<#
.SYNOPSIS
    Replaces hard-coded connection literals (server/database, etc.) in one or more PBIP
    semantic models' partitions with shared Power Query parameters, using the real
    Tabular Object Model (TOM) rather than text-level TMDL editing.

.DESCRIPTION
    Loads <SemanticModel>\definition via TmdlSerializer.DeserializeDatabaseFromFolder
    (Microsoft.AnalysisServices.Tabular.dll) into an in-memory Database/Model object,
    finds M-backed partitions whose source expression calls a known "connector root"
    function with literal server/database strings, e.g.:

        Source = Sql.Database("myserver.database.windows.net", "MyDB")

    rewrites them to reference shared parameters:

        Source = Sql.Database(#"SqlEndpoint", #"Database")

    creates one NamedExpression per distinct literal value (collapsing duplicates across
    tables), files them under a "Parameters" QueryGroup, and serializes the whole database
    back to TMDL via TmdlSerializer.SerializeDatabaseToFolder. Because this goes through
    the real object model, the write-back is structurally validated (bad references,
    invalid metadata, and malformed TMDL all surface as exceptions here rather than as a
    silent bad edit) -- but note that SerializeDatabaseToFolder rewrites every document in
    the definition folder from the in-memory model, which can reformat files this script
    didn't otherwise need to touch. Use -Backup, and diff before committing.

    TOM assembly selection is driven entirely by whether it can round-trip a given model's
    compatibility level unchanged -- not by version number. A locally found copy (SSMS,
    Visual Studio/SSDT, Tabular Editor, or a prior run of this script) is tried first, in
    an isolated child process so a bad candidate can never get loaded into this session;
    if it can deserialize the model without altering its declared compatibility level,
    it's used as-is. Otherwise the latest stable "Microsoft.AnalysisServices" package is
    downloaded from nuget.org, cached locally, and verified the same way before use. The
    compatibility level is re-checked immediately before writing and (on a real, non
    -WhatIf run) re-read from disk afterward -- the script aborts that model rather than
    silently letting it drift.

.PARAMETER SemanticModelPath
    Path to a "<Name>.SemanticModel" folder (the one containing a "definition" subfolder).
    Accepts pipeline input, including by property name -- pipe Get-ChildItem output
    (matches on its FullName property) directly in to process a batch of models.

.PARAMETER ConnectorMap
    Hashtable of connector function name -> two-element array of parameter base names
    for [server, database]. Extend this if your model uses a connector not listed by
    default. Only two-argument (string, string [, options-record]) connector shapes are
    supported out of the box.

.PARAMETER NavigationKeysToParameterize
    Record-key names in a navigation step, e.g. Source{[Schema="dbo", Item="Orders"]},
    whose literal value should be extracted into a shared parameter when it's a plain
    string literal. Default: @('Schema') -- the schema is normally the same across every
    table and worth sharing; "Item" (the table/view name) is deliberately never touched
    since it's unique per table by definition.

.PARAMETER QueryGroupName
    Name of the query group new parameters are filed under. Default: "Parameters".
    Skipped automatically if a given model's compatibility level is below 1480
    (QueryGroup support), since parameters themselves only require compatibility level
    1400+.

.PARAMETER SkipQueryGroup
    Skip creating/using a query group even if the model supports one.

.PARAMETER Backup
    Copy each SemanticModel folder to a timestamped sibling folder before writing.
    Strongly recommended: SerializeDatabaseToFolder rewrites every TMDL document.

.PARAMETER TomAssemblyPath
    Path to a folder that already contains Microsoft.AnalysisServices*.dll (e.g. a
    Tabular Editor 3 install directory, or a Visual Studio SSAS extension folder).
    When supplied, this is tried first (still verified for compatibility) instead of
    searching the machine.

.PARAMETER TomPackageVersion
    Pin a specific "Microsoft.AnalysisServices" NuGet package version instead of the
    latest stable release, if a NuGet download is needed. Version doesn't otherwise
    matter to this script -- only whether the assembly can represent a model's
    compatibility level without changing it.

.PARAMETER TomCacheDirectory
    Where downloaded/extracted NuGet package content is cached between runs.
    Default: $env:LOCALAPPDATA\PbipTom\packages

.EXAMPLE
    .\Convert-PbipConnectionsToParameters.ps1 -SemanticModelPath 'C:\repo\Sales.SemanticModel' -WhatIf

    Preview what would change for one model.

.EXAMPLE
    Get-ChildItem 'C:\repo' -Directory -Filter '*.SemanticModel' |
        .\Convert-PbipConnectionsToParameters.ps1 -Backup

    Process every semantic model under a repo in one batch, backing each up first. Piped
    models are processed sequentially within this one process, reusing the same
    TOM assembly across all of them (re-verified compatible with each model in turn) --
    a single process can only ever have one version of Microsoft.AnalysisServices.Tabular
    loaded at a time, so a model needing a genuinely different TOM version than the one
    already loaded is skipped with an error rather than silently mishandled.

.EXAMPLE
    Get-ChildItem 'C:\repo' -Directory -Filter '*.SemanticModel' | ForEach-Object -Parallel {
        & "C:\scripts\Convert-PbipConnectionsToParameters.ps1" -SemanticModelPath $_.FullName -Backup -Confirm:$false
    } -ThrottleLimit 4

    For genuine parallelism (e.g. many independent models, some possibly needing different
    TOM versions), invoke pwsh as a real child process per item instead -- ForEach-Object
    -Parallel runspaces still share this process's assembly-load context, so they cannot
    each load a different TOM version; separate processes can. Swap the scriptblock body
    for `Start-Process pwsh -ArgumentList '-NoProfile','-File',...,'-SemanticModelPath',$_.FullName -Wait`
    if that matters for your batch.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('FullName')]
    [string]$SemanticModelPath,

    [hashtable]$ConnectorMap = @{
        'Sql.Database'               = @('SqlEndpoint', 'Database')
        'Sql.Databases'              = @('SqlEndpoint', 'Database')
        'AnalysisServices.Database'  = @('AsServer', 'AsDatabase')
        'AnalysisServices.Databases' = @('AsServer', 'AsDatabase')
        'PostgreSQL.Database'        = @('PostgresServer', 'PostgresDatabase')
        'MySQL.Database'             = @('MySqlServer', 'MySqlDatabase')
    },

    [string[]]$NavigationKeysToParameterize = @('Schema'),

    [string]$QueryGroupName = 'Parameters',
    [switch]$SkipQueryGroup,
    [switch]$Backup,

    [string]$TomAssemblyPath,
    [string]$TomPackageVersion,
    [string]$TomCacheDirectory = (Join-Path $env:LOCALAPPDATA 'PbipTom\packages')
)

begin {
    $ErrorActionPreference = 'Stop'

    # This script's -WhatIf must only gate the two operations that actually change a model
    # (the backup copy and the definition-folder swap, both behind explicit
    # $PSCmdlet.ShouldProcess() checks below). Everything else -- writing/deleting the
    # isolated-probe temp script, creating the NuGet cache folder, cleaning up the staging
    # folder -- is internal plumbing that must run for real even under -WhatIf, or the
    # compatibility check itself breaks (Set-Content/New-Item/Remove-Item all silently
    # no-op under an inherited $WhatIfPreference, which is exactly what caused the "not
    # compatible" false negatives and the missing-nupkg-path failure). Reassigning the
    # preference *variable* here does not affect $PSCmdlet.ShouldProcess()'s own answer --
    # that stays tied to whether -WhatIf was actually bound to this script's invocation --
    # verified empirically before relying on it here.
    $WhatIfPreference = $false

    # PowerShell's $PWD and the process's actual .NET current directory can silently
    # diverge after Set-Location; a relative path handed straight to a .NET API resolves
    # against the (possibly stale) latter, not the former. Join against $PWD explicitly
    # rather than relying on Resolve-Path, since this needs to work for paths that don't
    # exist yet (e.g. a cache directory this script is about to create).
    function Resolve-ToAbsolutePath {
        param([string]$Path)
        if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
        return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
    }

    $TomCacheDirectory = Resolve-ToAbsolutePath $TomCacheDirectory

    # =====================================================================================
    # TOM assembly acquisition
    # =====================================================================================

    function Select-BestLibFolder {
        param([string]$LibRoot)
        $preference = if ($PSVersionTable.PSEdition -eq 'Core') {
            @('net9.0', 'net8.0', 'net7.0', 'net6.0', 'netstandard2.1', 'netstandard2.0')
        } else {
            @('net481', 'net48', 'net472', 'net462', 'net452', 'net45')
        }
        $available = Get-ChildItem -Path $LibRoot -Directory -ErrorAction SilentlyContinue
        foreach ($tfm in $preference) {
            $match = $available | Where-Object { $_.Name -eq $tfm }
            if ($match) { return $match.FullName }
        }
        if ($available) { return $available[0].FullName }
        throw "No target-framework subfolders found under $LibRoot"
    }

    function Find-ExistingTomFolder {
        $candidateRoots = @(
            "$env:ProgramFiles\Microsoft SQL Server Management Studio*",
            "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio*",
            "$env:ProgramFiles\Microsoft Visual Studio",
            "${env:ProgramFiles(x86)}\Microsoft Visual Studio",
            "$env:ProgramFiles\Tabular Editor*",
            "${env:ProgramFiles(x86)}\Tabular Editor*",
            "$env:LOCALAPPDATA\TabularEditor*"
        )

        foreach ($cmdName in 'te', 'TabularEditor', 'TabularEditor3') {
            $cmd = Get-Command $cmdName -ErrorAction SilentlyContinue
            if ($cmd) {
                $dir = Split-Path $cmd.Source -Parent
                $hit = Get-ChildItem -Path $dir -Filter 'Microsoft.AnalysisServices.Tabular.dll' -ErrorAction SilentlyContinue
                if ($hit) { return $hit[0].DirectoryName }
            }
        }

        foreach ($root in $candidateRoots) {
            $resolved = Resolve-Path $root -ErrorAction SilentlyContinue
            foreach ($r in $resolved) {
                $hit = Get-ChildItem -Path $r.Path -Filter 'Microsoft.AnalysisServices.Tabular.dll' -Recurse -Depth 6 -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                if ($hit) { return $hit.DirectoryName }
            }
        }

        if (Test-Path $TomCacheDirectory) {
            $hit = Get-ChildItem -Path $TomCacheDirectory -Filter 'Microsoft.AnalysisServices.Tabular.dll' -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | Select-Object -First 1
            if ($hit) { return $hit.DirectoryName }
        }

        return $null
    }

    function Install-TomFromNuget {
        param([string]$Version, [string]$CacheDirectory)

        $packageId = 'Microsoft.AnalysisServices'
        $idLower = $packageId.ToLowerInvariant()

        if (-not $Version) {
            Write-Host "Looking up latest stable '$packageId' version on nuget.org..." -ForegroundColor Cyan
            $index = Invoke-RestMethod -Uri "https://api.nuget.org/v3-flatcontainer/$idLower/index.json"
            # NuGet's flat-container version list can include prereleases (e.g. "19.115.0-preview1"),
            # which sort after their stable predecessor as plain strings -- exclude anything with a
            # semver prerelease tag (a hyphen) so a plain "last item" pick can't land on one.
            $Version = $index.versions | Where-Object { $_ -notmatch '-' } | Select-Object -Last 1
        }

        $pkgRoot = Join-Path $CacheDirectory "$packageId.$Version"
        $libRoot = Join-Path $pkgRoot 'lib'

        if (-not (Test-Path $libRoot)) {
            New-Item -ItemType Directory -Path $CacheDirectory -Force | Out-Null
            $nupkgPath = Join-Path $CacheDirectory "$idLower.$Version.nupkg"
            $url = "https://api.nuget.org/v3-flatcontainer/$idLower/$Version/$idLower.$Version.nupkg"
            Write-Host "Downloading $packageId $Version from nuget.org..." -ForegroundColor Cyan
            Invoke-WebRequest -Uri $url -OutFile $nupkgPath -UseBasicParsing
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($nupkgPath, $pkgRoot)
            Remove-Item $nupkgPath -Force
            Write-Host "Installed to $pkgRoot" -ForegroundColor Green
        } else {
            Write-Host "Using cached $packageId $Version at $pkgRoot" -ForegroundColor DarkGray
        }

        return Select-BestLibFolder -LibRoot $libRoot
    }

    function Import-TomAssemblies {
        param([string]$LibFolder)

        # Resolve any assembly TOM lazily loads (e.g. Tabular.Json.dll, needed for some
        # TMDL constructs but not eagerly Add-Type'd here) from the same folder. This is not
        # just defensive: without .GetNewClosure(), the scriptblock loses its binding to
        # $LibFolder once this function returns -- the Resolving event fires later, from
        # inside DeserializeDatabaseFromFolder, well after Import-TomAssemblies's own call
        # frame is gone, so the closure must be captured explicitly here.
        $resolverBody = {
            param($context, $assemblyName)
            $candidate = Join-Path $LibFolder "$($assemblyName.Name).dll"
            if (Test-Path $candidate) { return $context.LoadFromAssemblyPath($candidate) }
            return $null
        }.GetNewClosure()
        $resolver = [System.Func[System.Runtime.Loader.AssemblyLoadContext, System.Reflection.AssemblyName, System.Reflection.Assembly]] $resolverBody
        if ($PSVersionTable.PSEdition -eq 'Core') {
            [System.Runtime.Loader.AssemblyLoadContext]::Default.add_Resolving($resolver)
        } else {
            Register-ObjectEvent -InputObject ([System.AppDomain]::CurrentDomain) -EventName AssemblyResolve -Action {
                $name = ([Reflection.AssemblyName]$Event.SourceArgs[1].Name).Name
                $candidate = Join-Path $using:LibFolder "$name.dll"
                if (Test-Path $candidate) { [Reflection.Assembly]::LoadFrom($candidate) }
            } | Out-Null
        }

        foreach ($dllName in 'Microsoft.AnalysisServices.Core.dll', 'Microsoft.AnalysisServices.dll', 'Microsoft.AnalysisServices.Tabular.dll') {
            $path = Join-Path $LibFolder $dllName
            if (-not (Test-Path $path)) { throw "Expected assembly not found: $path" }
            Add-Type -Path $path
        }
    }

    function Get-DeclaredCompatibilityLevel {
        param([string]$DefinitionPath)
        $dbFile = Join-Path $DefinitionPath 'database.tmdl'
        if (-not (Test-Path $dbFile)) { throw "database.tmdl not found under $DefinitionPath" }
        $m = [regex]::Match((Get-Content $dbFile -Raw), '(?m)^\s*compatibilityLevel:\s*(\d+)')
        if (-not $m.Success) { throw "Could not find a compatibilityLevel property in $dbFile" }
        return [int]$m.Groups[1].Value
    }

    function Test-TomCandidate {
        # Runs in an isolated child process: loading a *wrong* Microsoft.AnalysisServices.Tabular.dll
        # into this session would be permanent (you cannot unload/swap a loaded assembly), so
        # candidates are proven compatible from the outside before this process ever touches them.
        param([string]$LibFolder, [string]$DefinitionPath, [int]$ExpectedCompatLevel)

        $probePath = Join-Path ([System.IO.Path]::GetTempPath()) "tom-probe-$([guid]::NewGuid().ToString('N')).ps1"
        $probeBody = @'
param([string]$LibFolder, [string]$DefinitionPath, [int]$ExpectedCompatLevel)
try {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $resolver = [System.Func[System.Runtime.Loader.AssemblyLoadContext, System.Reflection.AssemblyName, System.Reflection.Assembly]] {
            param($context, $assemblyName)
            $candidate = Join-Path $LibFolder "$($assemblyName.Name).dll"
            if (Test-Path $candidate) { return $context.LoadFromAssemblyPath($candidate) }
            return $null
        }
        [System.Runtime.Loader.AssemblyLoadContext]::Default.add_Resolving($resolver)
    }
    Add-Type -Path (Join-Path $LibFolder 'Microsoft.AnalysisServices.Core.dll')
    Add-Type -Path (Join-Path $LibFolder 'Microsoft.AnalysisServices.dll')
    Add-Type -Path (Join-Path $LibFolder 'Microsoft.AnalysisServices.Tabular.dll')
    $db = [Microsoft.AnalysisServices.Tabular.TmdlSerializer]::DeserializeDatabaseFromFolder($DefinitionPath)
    if ($db.CompatibilityLevel -ne $ExpectedCompatLevel) { exit 1 }
    exit 0
} catch {
    exit 1
}
'@
        Set-Content -Path $probePath -Value $probeBody -NoNewline
        try {
            $exe = (Get-Process -Id $PID).Path
            & $exe -NoProfile -NonInteractive -File $probePath -LibFolder $LibFolder -DefinitionPath $DefinitionPath -ExpectedCompatLevel $ExpectedCompatLevel *> $null
            return ($LASTEXITCODE -eq 0)
        } finally {
            Remove-Item $probePath -Force -ErrorAction SilentlyContinue
        }
    }

    function New-ConnectorRegex {
        param([string]$FunctionName, [int]$ArgCount)
        $escaped = [regex]::Escape($FunctionName)
        if ($ArgCount -ne 2) { throw "New-ConnectorRegex: only 2-argument connector shapes are supported (function '$FunctionName' requested $ArgCount)." }
        # Each argument slot is either a quoted string literal (captured into lit<N> for
        # possible parameterization) or a bare/quoted M identifier referencing another
        # shared query (captured into ident<N> and left completely untouched) -- e.g. a
        # Fabric Lakehouse SQL endpoint call is commonly
        # Sql.Database("server...", lh_structure), where the second argument already
        # references a shared query rather than being a literal.
        $argPattern = '(?:"(?<lit{0}>(?:[^"\\]|\\.)*)"|(?<ident{0}>#"[^"]*"|[A-Za-z_][A-Za-z0-9_]*))'
        $arg0 = $argPattern -f 0
        $arg1 = $argPattern -f 1
        return [regex]"\b$escaped\s*\(\s*$arg0\s*,\s*$arg1\s*(?:,\s*\[[^\]]*\])?\)"
    }

    function New-NavigationKeyRegex {
        # Matches a single "Key=value" fragment inside a navigation record, e.g. the
        # Schema= part of Source{[Schema="dbo", Item="Orders"]} -- deliberately scoped to
        # just that one key=value pair so replacing it never touches sibling keys (like
        # Item, which is unique per table and must never be parameterized).
        param([string]$KeyName)
        $escaped = [regex]::Escape($KeyName)
        return [regex]"\b$escaped\s*=\s*(?:""(?<lit>(?:[^""\\]|\\.)*)""|(?<ident>#""[^""]*""|[A-Za-z_][A-Za-z0-9_]*))"
    }

    function Test-IsParameterExpressionText {
        param([string]$ExpressionText)
        return [regex]::IsMatch($ExpressionText, '^"(?:[^"\\]|\\.)*"\s*meta\s*\[IsParameterQuery=true')
    }

    # Resolved once per invocation of this script and reused across every piped model --
    # a process can only ever have one Microsoft.AnalysisServices.Tabular.dll loaded.
    $script:tomLibFolder = $null
    $script:processedCount = 0
    $script:skippedCount = 0
    $script:failedCount = 0
}

process {
    $modelLabel = $SemanticModelPath
    try {
        if (-not (Test-Path (Join-Path $SemanticModelPath 'definition'))) {
            Write-Error "Skipping '$modelLabel': no 'definition' subfolder found -- not a PBIP semantic model folder."
            $script:skippedCount++
            return
        }

        # Resolve to an absolute path via PowerShell's own resolver before anything reaches
        # a raw .NET API (TmdlSerializer, File/Directory calls). PowerShell's $PWD and the
        # process's actual [System.IO.Directory]::GetCurrentDirectory() can silently diverge
        # after Set-Location -- Test-Path/Resolve-Path honor $PWD correctly, but a relative
        # path handed to a .NET method resolves against the (possibly stale) process
        # directory instead, and fails with a confusing "does not exist" even though the
        # folder is right there.
        $SemanticModelPath = (Resolve-Path -LiteralPath $SemanticModelPath).ProviderPath
        $definitionPath = Join-Path $SemanticModelPath 'definition'
        $expectedCompatLevel = Get-DeclaredCompatibilityLevel -DefinitionPath $definitionPath
        Write-Host "`n=== $modelLabel (compatibility level $expectedCompatLevel) ===" -ForegroundColor Cyan

        if (-not $script:tomLibFolder) {
            # First model in this batch: find or fetch a TOM assembly and prove it can
            # represent this model's compatibility level before loading it for real.
            $candidate = if ($TomAssemblyPath) {
                # Same relative-path hazard as $SemanticModelPath: this folder gets handed to
                # Add-Type and to a child process, both of which resolve relative paths
                # against the process's actual current directory, not PowerShell's $PWD.
                $resolvedTomAssemblyPath = (Resolve-Path -LiteralPath $TomAssemblyPath).ProviderPath
                if ((Get-Item $resolvedTomAssemblyPath).PSIsContainer) { $resolvedTomAssemblyPath } else { Split-Path $resolvedTomAssemblyPath -Parent }
            } else {
                Find-ExistingTomFolder
            }

            if ($candidate) {
                Write-Host "Checking whether the TOM assembly at $candidate can represent this model without changing its compatibility level..." -ForegroundColor DarkGray
                if (Test-TomCandidate -LibFolder $candidate -DefinitionPath $definitionPath -ExpectedCompatLevel $expectedCompatLevel) {
                    Write-Host "Compatible -- using it." -ForegroundColor DarkGray
                    $script:tomLibFolder = $candidate
                } else {
                    Write-Host "Not compatible (or failed to load) -- falling back to nuget.org." -ForegroundColor Yellow
                }
            }

            if (-not $script:tomLibFolder) {
                $script:tomLibFolder = Install-TomFromNuget -Version $TomPackageVersion -CacheDirectory $TomCacheDirectory
                if (-not (Test-TomCandidate -LibFolder $script:tomLibFolder -DefinitionPath $definitionPath -ExpectedCompatLevel $expectedCompatLevel)) {
                    throw "The downloaded Microsoft.AnalysisServices package still cannot represent this model's compatibility level ($expectedCompatLevel) without changing it. This model may use a compatibility level newer than any currently published package supports."
                }
            }

            Import-TomAssemblies -LibFolder $script:tomLibFolder
        } else {
            # Later model in the same batch: the assembly is already loaded and can't be
            # swapped, so just confirm it also handles this model before reusing it.
            if (-not (Test-TomCandidate -LibFolder $script:tomLibFolder -DefinitionPath $definitionPath -ExpectedCompatLevel $expectedCompatLevel)) {
                Write-Error "Skipping '$modelLabel': it needs a different TOM assembly than the one already loaded for this batch (from an earlier model). Process it in a separate invocation of this script."
                $script:skippedCount++
                return
            }
        }

        Write-Host "Loading model via TOM..." -ForegroundColor Cyan
        $database = [Microsoft.AnalysisServices.Tabular.TmdlSerializer]::DeserializeDatabaseFromFolder($definitionPath)
        $model = $database.Model
        # Real Power BI Desktop-authored database.tmdl files often declare an unnamed
        # database block (no name after the "database" keyword) -- Database.Name is blank
        # in that case, so fall back to the model path for a readable log line.
        $displayName = if ($database.Name) { $database.Name } else { $modelLabel }
        Write-Host "Loaded '$displayName', $($model.Tables.Count) table(s)." -ForegroundColor DarkGray

        $supportsExpressions = $database.CompatibilityLevel -ge 1400
        $localSkipQueryGroup = $SkipQueryGroup.IsPresent
        $supportsQueryGroups = $database.CompatibilityLevel -ge 1480
        if (-not $supportsExpressions) {
            throw "Model compatibility level $($database.CompatibilityLevel) is below 1400 -- shared M parameters are not supported."
        }
        if (-not $supportsQueryGroups -and -not $localSkipQueryGroup) {
            Write-Warning "Compatibility level $($database.CompatibilityLevel) is below 1480 -- skipping query group (needs 1480+)."
            $localSkipQueryGroup = $true
        }

        # ---- Existing parameters: reuse instead of duplicating ----
        $existingParams = @{}   # value -> parameter name
        foreach ($expr in $model.Expressions) {
            if ($expr.Kind -ne [Microsoft.AnalysisServices.Tabular.ExpressionKind]::M) { continue }
            if (-not (Test-IsParameterExpressionText $expr.Expression)) { continue }
            $m = [regex]::Match($expr.Expression, '^"(?<val>(?:[^"\\]|\\.)*)"\s*meta\s*\[IsParameterQuery=true')
            if ($m.Success) { $existingParams[$m.Groups['val'].Value] = $expr.Name }
        }

        # ---- Build the set of things with an .Expression to scan: every M-backed table
        # partition, PLUS every staging/non-loaded query -- those are NOT partitions at
        # all, they're plain NamedExpression objects in Model.Expressions (what
        # expressions.tmdl holds) since they never got "Enable Load" turned on. Skip the
        # ones that are actually parameters (handled above as reuse targets, not rewrite
        # targets -- a parameter's own value text isn't a query to scan).
        $expressionTargets = New-Object System.Collections.Generic.List[object]

        foreach ($table in $model.Tables) {
            foreach ($partition in $table.Partitions) {
                if ($partition.Source -isnot [Microsoft.AnalysisServices.Tabular.MPartitionSource]) { continue }
                $expressionTargets.Add([pscustomobject]@{
                    Location = $table.Name
                    Label    = "partition '$($partition.Name)' on table '$($table.Name)'"
                    Target   = $partition.Source
                })
            }
        }
        foreach ($expr in $model.Expressions) {
            if ($expr.Kind -ne [Microsoft.AnalysisServices.Tabular.ExpressionKind]::M) { continue }
            if (Test-IsParameterExpressionText $expr.Expression) { continue }
            $expressionTargets.Add([pscustomobject]@{
                Location = "[staging] $($expr.Name)"
                Label    = "staging query '$($expr.Name)'"
                Target   = $expr
            })
        }

        # ---- Pass 1: discover every literal connection value and navigation schema value ----
        $paramAssignments = @{}   # "<Function>|<RoleIndex>|<Value>" or "NavKey:<Key>|<Value>" -> parameter name
        $usedNames = New-Object System.Collections.Generic.HashSet[string]
        $existingParams.Values | ForEach-Object { [void]$usedNames.Add($_) }

        $discoveries = New-Object System.Collections.Generic.List[object]

        function Resolve-ParamAssignment {
            param([string]$Key, [string]$Value, [string]$BaseName)
            if ($paramAssignments.ContainsKey($Key)) { return }
            if ($existingParams.ContainsKey($Value)) {
                $paramAssignments[$Key] = $existingParams[$Value]
            } else {
                $candidate = $BaseName
                $suffix = 1
                while ($usedNames.Contains($candidate)) { $suffix++; $candidate = "$BaseName$suffix" }
                [void]$usedNames.Add($candidate)
                $paramAssignments[$Key] = $candidate
            }
        }

        foreach ($entry in $expressionTargets) {
            $text = $entry.Target.Expression

            foreach ($funcName in $ConnectorMap.Keys) {
                $roles = $ConnectorMap[$funcName]
                $regex = New-ConnectorRegex -FunctionName $funcName -ArgCount $roles.Count
                foreach ($m in $regex.Matches($text)) {
                    for ($i = 0; $i -lt $roles.Count; $i++) {
                        $litGroup = $m.Groups["lit$i"]
                        if (-not $litGroup.Success) { continue }   # already a reference to another query -- leave it alone
                        $value = $litGroup.Value
                        if ([string]::IsNullOrEmpty($value)) { continue }
                        $key = "$funcName|$i|$value"
                        Resolve-ParamAssignment -Key $key -Value $value -BaseName $roles[$i]
                        $discoveries.Add([pscustomobject]@{
                            Table     = $entry.Location
                            Function  = $funcName
                            Role      = $roles[$i]
                            Value     = $value
                            Parameter = $paramAssignments[$key]
                            IsNew     = -not $existingParams.ContainsKey($value)
                        })
                    }
                }
            }

            foreach ($keyName in $NavigationKeysToParameterize) {
                $navRegex = New-NavigationKeyRegex -KeyName $keyName
                foreach ($m in $navRegex.Matches($text)) {
                    $litGroup = $m.Groups['lit']
                    if (-not $litGroup.Success) { continue }   # already a reference -- leave it alone
                    $value = $litGroup.Value
                    if ([string]::IsNullOrEmpty($value)) { continue }
                    $key = "NavKey:$keyName|$value"
                    Resolve-ParamAssignment -Key $key -Value $value -BaseName $keyName
                    $discoveries.Add([pscustomobject]@{
                        Table     = $entry.Location
                        Function  = 'Navigation'
                        Role      = $keyName
                        Value     = $value
                        Parameter = $paramAssignments[$key]
                        IsNew     = -not $existingParams.ContainsKey($value)
                    })
                }
            }
        }

        if ($discoveries.Count -eq 0) {
            Write-Host "No recognized hard-coded connection literals found in any M partition or staging query." -ForegroundColor Yellow
            $script:processedCount++
            return
        }

        Write-Host "Discovered connection literals:" -ForegroundColor Cyan
        $discoveries | Sort-Object Table, Function, Role | Format-Table Table, Function, Role, Value, Parameter, IsNew -AutoSize | Out-Host

        # ---- Pass 2: rewrite every M partition's and staging query's expression in-memory ----
        foreach ($entry in $expressionTargets) {
            $original = $entry.Target.Expression
            $updated = $original

            foreach ($funcName in $ConnectorMap.Keys) {
                $roles = $ConnectorMap[$funcName]
                $regex = New-ConnectorRegex -FunctionName $funcName -ArgCount $roles.Count
                $evaluator = {
                    param($m)
                    $refs = for ($i = 0; $i -lt $roles.Count; $i++) {
                        $litGroup = $m.Groups["lit$i"]
                        if ($litGroup.Success) {
                            $value = $litGroup.Value
                            $key = "$funcName|$i|$value"
                            if ($paramAssignments.ContainsKey($key)) { '#"{0}"' -f $paramAssignments[$key] } else { '"{0}"' -f $value }
                        } else {
                            $m.Groups["ident$i"].Value   # already a reference to another query -- reproduce it verbatim
                        }
                    }
                    "$funcName($($refs -join ', '))"
                }
                $updated = [regex]::Replace($updated, $regex, [System.Text.RegularExpressions.MatchEvaluator]$evaluator)
            }

            foreach ($keyName in $NavigationKeysToParameterize) {
                $navRegex = New-NavigationKeyRegex -KeyName $keyName
                $navEvaluator = {
                    param($m)
                    $litGroup = $m.Groups['lit']
                    if ($litGroup.Success) {
                        $value = $litGroup.Value
                        $key = "NavKey:$keyName|$value"
                        $ref = if ($paramAssignments.ContainsKey($key)) { '#"{0}"' -f $paramAssignments[$key] } else { '"{0}"' -f $value }
                        "$keyName=$ref"
                    } else {
                        "$keyName=$($m.Groups['ident'].Value)"   # already a reference -- reproduce it verbatim
                    }
                }
                $updated = [regex]::Replace($updated, $navRegex, [System.Text.RegularExpressions.MatchEvaluator]$navEvaluator)
            }

            if ($updated -ne $original) {
                $entry.Target.Expression = $updated
                Write-Host "Rewrote $($entry.Label)" -ForegroundColor Green
            }
        }

        # ---- Create NamedExpression objects for newly discovered parameters ----
        $newParams = $discoveries | Where-Object { $_.IsNew } | Sort-Object Parameter -Unique
        if ($newParams) {
            $queryGroup = $null
            if (-not $localSkipQueryGroup) {
                $queryGroup = $model.QueryGroups | Where-Object { $_.Name -eq $QueryGroupName } | Select-Object -First 1
                if (-not $queryGroup) {
                    $queryGroup = New-Object Microsoft.AnalysisServices.Tabular.QueryGroup
                    $queryGroup.Folder = $QueryGroupName
                    $model.QueryGroups.Add($queryGroup)
                }
            }

            foreach ($p in $newParams) {
                $expr = New-Object Microsoft.AnalysisServices.Tabular.NamedExpression
                $expr.Name = $p.Parameter
                $expr.Kind = [Microsoft.AnalysisServices.Tabular.ExpressionKind]::M
                $expr.Expression = '"{0}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]' -f $p.Value
                $expr.LineageTag = [guid]::NewGuid().ToString()
                if ($queryGroup) { $expr.QueryGroup = $queryGroup }
                $model.Expressions.Add($expr)
            }
            Write-Host "Created $($newParams.Count) new parameter(s): $($newParams.Parameter -join ', ')" -ForegroundColor Green
        } else {
            Write-Host "No new parameters needed -- all matched literals reused existing parameters." -ForegroundColor Yellow
        }

        # =================================================================================
        # Write back
        # =================================================================================

        if ($Backup) {
            $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            $backupPath = "$SemanticModelPath.bak-$stamp"
            if ($PSCmdlet.ShouldProcess($backupPath, "Create backup copy of $SemanticModelPath")) {
                Copy-Item -Path $SemanticModelPath -Destination $backupPath -Recurse
                Write-Host "Backed up to $backupPath" -ForegroundColor Green
            }
        }

        if ($database.CompatibilityLevel -ne $expectedCompatLevel) {
            throw "Internal error: in-memory compatibility level ($($database.CompatibilityLevel)) no longer matches the model's original declared level ($expectedCompatLevel) -- aborting without writing."
        }

        # Serialize to a scratch folder and verify it there FIRST. The live 'definition'
        # folder is never touched by anything that hasn't already been proven good --
        # a failed verification here leaves the original completely untouched.
        $stagingPath = Join-Path ([System.IO.Path]::GetTempPath()) "tmdl-write-$([guid]::NewGuid().ToString('N'))"
        try {
            [Microsoft.AnalysisServices.Tabular.TmdlSerializer]::SerializeDatabaseToFolder($database, $stagingPath)
            $writtenLevel = Get-DeclaredCompatibilityLevel -DefinitionPath $stagingPath
            if ($writtenLevel -ne $expectedCompatLevel) {
                throw "Serializing would change the compatibility level (was $expectedCompatLevel, would become $writtenLevel). Nothing was written to $definitionPath."
            }

            if ($PSCmdlet.ShouldProcess($definitionPath, 'Replace definition folder with newly-serialized TMDL (verified compatibility level unchanged)')) {
                # Swap by renaming the original aside first rather than deleting it outright,
                # so a failure partway through the swap still leaves a complete copy on disk
                # to roll back to -- the live folder is never in a half-written state.
                $displacedPath = "$definitionPath.displaced-$([guid]::NewGuid().ToString('N'))"
                Rename-Item -Path $definitionPath -NewName (Split-Path $displacedPath -Leaf)
                try {
                    Move-Item -Path $stagingPath -Destination $definitionPath
                    Remove-Item -Path $displacedPath -Recurse -Force
                } catch {
                    if (Test-Path $definitionPath) { Remove-Item -Path $definitionPath -Recurse -Force -ErrorAction SilentlyContinue }
                    Rename-Item -Path $displacedPath -NewName (Split-Path $definitionPath -Leaf)
                    throw "Swap failed and was rolled back to the original: $($_.Exception.Message)"
                }
                Write-Host "Wrote changes (compatibility level unchanged: $writtenLevel)" -ForegroundColor Green
            }
        } finally {
            if (Test-Path $stagingPath) { Remove-Item -Path $stagingPath -Recurse -Force -ErrorAction SilentlyContinue }
        }

        $script:processedCount++
    } catch {
        Write-Error "Failed on '$modelLabel': $($_.Exception.Message)"
        $script:failedCount++
    }
}

end {
    Write-Host "`nBatch complete: $script:processedCount processed, $script:skippedCount skipped, $script:failedCount failed." -ForegroundColor Cyan
    Write-Host "Open each changed project in Power BI Desktop to confirm it loads cleanly." -ForegroundColor Cyan
}
