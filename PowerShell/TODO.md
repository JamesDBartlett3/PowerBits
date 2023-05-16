# PowerBits PowerShell TODOs

## Technical

- SQL Server
- SQL Agent Job
- Dataflow w/ REST API
- Test Dataflow refresh on job completion
- Add "-Warningaction SilentlyContinue" to all Connect-*ServiceAccount" calls
- Debug "takeover dataset" workflow. This series of piped cmdlets crashed the terminal in both the Havens Consulting and Lagos PBIUG livestreams, so there's definitely something wrong.
- Re-word intro to, and refactor synopsis of Copy-PowerBIReportContentToBlankPBIXFile, explicitly calling out its usefulness for downloading thin reports connected to datasets in other workspaces
- Add Power BI security audit script
- Add output showing all datasets that were taken over

- Re-examine the "??" operator, which was recently replaced. Determine if that was actually necessary, and change back if possible. 
- Consider renaming Copy-PowerBIReportContentToBlankPBIXFile to something more semantic.
- Refactor Copy-PowerBIReportContentToBlankPBIXFile default behavior to output a report file named the same as the original report, save it in the temp folder, and then open that location, using the same method as in Export-PowerBIReportsFromWorkspaces. 
- Refactor all functions into scripts
- Handle thin reports & datasets

## Presentation
- PPT w/ offline demo & space in upper-right corner
- Backup VM
- Practice
