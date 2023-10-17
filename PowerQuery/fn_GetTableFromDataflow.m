(WorkspaceNameOrId as text, DataflowNameOrId as text, TableName as text, RowLimit as number) => 
let
  Source = PowerPlatform.Dataflows(null),
  WorkspaceList = Source{[Id="Workspaces"]}[Data],
  SelectedWorkspace = 
    try WorkspaceList{[workspaceName=WorkspaceNameOrId]}[Data]
    otherwise WorkspaceList{[workspaceId=WorkspaceNameOrId]}[Data],
  SelectedDataflow = 
    try SelectedWorkspace{[dataflowName=DataflowNameOrId]}[Data]
    otherwise SelectedWorkspace{[dataflowId=DataflowNameOrId]}[Data],
  SelectedTable = SelectedDataflow{[entity=TableName,version=""]}[Data],
  FilterLogic = if RowLimit < 0
    then SelectedTable
    else Table.FirstN(SelectedTable, RowLimit)
in
  FilterLogic