let
  Source = each List.First(List.RemoveFirstN(_, each _ = null), null)
in
  Source
