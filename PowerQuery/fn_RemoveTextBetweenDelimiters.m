(
  stringToClean as text, 
  startDelimiter as text, 
  endDelimiter as text, 
  optional trimWhitespace as logical
) => [
  trim = if List.Contains({false, null, ""}, trimWhitespace) then false else true,
  fnRemoveFirstTag = (DELIM as text) =>
    let
      OpeningTag = Text.PositionOf(DELIM, startDelimiter),
      ClosingTag = Text.PositionOf(DELIM, endDelimiter),
      Output = 
        if OpeningTag = -1 
        then DELIM 
        else Text.RemoveRange(DELIM, OpeningTag, ClosingTag - OpeningTag + 1)
    in
      Output,
  fnRemoveDELIM = (y as text) =>
    if fnRemoveFirstTag(y) = y
    then y 
    else @fnRemoveDELIM(fnRemoveFirstTag(y)),
  Output = if trim then Text.Trim(@fnRemoveDELIM(stringToClean)) else @fnRemoveDELIM(stringToClean)
][Output]