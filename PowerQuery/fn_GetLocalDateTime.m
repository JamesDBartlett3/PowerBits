(
  DaylightSavingTime_Month_Number as number,  // Number of the month when DST starts
  DaylightSavingTime_Sunday_Number as number, // Nth Sunday of the month when DST starts
  StandardTime_Month_Number as number,        // Number of the month when DST ends
  StandardTime_Sunday_Number as number,       // Nth Sunday of the month when DST ends
  DaylightSavingTime_GMT_Offset as number,    // Hours offset from GMT during Daylight Saving Time
  StandardTime_GMT_Offset as number,          // Hours offset from GMT during Standard Time
  Include_LocalTimeZone as logical            // Whether to include the local timezone in the result
) =>
let
  UTC_DateTimeZone = DateTimeZone.UtcNow(),
  UTC_Date = Date.From( UTC_DateTimeZone ),
  Start_DaylightSavingTime = Date.StartOfWeek( #date( Date.Year( UTC_Date ) , DaylightSavingTime_Month_Number , DaylightSavingTime_Sunday_Number * 7 ), Day.Sunday ),
  Start_StandardTime = Date.StartOfWeek( #date( Date.Year( UTC_Date ) , StandardTime_Month_Number, StandardTime_Sunday_Number * 7 ), Day.Sunday ),
  UTC_Offset = if UTC_Date >= Start_DaylightSavingTime and UTC_Date < Start_StandardTime then DaylightSavingTime_GMT_Offset else StandardTime_GMT_Offset,
  Local_TimeZone = DateTimeZone.SwitchZone( UTC_DateTimeZone, UTC_Offset),
  Result = if Include_LocalTimeZone then Local_TimeZone else DateTimeZone.RemoveZone( Local_TimeZone )
in
  Result
