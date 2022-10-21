#Login-PowerBI

$activities = Get-PowerBIActivityEvent -StartDateTime "2020-11-01T00:00:00" -EndDateTime "2020-11-07T23:59:59" -ActivityType 'ViewDashboard' -User 'jbartlett@dmu.edu' | ConvertFrom-Json

$activities.Count
$activities[0]