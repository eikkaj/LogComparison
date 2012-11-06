Const ForReading = 1

Dim currentDate
Dim currentYear
Dim currentMonth
Dim currentDay

Dim yesterdayDay
Dim yesterdayMonth
Dim yesterdayYear
Dim yesterdayDate

Dim todayFile
Dim yesterdayFile

Dim file1,file2

currentYear = DatePart("yyyy", Now)
currentMonth = DatePart("m",Now)
currentDay = DatePart("d", Now)

'********build today's date and today's logfile name********'
'adds 0 to single digit months'
If currentMonth = 10 OR currentMonth = 11 OR currentMonth = 12 Then
		currentMonth = currentMonth
Else
	currentMonth = 0 & currentMonth
End If

if currentDay < 10 Then
	currentDay = 0 & currentDay
End If
'end single digit month'

currentDate = currentYear & currentMonth & currentDay
todayFile = currentDate & "-0000.txt"
'*********end today's date and today's logfile name*********'

'******Build Yesterday's Date and Yesterday's Filename******'
'Gets yesterday's day'
If currentDay = 1 Then
	If currentMonth = 2 Then
		yesterdayDay = 29
	ElseIf currentMonth = 4 Then
		yesterdayDay = 30
	ElseIf currentMonth = 6 Then
		yesterdayDay = 30
	ElseIf currenMonth = 9 Then
		yesterdayDay = 30
	ElseIf currentMonth = 11 Then
		yesterdayDay = 30
	Else
		yesterdayDay = 31
	End If
Else
	yesterdayDay = currentDay - 1
End If
'end yesterday's day'

'Check if it is the same month or not'
If currentDay = 1 Then
	yesterdayMonth = currentMonth - 1
Else
	yesterdayMonth = currentMonth
End If
'End Month Check'

'Check for same year or not'
If currentDay = 1 Then
	If currentMonth = 1 Then
		yesterdayYear = currentYear - 1
	Else
		yesterdayYear = currentYear
	End If
Else
	yesterdayYear = currentYear
End If
'End year check'

'Add zero's to single digits'
'adds 0 to single digit months'
If yesterdayMonth = 10 OR yesterdayMonth = 11 OR yesterdayMonth = 12 Then
		yesterdayMonth = yesterdayMonth
Else
	yesterdayMonth = 0 & yesterdayMonth
End If

if yesterdayDay < 10 Then
	yesterdayDay = 0 & yesterdayDay
End If
'end single digit month'
'End of zero's'

'Build Yesterday's Filename'
yesterdayDate = yesterdayYear & yesterdayMonth & yesterdayDay
yesterdayFile = yesterdayDate & "-0000.txt"
'End FileBuild'

'*******End Yesterday's Date build and Yesterday's FileName********'

file1 = "C:\Users\jaldama\Desktop\vb_test\" &todayFile
file2 = "C:\Users\jaldama\Desktop\vb_test\" &yesterdayFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile1 = objFSO.OpenTextFile(file2, ForReading)

strCurrentDevices = objFile1.ReadAll
objFile1.Close

Set objFile2 = objFSO.OpenTextFile(file1, ForReading)

Do Until objFile2.AtEndOfStream
    strAddress = objFile2.ReadLine
    If InStr(strCurrentDevices, strAddress) = 0 Then
        strNotCurrent = strNotCurrent & strAddress & vbCrLf
    End If
Loop

objFile2.Close

Wscript.Echo "Changes since last revision: " & vbCrLf & strNotCurrent & currentDate
Wscript.Echo "Yesterday's Date was: " & yesterdayDate
Wscript.Echo "Today's Filename is: " & todayFile

Set objFile3 = objFSO.CreateTextFile("C:\Users\jaldama\Desktop\vb_test\changes.txt")

objFile3.WriteLine "****************DETECTED CHANGES IN TODAY'S LOG FILE****************"
objFile3.WriteLine "Today's Date: "  & currentDate
objFile3.WriteLine strNotCurrent
objFile3.Close

If strNotCurrent = "" Then
	Wscript.Echo " Nothing New To Report!"
Else
	Dim SMTPServer
	SMTPServer = "WFDS-ExchMB.picis.com"
	Dim EmailPass
	EmailPass = ""
	Dim UserName
	UserName = ""
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = "Log File Checker"
	objMessage.From = ""
	objMessage.AddAttachment "\AddressTo\changes.txt"
	objMessage.To = ""
	objMessage.TextBody = "Results for log checker"
	
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
	
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objMessage.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    objMessage.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
    objMessage.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objMessage.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = UserName
    objMessage.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EmailPass
	objMessage.Configuration.Fields.Update
	objMessage.Send
End If