Dim currentDate
Dim currentYear
Dim currentMonth
Dim currentDay

Dim todayFile
Dim FSO

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
todayFile = "C:\AddressToFile\" &currentDate & "-0000.txt"

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile todayFile, "C:\Address\To\new.txt",[true]