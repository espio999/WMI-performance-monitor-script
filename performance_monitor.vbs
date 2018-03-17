option explicit

'dim start_date
dim dead_line
dim log_folder
dim server_list
dim server_array()
dim log_array()
dim now_year
dim now_month
dim now_day
dim perf_log_name
dim my_fso
dim my_file
dim my_flg

on error resume next

'start_date = now
dead_line = cDate("23:55:00")
log_folder = "c:\performance_monitor\log\"
server_list = "server_list.txt"
now_year = cStr(Year(now))
now_month = attachZero(cStr(Month(now)), 2)
now_day = attachZero(cStr(Day(now)), 2)
perf_log_name = now_year & now_month & now_day & ".csv"
'perf_log_name = cstr(now_year) & cstr(now_month) & cstr(now_day) & ".csv"
my_flg = true

set my_fso = CreateObject("Scripting.FileSystemObject")

CreateServerArray my_fso, server_array, log_folder & server_list
CreateLogArray server_array, log_array, log_folder, perf_log_name
CreatePerfLogLoop log_array, my_fso, my_file

while my_flg = true
	CreateCSVLoop my_flg, my_fso, my_file, log_array, server_array, dead_line
'	CreateCSVLoop my_flg, my_fso, my_file, log_array, server_array, start_date
wend

closer my_fso


Function attachZero(val, length)
	dim i
	dim ret
	
	i = len(val)
	ret = ""
	
	While(i < length)
		ret = ret & "0"
		i = i + 1
	WEnd
	
	attachZero = ret & cStr(val)
End Function

sub CreateServerArray(byRef myFSO, byRef myArray, byVal myPath)
	dim i
	dim server_list
	
	i = 1
	
	set server_list = myFSO.OpenTextfile(myPath)
	
	if CheckError("create server array") = false then
		close myFSO
	end if
	
 	Do While Not server_list.AtEndOfLine
 		redim preserve myArray(i)
 		myArray(i - 1) = cstr(server_list.ReadLine)
 		i = i + 1
 	Loop
	
	server_list.close
end sub

sub CreateLogArray(byRef myServerArray, byRef myTargetArray, byVal myFolder, byVal myLog)
	dim i
	dim val
	
	i = 1
	
	For Each val In myServerArray
		redim preserve myTargetArray(i)
		
		if CheckZeroLengthString(val) = true then
		else
			myTargetArray(i - 1) = myFolder & val & "_" & myLog
			i = i + 1
		end if
	next
end sub

sub CreatePerfLogLoop(byRef myArray, byRef myFSO, byRef myFile)
	dim val
	
	For Each val In myArray
		if CheckZeroLengthString(val) = true then
		else
			if myFSO.FileExists(val) = false then
				CreatePerfLog myFSO, myFile, val
			end if
		end if
	next
end sub

sub CreatePerfLog(byRef myFSO, byRef myFile, byVal file_path)
	set myFile = myFSO.CreateTextFile(file_path)
	
	if CheckError("create perf_log") = false then
		closer myFSO
	else
		CreateTitle myFile
		myFile.close
	end if
end sub

sub CreateTitle(byRef myObj)
	myObj.WriteLine "Time" & vbTab &_
		"PercentProcessorTime" & vbTab &_
		"AvailableMBytes" & vbTab &_
		"FileReadBytesPerSec" & vbTab &_
		"FileWriteBytesPerSec" & vbTab &_
		"FilesOpen" & vbTab &_
		"LogonPerSec" & vbTab &_
		"PoolNonPagedBytes" & vbTab &_
		"PoolNonPagedFailures" & vbTab &_
		"PoolPagedBytes" & vbTab &_
		"PoolPagedFailures" & vbTab &_
		"ServerSessions" & vbTab &_
		"CurrentDiskQueueLength"
end sub

sub CreateCSVLoop(byRef myFlg, byRef myFSO, byRef MyFile, byRef myLogArray, byRef myServerArray, byVal myDate)
	dim val
	dim i
	dim my_wmi
	dim my_processor
	dim my_memory
	dim my_system
	dim my_server
	dim my_disk
	
	i = 1
	
	For Each val In myServerArray
		if CheckZeroLengthString(val) = true then
		else
			OpenPerfLog myFSO, myFile, myLogArray(i - 1)
			InitObj my_wmi, val, my_processor, my_memory, my_system, my_server, my_disk
			CreateCSV my_file, my_processor, my_memory, my_system, my_server, my_disk
			KillObj my_wmi, my_processor, my_memory, my_system, my_server, my_disk
			my_file.close
			i = i + 1
		end if
	next
	
	wscript.sleep 5 * 60 * 1000
	
	if CheckLifeTime(myDate) = true then
		my_flg = false
	end if
end sub

sub OpenPerfLog(byRef myFSO, byRef myFile, byVal file_path)
	set myFile = myFSO.OpenTextFile(file_path, 8, false, 0)
	
	if CheckError("open perf_log") = false then
		closer myFSO
	end if
end sub

sub InitObj(byRef myWMI, byVal myServer, byRef myCPU, byRef myRAM, byRef mySYS, byRef mySVR, byRef myDSK)
	set myWMI = GetObject("winmgmts:!\\" & myServer & "\root\cimv2")
	set myCPU = myWMI.get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
	set myRAM = myWMI.get("Win32_PerfRawData_PerfOS_Memory=@")
	set mySYS = myWMI.get("Win32_PerfRawData_PerfOS_System=@")
	set mySVR = myWMI.get("Win32_PerfRawData_PerfNet_Server=@")
	set myDSK = myWMI.get("Win32_PerfRawData_PerfDisk_PhysicalDisk.Name='_Total'")
end sub

sub CreateCSV (byRef myFile, byRef myCPU, byRef myRAM, byRef mySYS, byRef mySVR, byRef myDSK)
	myFile.WriteLine cstr(now) & vbTab &_
		myCPU.PercentProcessorTime & vbTab &_
		myRAM.AvailableMBytes & vbTab &_
		mySYS.FileReadBytesPerSec & vbTab &_
		mySYS.FileWriteBytesPerSec & vbTab &_
		mySVR.FilesOpen & vbTab &_
		mySVR.LogonPerSec & vbTab &_
		mySVR.PoolNonPagedBytes & vbTab &_
		mySVR.PoolNonPagedFailures & vbTab &_
		mySVR.PoolPagedBytes & vbTab &_
		mySVR.PoolPagedFailures & vbTab &_
		mySVR.ServerSessions & vbTab &_
		myDSK.CurrentDiskQueueLength
end sub

sub KillObj (byRef myWMI, byRef myCPU, byRef myRAM, byRef mySYS, byRef mySVR, byRef myDSK)
	set myDSK = nothing
	set mySVR = nothing
	set mySYS = nothing
	set myRAM = nothing
	set myCPU = nothing
	set myWMI = nothing
end sub

function CheckLifeTime(myDeadLine)
	if Time >= myDeadLine then
		CheckLifeTime = true
	else
		CheckLifeTime = false
	end if
	
' 	dim life_time
' 	
' 	life_time = 24 * 60 - 5
' 		
' 	if DateDiff("n", val, now) >= life_time then
' 		CheckLifeTime = true
' 	else
' 		CheckLifeTime = false
' 	end if
end function

sub closer(byRef myObj)
	set myObj = nothing
	msgbox "end"
	wscript.quit
end sub

Function CheckZeroLengthString(byVal val)
	if len(val) = 0 then
		CheckZeroLengthString = true
	end if
	
	if isNull(val) = true then
		CheckZeroLengthString = true
	end if
end function

Function CheckError(caller)
	dim msg
	
	On Error Resume Next
	
	if Err.Number <> 0 then
		msg = addLine(msg, "Error occured at " & caller)
		msg = addLine(msg, "Error Number: " & cStr(Err.Number))
		msg = addLine(msg, "Error Message: " & Err.Description)
		writeLog 2, msg
		
		checkError = false
	else
		checkError = true
	end if
	
	Err.Clear
End Function

Function addLine(myLine, val)
	On Error Resume Next
	
	myLine = myLine & vbCrLf & val
	addLine = myLine
End Function

Sub writeLog(stat, msg)
	dim myObj
	
	On Error Resume Next
	
	set myObj = WScript.CreateObject("WScript.Shell")
	myObj.LogEvent stat, msg
End Sub