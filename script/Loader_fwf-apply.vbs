#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'	SCRIPT USED TO APPLY FIREWALL FILTERS TO UNI INTERFACES ACCORDING TO MEF BW PROFILE
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 75
Dim nResult
Dim strLine
Dim nOverwrite
Dim strMonthMaxFileName, strFileString, strSkip, strFileButton, strFileInventory, strFileSession
Dim strDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke
Dim strDeviceID, strAccountID
Dim nDebug
Dim nIndex, nInd, nCount
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile
Dim vSession(30), vSettings
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand
Dim Platform
Dim objTab_L, objTab_R
Dim vWaitForCommit, vWaitForFtp
vWaitForftp = Array("ftp: connect: No route to host","ftp>")
vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
Dim strFileSettings
Dim vDelim, vParamNames
    Const SECURECRT_FOLDER = "SecureCRT Folder"
    Const WORK_FOLDER = "Work Folder"
    Const CONFIGS_FOLDER = "Configuration Files Folder"
    Const CONFIGS_PARAM  = "MEF Service Parameters"
    Const CONFIGS_GLOBAL  = "CONFIGS_GLOBAL"
    Const CONFIGS_RE0  = "CONFIGS_RE0"
    Const CONFIGS_RE1  = "CONFIGS_RE1"
    Const Node_Left_IP  = "Left Node IP"
    Const Node_Right_IP  = "Right Node IP"
    Const FTP_IP  = "FTP IP"
    Const FTP_User  = "FTP User"
    Const FTP_Password  = "FTP Password"
	Const UNI_A = "UNI-A"
	Const UNI_B = "UNI-B"
	Const UNI_C = "UNI-C"
	Const UNI_D = "UNI-D"
	Const UNI_CC = "UNI-CC"
	Const UNI_DD = "UNI-DD"
	Const NNI = "NNI"
	Const PLATFORM_NAME = "Platform Name"
	Const PLATFORM_INDEX = "Node Name Prefix"
	Const Template = "XLS TEMPLATE"
	Const Orig_Folder = "Original TCG Templates"
	Const Dest_Folder = "Exported TCG Templates"
	Const WorkBookPrefix = "WorkBookPrefix"
	Const SECURECRT_L_SESSION = "Left Node Session"
	Const SECURECRT_R_SESSION = "Right Node Session"
ReDim vSettings(30)
vDelim = Array("=",",",":")	
nDebug = 0
nInfo = 1
Platform = "acx"
strVersion = "None"
strFileSettings = "settings.dat"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = CreateObject("WScript.Shell")
Sub Main()
'------------------------------------------------------------------
'	CHECK NUMBER OF ARGUMENTS AND EXIT IF LESS THEN 3
'------------------------------------------------------------------
	If crt.Arguments.Count < 3 Then
			MsgBox "ERROR: Wrong number of arguments" & chr(13) &_
			"ARG1: Name of the fw filter to submit" & chr(13) &_
			"ARG2: Full Path to settings.dat file" & chr(13) &_
			"ARG3: Full Path for Loader Work Folder"
		crt.quit
		Exit Sub
	End If
	strFilter = crt.Arguments(0)
	strFileSettings = crt.Arguments(1)
	strDirectoryWork = crt.Arguments(2)	
'----------------------------------------------------------------
'	Open log File
'----------------------------------------------------------------
			n = 5
			i = 0
			nRetries = 5
				Do While i < nRetries
					On Error Resume Next
					Err.Clear
						Set objDebug = objFSO.OpenTextFile(strDirectoryWork & "\Log\" & "debug-terminal.log",ForWriting,True)
						Select Case Err.Number
							Case 0
								Exit Do
							Case 70
								i =  i + 1
								n = 3
								crt.sleep 100 * n
							Case Else 
								Exit Do		
						End Select
				Loop
				On Error goto 0
'--------------------------------------------------------------------
'   LOOKING FOR EXISTED MONITOR SESSION (tail.exe)
'--------------------------------------------------------------------
strLaunch = strDirectoryWork & "\bin\tail.exe -f " & strDirectoryWork & "\log\debug-terminal.log"
If Not GetAppPID(strPID, strParentPID, "tail.exe") Then 
    objEnvar.run (strLaunch)
Else
    Call FocusToParentWindow(strPID)
End If
Call TrDebug_No_Date ("GetMyPID: PID = " & strPID & " ParentPID = " & strParentPID,"",objDebug, MAX_LEN, 1, nDebug)				
'-------------------------------------------------------------------------------------------
'  	LOAD INITIAL CONFIGURATION FROM SETTINGS FILE
'-------------------------------------------------------------------------------------------
	If objFSO.FileExists(strFileSettings) Then 
		nSettings = GetFileLineCountByGroup(strFileSettings, vLines,"Settings","","",0)
		For nInd = 0 to nSettings - 1 
			Select Case Split(vLines(nInd),"=")(0)
					Case SECURECRT_FOLDER
								vSettings(5) = vLines(nInd)
								strDirectoryVandyke = Split(vLines(nInd),"=")(1)
					Case WORK_FOLDER
								vSettings(6) = WORK_FOLDER & "=" & strDirectoryWork
					Case CONFIGS_FOLDER
								vSettings(7) = vLines(nInd)
								strDirectoryConfig =  Split(vLines(nInd),"=")(1)
					Case CONFIGS_PARAM
								vSettings(12) = vLines(nInd)
								strFileParam =  Split(vLines(nInd),"=")(1)
					Case CONFIGS_GLOBAL
								vSettings(27) = vLines(nInd)
								strCfgGlobal =  Split(vLines(nInd),"=")(1)
					Case CONFIGS_RE0
								vSettings(28) = vLines(nInd)
								strCfgRE0 =  Split(vLines(nInd),"=")(1)
					Case CONFIGS_RE1
								vSettings(29) = vLines(nInd)
								strCfgRE1 =  Split(vLines(nInd),"=")(1)
					Case PLATFORM_NAME
					            vSettings(13) = vLines(nInd)
								DUT_Platform = Split(vLines(nInd),"=")(1)
					Case PLATFORM_INDEX
					            vSettings(14) = vLines(nInd)					
								Platform = Split(vLines(nInd),"=")(1)
					Case Node_Left_IP
								vSettings(0) = vLines(nInd)
								strLeft_ip =  Split(vLines(nInd),"=")(1)
					Case Node_Right_IP
								vSettings(1) = vLines(nInd)
								strRight_ip =  Split(vLines(nInd),"=")(1)
					Case FTP_IP
								vSettings(2) = vLines(nInd)
								strFTP_ip =  Split(vLines(nInd),"=")(1)
					Case FTP_User
								vSettings(3) = vLines(nInd)
								strFTP_name =  Split(vLines(nInd),"=")(1)
					Case FTP_Password
								vSettings(4) = vLines(nInd)
								strFTP_pass =  Split(vLines(nInd),"=")(1)
			End Select
		Next
	End If
'--------------------------------------------------------------------------------
'          GET NAME OF THE TELNET SESSIONS
'--------------------------------------------------------------------------------
	nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Sessions","","",0)
	For nInd = 0 to nInventory - 1
		Select Case Split(vLines(nInd),"=")(0)
			Case SECURECRT_L_SESSION
						vSettings(10) = Split(vLines(nInd),"=")(1)
						strFolder = Split(vSettings(10), ",")(0) & "/"
						strSessionL = Split(vSettings(10), ",")(1)
						strHostL = Split(vSettings(10), ",")(2) 
						strLogin = Split(vSettings(10), ",")(3)
			Case SECURECRT_R_SESSION
						vSettings(11) = Split(vLines(nInd),"=")(1)
						strFolder = Split(vSettings(11), ",")(0) & "/"
						strSessionR = Split(vSettings(11), ",")(1)
						strHostR = Split(vSettings(11), ",")(2) 
						strLogin = Split(vSettings(11), ",")(3)
		End Select
	Next
'-------------------------------------------------------------------------------------------
'        SET SERVICE PARAM FULL PATH
'-------------------------------------------------------------------------------------------
	strFileParam = strDirectoryWork & "\config\" & strFileParam
'-------------------------------------------------------------------------------------------
'  	LOAD TESTBED TOPOLOGY
'-------------------------------------------------------------------------------------------
	If Not objFSO.FileExists(strFileParam) Then 
		MsgBox "MEF Parameters File: " & chr(13) & strFileParam  & " not found!"
		Exit Sub
	End If
'------------------------------------------------------------------
'	INITIAL CONFIGURATION FILES
'------------------------------------------------------------------
	strGlobalFileL = strCfgGlobal & "-" & Platform & "-l.conf"
	strGlobalFileR = strCfgGlobal & "-" & Platform & "-r.conf"
	strRe0FileL = strCfgRE0 & "-" & Platform & "-l.conf"
	strRe0FileR = strCfgRE0 & "-" & Platform & "-r.conf"
'------------------------------------------------------------------
'	LOG MAIN VARIABLES
'------------------------------------------------------------------
	Call TrDebug_No_Date ("TelnetScript: strSessionL, strHostL: " & strFolder & strSessionL & ", " & strHostL,"", objDebug, MAX_LEN, 1, nDebug)						
	Call TrDebug_No_Date ("TelnetScript: " & strGlobalFileL & ", " & strGlobalFileR,"", objDebug, MAX_LEN, 1, nDebug)						
	Call TrDebug_No_Date ("TelnetScript: " & strRe0FileL & ", " & strRe0FileR, "", objDebug, MAX_LEN, 1, nDebug)						
'------------------------------------------------------------------
'	START MAIN PROGRAM
'------------------------------------------------------------------
	'--------------------------------------------------------------------------------
    '  Start SSH session to Left Node
    '--------------------------------------------------------------------------------
	On Error Resume Next
	Err.Clear
	Set objTab_L = crt.Session.ConnectInTab("/S " & strFolder & strSessionL)
	If Err.Number <> 0 Then 
		Call  TrDebug_No_Date ("ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, 1)
		Exit Sub
	End If
	On Error Goto 0
	Call TrDebug_No_Date ("APPLYING FILTER: " & strFilter , "", objDebug, MAX_LEN, 3, nInfo)	
	objTab_L.Caption = strHostL
	objTab_L.Screen.Synchronous = True
	objTab_L.Screen.Send chr(13)
	nResult = objTab_L.Screen.WaitForString ("@" & strHostL & ">",2)
    If nResult = 0  Then
		If IsObject(objDebug) Then Call  TrDebug_No_Date (strHostL & " ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1) End If
		objTab_L.Session.Disconnect
    	Exit Sub
	End If
	objTab_L.Screen.Send "edit" & chr(13)
	objTab_L.Screen.WaitForString "@" & strHostL & "#"
	objTab_L.Screen.Send "edit interfaces" & chr(13)
	objTab_L.Screen.WaitForString "@" & strHostL & "#"
	objTab_L.Screen.Send "show |no-more" & chr(13)
	strLine = objTab_L.Screen.ReadString ("#")
	i = 0
	strOldFilter = ""
	nResult = 0
	Do While i < UBound(Split(strLine,chr(13)))
		If InStr(Split(strLine,chr(13))(i),"f-tc-") Then 
		  if InStr(Split(strLine,chr(13))(i),strFilter) Then nResult = 2 : Exit Do : End If
		  strOldFilter = Split(strLine,chr(13))(i)
		  j = 0
		  Do While j < UBound(Split(strOldFilter," ")) + 1
			If InStr(Split(strOldFilter," ")(j),"f-tc-") Then 
				strOldFilter = Split(strOldFilter," ")(j)
				If Right(strOldFilter,1) = ";" Then 
					strOldFilter = Left(strOldFilter,Len(strOldFilter)-1)
				End If
				Call TrDebug_No_Date ("FOUND OLD FILTER: " & strOldFilter , "OK", objDebug, MAX_LEN, 1, nInfo)
				nResult = 1
				Exit Do
			End If
			j = j + 1
		  Loop
		  Exit Do
		End If
		i = i + 1
	Loop
	objTab_L.Screen.Send chr(13)
	objTab_L.Screen.WaitForString "@" & strHostL & "#"
	Select Case nResult 
        Case 1
		objTab_L.Screen.Send "replace pattern " & strOldFilter & " with " & strFilter & chr(13)
		objTab_L.Screen.WaitForString "@" & strHostL & "#"
		Call TrDebug_No_Date ("REPLACING PATTERN: " & strOldFilter & " with " & strFilter, "OK", objDebug, MAX_LEN, 1, nInfo)
		Call  TrDebug_No_Date ("COMMIT " & strHostL, "......IN PROGRESS", objDebug, MAX_LEN, 1, nInfo)   
		'---------------------------------------------
		'   COMMIT L
		'---------------------------------------------	
		objTab_L.Screen.Send "commit" & chr(13)
		nResult = objTab_L.Screen.WaitForStrings (vWaitForCommit, 60)
		Select Case nResult
			Case 0
				Call  TrDebug_No_Date ("COMMIT " & strHostL, "TIME OUT", objDebug, MAX_LEN, 1, nInfo)   
				Exit Sub
			Case 1 
				Call  TrDebug_No_Date ("COMMIT " & strHostL, "ERROR 1", objDebug, MAX_LEN, 1, nInfo)   
				Exit Sub
			Case 2 
				Call  TrDebug_No_Date ("COMMIT " & strHostL, "ERROR 2", objDebug, MAX_LEN, 1, nInfo)   
				Exit Sub
			Case Else
				Call  TrDebug_No_Date ("COMMIT " & strHostL, "OK", objDebug, MAX_LEN, 1, nInfo)   
		End Select	
		objTab_L.Screen.Send chr(13)	
		objTab_L.Screen.WaitForString "@" & strHostL & "#"
		objTab_L.Screen.Send "exit" & chr(13)
		objTab_L.Screen.WaitForString "@" & strHostL & "#"
	    Case 0 
			Call TrDebug_No_Date ("WARNING: L-NODE OLD FILTER WAS NOT FOUND ", "", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("WARNING: L-NODE NEW FILTER WAS NOT APPLIED ", "", objDebug, MAX_LEN, 1, nInfo)		
			objTab_L.Screen.Send "exit" & chr(13)
			objTab_L.Screen.WaitForString "@" & strHostL & "#"
			objTab_L.Screen.Send "rollback" & chr(13)
			objTab_L.Screen.WaitForString "@" & strHostL & "#"		
        Case 2		
  			Call TrDebug_No_Date ("FILTER HAS BEEN ALREDY APPLYED ON L-NODE ", "", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("NO CONFIGURATION CHANGE WAS DONE ON L-NODE ", "", objDebug, MAX_LEN, 1, nInfo)		
			objTab_L.Screen.Send "exit" & chr(13)
			objTab_L.Screen.WaitForString "@" & strHostL & "#"
			objTab_L.Screen.Send "rollback" & chr(13)
			objTab_L.Screen.WaitForString "@" & strHostL & "#"		
	End Select
	objTab_L.Screen.Send "exit" & chr(13)	
	objTab_L.Screen.WaitForString "@" & strHostL & ">"

	'--------------------------------------------------------------------------------
    '  Start SSH session to Right Node
    '--------------------------------------------------------------------------------
	On Error Resume Next
	Err.Clear
	Set objTab_R = crt.Session.ConnectInTab("/S " & strFolder & strSessionR)
	If Err.Number <> 0 Then 
		Call  TrDebug_No_Date ("ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, nInfo)
		Exit Sub
	End If
	On Error Goto 0
	objTab_R.Caption = strHostR
	objTab_R.Screen.Synchronous = True
	objTab_R.Screen.Send chr(13)
	nResult = objTab_R.Screen.WaitForString ("@" & strHostR & ">",2)
    If nResult = 0  Then
		If IsObject(objDebug) Then Call  TrDebug_No_Date (strHostR & " ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, nInfo) End If
		objTab_R.Session.Disconnect
    	Exit Sub
	End If
	objTab_R.Screen.Send "edit" & chr(13)
	objTab_R.Screen.WaitForString "@" & strHostR & "#"
	objTab_R.Screen.Send "edit interfaces" & chr(13)
	objTab_R.Screen.WaitForString "@" & strHostR & "#"
	objTab_R.Screen.Send "show |no-more" & chr(13)
	strLine = objTab_R.Screen.ReadString ("#")
	i = 0
	strOldFilter = ""
	nResult = 0
	Do While i < UBound(Split(strLine,chr(13)))
		If InStr(Split(strLine,chr(13))(i),"f-tc-") Then
		  if InStr(Split(strLine,chr(13))(i),strFilter) Then nResult = 2 : Exit Do : End If		
		  strOldFilter = Split(strLine,chr(13))(i)
		  j = 0
		  Do While j < UBound(Split(strOldFilter," ")) + 1
			If InStr(Split(strOldFilter," ")(j),"f-tc-") Then 
				strOldFilter = Split(strOldFilter," ")(j)
				If Right(strOldFilter,1) = ";" Then 
					strOldFilter = Left(strOldFilter,Len(strOldFilter)-1)
				End If
				Call TrDebug_No_Date ("FOUND OLD FILTER: " & strOldFilter , "OK", objDebug, MAX_LEN, 1, nInfo)
				nResult = 1
				Exit Do
			End If
			j = j + 1
		  Loop
		  Exit Do
		End If
		i = i + 1
	Loop
	objTab_R.Screen.Send chr(13)
	objTab_R.Screen.WaitForString "@" & strHostR & "#"
	Select Case nResult
	    Case 1
			objTab_R.Screen.Send "replace pattern " & strOldFilter & " with " & strFilter & chr(13)
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
			Call TrDebug_No_Date ("REPLACING PATTERN: " & strOldFilter & " with " & strFilter, "OK", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("COMMIT " & strHostR, "......IN PROGRESS", objDebug, MAX_LEN, 1, nInfo)   
			'---------------------------------------------
			'   COMMIT R
			'---------------------------------------------	
			objTab_R.Screen.Send "commit" & chr(13)
			nResult = objTab_R.Screen.WaitForStrings (vWaitForCommit, 60)
			Select Case nResult
				Case 0
					Call  TrDebug_No_Date ("COMMIT " & strHostR, "TIME OUT", objDebug, MAX_LEN, 1, nInfo)   
					Exit Sub
				Case 1 
					Call  TrDebug_No_Date ("COMMIT " & strHostR, "ERROR 1", objDebug, MAX_LEN, 1, nInfo)   
					Exit Sub
				Case 2 
					Call  TrDebug_No_Date ("COMMIT " & strHostR, "ERROR 2", objDebug, MAX_LEN, 1, nInfo)   
					Exit Sub
				Case Else
					Call  TrDebug_No_Date ("COMMIT " & strHostR, "OK", objDebug, MAX_LEN, 1, nInfo)   
			End Select	
			objTab_R.Screen.Send chr(13)	
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
			objTab_R.Screen.Send "exit" & chr(13)
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
		Case 0 
			Call TrDebug_No_Date ("WARNING: R-NODE OLD FILTER WAS NOT FOUND ", "", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("WARNING: R-NODE NEW FILTER WAS NOT APPLIED ", "", objDebug, MAX_LEN, 1, nInfo)		
			objTab_R.Screen.Send "exit" & chr(13)
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
			objTab_R.Screen.Send "rollback" & chr(13)
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
		Case 2
  			Call TrDebug_No_Date ("FILTER HAS BEEN ALREDY APPLYED ON R-NODE ", "", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("NO CONFIGURATION CHANGE WAS DONE ON R-NODE ", "", objDebug, MAX_LEN, 1, nInfo)		
			objTab_R.Screen.Send "exit" & chr(13)
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
			objTab_R.Screen.Send "rollback" & chr(13)
			objTab_R.Screen.WaitForString "@" & strHostR & "#"
	End Select
	objTab_R.Screen.Send "exit" & chr(13)	
	objTab_R.Screen.WaitForString "@" & strHostR & ">"
    objTab_R.Session.Disconnect	
    objTab_L.Session.Disconnect	
	Call TrDebug_No_Date ("JOB DONE ", "", objDebug, MAX_LEN, 3, 1)	
	If IsObject(objDebug) Then objDebug.close : End If
'	objEnvar.Run "notepad.exe " & strDirectoryWork & "\Log\debug-terminal.log"	

	Set objFSO = Nothing
	Set objEnvar = Nothing
	crt.quit		
End Sub

'#######################################################################
' Function GetFileLineCount - Returns number of lines int the text file
'#######################################################################
 Function GetFileLineCount(strFileName, ByRef vFileLines, nDebug)
    Dim nIndex
	Dim strLine
	
    strFileWeekStream = ""	
	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		If 	InStr(strLine,"#") = 0 and InStr(strLine,"$") = 0 and strLine <> "" Then
			vFileLines(nIndex) = strLine
			If IsObject(objDebug) and nDebug = 1 Then objDebug.WriteLine "GetFileLineCount: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
			nIndex = nIndex + 1
		End If
	Loop
	objDataFileName.Close

    GetFileLineCount = nIndex
End Function
 '-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function

' ----------------------------------------------------------------------------------------------
'   Function  TrDebug_No_Date (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug_No_Date (strTitle, strString, objDebug, nChar, nFormat, nDebug)
Dim strLine
strLine = ""
If nDebug <> 1 Then Exit Function End If
If IsObject(objDebug) Then 
	Select Case nFormat
		Case 0
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) 
			strLine = strLine & ":  " & strTitle
			strLine = strLIne & strString
			objDebug.WriteLine strLine
			
		Case 1
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3)
			strLine = strLine & ":  " & strTitle
			If nChar - Len(strLine) - Len(strString) > 0 Then 
				strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
			Else 
				strLine = strLine & " " & strString
			End If
			objDebug.WriteLine strLine
		Case 2
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			
			If nChar - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
		Case 3
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine

	End Select
						
End If
End Function
'#######################################################################
 ' Function GetFileLineCountSelect - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCountSelect(strFileName, ByRef vFileLines,strChar1, strChar2, strChar3, nDebug)
    Dim nIndex
	Dim strLine, nCount, nSize
	Dim objDataFileName
	
    strFileWeekStream = ""	
	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	If nDebug = 1 Then objDebug.WriteLine "           GETTING SIZE OF THE FILE FIRST        "
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		Select Case Left(strLine,1)
			Case strChar1
			Case strChar2
			Case strChar3
			Case Else
					nIndex = nIndex + 1
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect:    " & strLine  End If  
		End Select
	Loop
	nCount = nIndex
	objDataFileName.close

    Redim vFileLines(nCount)
	nSize = UBound(vFileLines)
	Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	If nDebug = 1 Then objDebug.WriteLine "File Size: " & nCount & " Array Size: " & nSize
	If nDebug = 1 Then objDebug.WriteLine "           NOW TRYING TO RIGHT INTO AN ARRAY        "
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		Select Case Left(strLine,1)
			Case strChar1
			Case strChar2
			Case strChar3
			Case Else
					vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountSelect = nIndex
End Function
 '#######################################################################
 ' Function GetFileLineCountByGroup - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCountByGroup(strFileName, ByRef vFileLines, strGroup1, strGroup2, strGroup3, nDebug)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	
	GetFileLineCountByGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: -------------- GET FILE SIZE FIRST-----------" End If  
    Do While objDataFileName.AtEndOfStream <> True
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
'		strLine = All_Trim(objDataFileName.ReadLine)
		Select Case Left(strLine,1)
			Case "#"
'			Case "$"
			Case ""
			Case "["
				If strGroup1 = "All" Then 
					nGroupSelector = 1 
				Else 
					Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case "[" & strGroup2 & "]"
							nGroupSelector = 1
						Case "[" & strGroup3 & "]"
							nGroupSelector = 1
						Case Else
							nGroupSelector = 0
					End Select
				End If
			Case Else	
				If nGroupSelector = 1 Then nIndex = nIndex + 1 End If
		End Select
	Loop
	objDataFileName.Close
	Redim vFileLines(nIndex)
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: -------------- NOW READ TO ARRAY -----------" End If  
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
'		strLine = All_Trim(objDataFileName.ReadLine)
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
		Select Case Left(strLine,1)
			Case "#"
'			Case "$"
			Case ""
			Case "["
				If strGroup1 = "All" Then 
					nGroupSelector = 1 
				Else 
					Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case "[" & strGroup2 & "]"
							nGroupSelector = 1
						Case "[" & strGroup3 & "]"
							nGroupSelector = 1
						Case Else
							nGroupSelector = 0
					End Select
				End If
			Case Else	
				If nGroupSelector = 1 Then
					vFileLines(nIndex) = NormalizeStr(strLine, vDelim)
					' vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: vFileLines(" & nIndex & "): "  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
				End If
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountByGroup = nIndex
End Function
'-----------------------------------------------------------------
'     Function Normalize(strLine) - Removes all spaces around delimiters: arg1...arg3
'-----------------------------------------------------------------
Function NormalizeStr(strLine, vDelim)
Dim strNew
	strLine = LTrim(RTrim(strLine))
	strNew = ""
	For nInd = 0 to UBound(vDelim) - 1
		i = 0 
		Do While i <= UBound(Split(strLine,vDelim(nInd)))
				If UBound(Split(strLine,vDelim(nInd))) = 0 Then Exit Do End If
				If i < UBound(Split(strLine,vDelim(nInd))) Then strNew = strNew & LTrim(RTrim(Split(strLine,vDelim(nInd))(i))) & vDelim(nInd) End If
				If i = UBound(Split(strLine,vDelim(nInd))) Then strNew = strNew & LTrim(RTrim(Split(strLine,vDelim(nInd))(i))) End If
				i = i + 1
		Loop
		If i > 0 Then strLine = strNew End If
		strNew = ""
	Next
	NormalizeStr = strLine
End Function 
'-----------------------------------------------------------------
'     Function All_Trim(strLine) - Removes all speces form the string
'-----------------------------------------------------------------
Function All_Trim(strLine)
Dim nChar, strChar, i, strResult
	strResult = ""
	nChar = Len(strLine)
	For i = 1 to nChar
		strChar = Mid(strLine,i,1)
		If strChar <> " " Then strResult = strResult & strChar End If
	Next
		All_Trim = strResult
End Function

' ----------------------------------------------------------------------------------------------
'   Function  TrDebug_No_Date (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug_No_Date (strTitle, strString, objDebug, nChar, nFormat, nDebug)
Dim strLine
strLine = ""
If nDebug <> 1 Then Exit Function End If
If IsObject(objDebug) Then 
	Select Case nFormat
		Case 0
			strLine = ""
			strLine = strLine & ":  " & strTitle
			strLine = strLIne & strString
			objDebug.WriteLine strLine
			
		Case 1
			strLine = ""
			strLine = strLine & ":  " & strTitle
			If nChar - Len(strLine) - Len(strString) > 0 Then 
				strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
			Else 
				strLine = strLine & " " & strString
			End If
			objDebug.WriteLine strLine
		Case 2
			strLine = ""
			
			If nChar - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
		Case 3
			strLine = ""
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine
			strLine = ""
			If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
			strLine = ""
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine

	End Select
						
End If
End Function
'----------------------------------------------------------------
'   Function FocusToParentWindow(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function FocusToParentWindow(strPID)
Dim objShell
Call TrDebug_No_Date ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 0) 
Const IE_PAUSE = 70
	Set objShell = CreateObject("WScript.Shell")
	crt.sleep IE_PAUSE  
	objShell.SendKeys "%"	
	crt.sleep IE_PAUSE
	objShell.AppActivate strPID			
	crt.sleep IE_PAUSE  
	objShell.SendKeys "% "
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function GetAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetAppPID(ByRef strPID, ByRef strParentPID, strAppName)
Dim objWMI, colItems
Const IE_PAUSE = 70
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug_No_Date ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 0)
				On error Goto 0 
				Exit Do
		End If 
'		wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'Launcher Ver.'"  WHERE Name = 'iexplore.exe' OR Name = 'wscript.exe'
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug_No_Date ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug_No_Date ("GetMyPID: RESTORE IE WINDOW:", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 0) 
			If pUser = strUser then 
				strPID = process.ProcessId
				strParentPID = Process.ParentProcessId
'				Call TrDebug_No_Date ("GetMyPID: ", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
'				Call TrDebug_No_Date ("GetMyPID: ", "Caption: " & process.Caption & ", CSName " & process.CSName & ", Description: " & process.Description & ", Handle: " &  Process.Handle,objDebug, MAX_LEN, 1, 1) 
			GetAppPID = True
				Exit For
			End If
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUserSYS
'----------------------------------------------------------------------------------
Function GetScreenUserSYS()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserSYS = strScreenUser
End Function