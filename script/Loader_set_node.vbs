#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'	JUNIPER MEF CONFIG UPLOAD SCRIPT
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 75
Const HttpTextColor1 = "#292626"
Const HttpTextColor2 = "#F0BC1F"
Const HttpTextColor3 = "#EBEAF7"
Const HttpTextColor4 = "#A4A4A4"
Const HttpBgColor1 = "Grey"
Const HttpBgColor2 = "#292626" 
Const HttpBgColor3 = "#2C2A23" 
Const HttpBgColor4 = "#504E4E"
Const HttpBgColor5 = "#0D057F"
Const HttpBgColor6 = "#8B9091"
Const HttpBdColor1 = "Grey"
Dim nResult
Dim strLine
Dim nOverwrite
Dim strDirectoryWork, strDirectoryVandyke
Dim nDebug, nInfo
Dim nIndex, nInd, nCount
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile
Dim vSession(30), vSettings
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand
Dim Platform
Dim vWaitForCommit, vModels, vWaitForPrompt
Dim vIE_Scale
vWaitForPrompt = Array(">","#","%")
vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
vModels = Array("acx5096","acx5048","acx1100","acx1000","acx2100","acx2200","mx80","mx104","mx240","mx480","mx960")
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
    Const CRT_REG_INSTALL = "HKEY_LOCAL_MACHINE\SOFTWARE\VanDyke\SecureCRT\Install\Main Directory"
    Const CRT_REG_SESSION = "HKEY_CURRENT_USER\Software\VanDyke\SecureCRT\Config Path"
	Const MEF_CFG_LDR_REG = "HKEY_CURRENT_USER\Software\JnprMCL\Main Directory"
	Const FTP_REG_INSTALL = "HKEY_CURRENT_USER\Software\FileZilla Server\Install_Dir"
	Const NOTEPAD_PP = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe\"
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	
ReDim vSettings(30)
vDelim = Array("=",",",":")	
nDebug = 0
nInfo = 1
Platform = "acx"
strVersion = "None"
strFileSettings = "settings.dat"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = CreateObject("WScript.Shell")
Set objShell = CreateObject("WScript.Shell")
crt.Screen.Synchronous = True
Sub Main()
	'-----------------------------------------------------------------
	'  GET SCREEN RESOLUTION
	'-----------------------------------------------------------------
		Call GetScreenResolution(vIE_Scale, 0)
	'------------------------------------------------------------------
	'	CHECK NUMBER OF ARGUMENTS AND EXIT IF LESS THEN 3
	'------------------------------------------------------------------
	nResult = 0
	On Error Resume Next
		Err.Clear
		strDirectoryWork = objShell.RegRead(MEF_CFG_LDR_REG)
		if Err.Number <> 0 Then 
			strDirectoryWork = "Not Set"
			MsgBox "Error: Can't find work folder for MEF Loader"  & chr(13) & "Exit Now!"
			Exit Sub
		End If
	On Error Goto 0
	strPlatform = ""
	strFileSettings = strDirectoryWork & "\config\settings.dat"						
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
'strLaunch = strDirectoryWork & "\bin\tail.exe -f " & strDirectoryWork & "\log\debug-terminal.log"
'If Not GetAppPID(strPID, strParentPID, "tail.exe") Then 
'    objEnvar.run (strLaunch)
'Else
'    Call FocusToParentWindow(strPID)
'End If
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
								strPlatform = Split(vLines(nInd),"=")(1)
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
	'-------------------------------------------------------
	' CHECK THE PLATFORM TYPE
	'-------------------------------------------------------
	For nResult = 0 to Ubound(vModels)
		If LCase(strPlatform) = LCase(vModels(nResult)) Then Exit For
    Next
    Select Case nResult
        Case 0, 1 
		  strRe0 = " member0 "
        Case Else
          strRe0 = " re0 "
     End Select
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
	strGlobalFileL = strDirectoryConfig & "\" & strCfgGlobal & "-" & Platform & "-l.conf"
	strGlobalFileR = strDirectoryConfig & "\" &  strCfgGlobal & "-" & Platform & "-r.conf"
	strRe0FileL = strDirectoryConfig & "\" &  strCfgRE0 & "-" & Platform & "-l.conf"
	strRe0FileR = strDirectoryConfig & "\" &  strCfgRE0 & "-" & Platform & "-r.conf"
'------------------------------------------------------------------
'	LOG MAIN VARIABLES
'------------------------------------------------------------------
	Call TrDebug_No_Date ("TelnetScript: strSessionL, strHost: " & strFolder & strSessionL & ", " & strHost,"", objDebug, MAX_LEN, 1, nDebug)						
	Call TrDebug_No_Date ("TelnetScript: " & strGlobalFileL & ", " & strGlobalFileR,"", objDebug, MAX_LEN, 1, nDebug)						
	Call TrDebug_No_Date ("TelnetScript: " & strRe0FileL & ", " & strRe0FileR, "", objDebug, MAX_LEN, 1, nDebug)						
'------------------------------------------------------------------
'	START MAIN PROGRAM
'------------------------------------------------------------------
    Call TrDebug_No_Date ("NOW WE SET UP DUT Node INITIAL CONFIGURATION ", "", objDebug, MAX_LEN, 3, nInfo)
	'--------------------------------------------------------------------------------
    '  Connect to existed terminal session
    '--------------------------------------------------------------------------------
	If Not crt.Session.Connected Then 
		Call TrDebug_No_Date ("CAN'T FIND ACTIVE CRT SESSION: ", "", objDebug, MAX_LEN, 1, nInfo)
		Call TrDebug_No_Date ("1. Open terminal session from SecureCRT to one of the DUT Node", "", objDebug, MAX_LEN, 1, nInfo)
		Call TrDebug_No_Date ("   Use Telnet, ssh or consol port to open a session", "", objDebug, MAX_LEN, 1, nInfo)
				Call TrDebug_No_Date ("CONNECTION FAILED", "", objDebug, MAX_LEN, 3, nInfo)
		Exit Sub
	End If
    If Not IE_Node_Input(vIE_Scale, "Select DUT Node ID", strNodeId, nDebug) Then Exit Sub
	Call TrDebug_No_Date ("START CONFIGURING " & strNodeId & " NODE", "", objDebug, MAX_LEN, 1, nInfo)
	'---------------------------------------------
	'   LOOKING FOR PROMPT
	'---------------------------------------------
	crt.Screen.Send chr(13)
	nResult = crt.Screen.WaitForStrings (vWaitForPrompt, 5)
     Select Case nResult
        Case 0
			Call  TrDebug_No_Date ("TIME OUT WAITING FOR THE PROMPT", "",objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("NODE SETUP FAILED " , "", objDebug, MAX_LEN, 3, nInfo)					
'           crt.Session.Disconnect			
'			crt.quit
			Exit Sub
        Case 1
           	crt.Screen.Send chr(13)
			strLine = crt.Screen.ReadString (">")
			If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
			crt.Screen.Send "edit" & chr(13)
			Call  TrDebug_No_Date ("GET > PROMPT","", objDebug, MAX_LEN, 1, nInfo)
        Case 2 
           	crt.Screen.Send chr(13)
			strLine = crt.Screen.ReadString ("#")
			If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
			crt.Screen.Send chr(13)
			Call  TrDebug_No_Date ("GET # PROMPT","", objDebug, MAX_LEN, 1, nInfo)
		Case 3
           	crt.Screen.Send "cli" & chr(13)
			strLine = crt.Screen.ReadString (">")
			If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
			crt.Screen.Send "edit" & chr(13)
			Call  TrDebug_No_Date ("GET % PROMPT", "",objDebug, MAX_LEN, 1, nInfo)		
     End Select	
	 crt.Screen.WaitForString "@" & strHost & "#"
'---------------------------------------------
'   READ Configuration for Global Configuration group
'---------------------------------------------
    Select Case strNodeID 
	    Case "Left"  
    	    Set configFile = objFSO.OpenTextFile(strGlobalFileL, 1, False)
	    Case "Right"  
     	    Set configFile = objFSO.OpenTextFile(strGlobalFileR, 1, False)
    End Select
	strConfig = configFile.ReadAll
	configFile.Close
	Set configFile = Nothing
'---------------------------------------------
'   LOAD global 
'---------------------------------------------
    crt.Screen.Send "edit groups global" & chr(13)	
	crt.Screen.WaitForString "@" & strHost & "#"
	crt.Screen.Send "load merge terminal relative" & chr(13)
	If Not crt.Screen.WaitForString ("input]", 5) Then
	    Call  TrDebug_No_Date ("LOAD CONFIGURATION TIME OUT", "FAILED", objDebug, MAX_LEN, 1, nInfo)
		exit sub		
    End If
	crt.Screen.Send strConfig & chr(13) & chr(4)
	If Not crt.Screen.WaitForString ("load complete", 5) Then
	    Call  TrDebug_No_Date ("GLOBAL CONFIGURATION LOAD", "FAILED", objDebug, MAX_LEN, 1, nInfo)
		exit sub
    Else
	    Call  TrDebug_No_Date ("GLOBAL CONFIGURATION LOADED ", "OK", objDebug, MAX_LEN, 1, nInfo)   
	End If
	crt.Screen.WaitForString "@" & strHost & "#"
	crt.Screen.Send "exit" & chr(13)
	crt.Screen.WaitForString "@" & strHost & "#"
'---------------------------------------------
'   READ Configuration for RE0  group
'---------------------------------------------
    Select Case strNodeID 
	    Case "Left"  
    	    Set configFile = objFSO.OpenTextFile(strRe0FileL, 1, False)
	    Case "Right"  
     	    Set configFile = objFSO.OpenTextFile(strRe0FileR, 1, False)
    End Select
	strConfig = configFile.ReadAll
	configFile.Close
	Set configFile = Nothing
'---------------------------------------------
'   LOAD RE0 Group Configuration
'---------------------------------------------
    crt.Screen.Send "edit groups" & strRe0  & chr(13)
	crt.Screen.WaitForString "@" & strHost & "#"
	crt.Screen.Send "load update terminal relative" & chr(13)
	If Not crt.Screen.WaitForString ("input]", 5) Then
	    Call  TrDebug_No_Date ("LOAD RE0 CONFIGURATION TIME OUT", "FAILED", objDebug, MAX_LEN, 1, nInfo)
		exit sub		
    End If
	crt.Screen.Send strConfig & chr(13) & chr(4)
	If Not crt.Screen.WaitForString ("load complete", 5) Then
	    Call  TrDebug_No_Date ("RE0 CONFIGURATION LOAD","FAILED", objDebug, MAX_LEN, 1, nInfo)
		exit sub
    Else
	    Call  TrDebug_No_Date ("RE0 CONFIGURATION LOADED ", "OK", objDebug, MAX_LEN, 1, nInfo)   
	End If
	crt.Screen.WaitForString "@" & strHost & "#"
	crt.Screen.Send "exit" & chr(13)
	crt.Screen.WaitForString "@" & strHost & "#"
'---------------------------------------------
'   APPLY-GROUPS re0 and global on L NODE
'---------------------------------------------
    crt.Screen.Send "set apply-groups global" & chr(13)	
	crt.Screen.WaitForString "@" & strHost & "#"	
    crt.Screen.Send "set apply-groups" & strRe0 & chr(13)	
	crt.Screen.WaitForString "@" & strHost & "#"	
'---------------------------------------------
'   COMMIT
'---------------------------------------------	
	crt.Screen.Send "commit" & chr(13)
	' - After commit change the prompt
	nResult = crt.Screen.WaitForStrings (vWaitForCommit, 30)
     Select Case nResult
        Case 0
			Call  TrDebug_No_Date ("COMMIT " & strHost, "TIME OUT", objDebug, MAX_LEN, 1, nInfo) 
			Call  TrDebug_No_Date ("Most probably You have lost your connectivity","", objDebug, MAX_LEN, 1, nInfo) 
			Call  TrDebug_No_Date ("as a result of setting up new IP address to the", "",objDebug, MAX_LEN, 1, nInfo) 
			Call  TrDebug_No_Date ("management address of the DUT Node", "",objDebug, MAX_LEN, 1, nInfo) 			
			Call  TrDebug_No_Date ("Try to to reconnect using it new IP-address", "", objDebug, MAX_LEN, 1, nInfo) 						
	        Call  TrDebug_No_Date ("JOB DONE", "",objDebug, MAX_LEN, 3, nInfo) 			
			Exit Sub
        Case 1 
			Call  TrDebug_No_Date ("COMMIT " & strHost, "ERROR 1", objDebug, MAX_LEN, 1, nInfo)   
		    Call  TrDebug_No_Date ("LOADING INITIAL CONFIGURATION HAS FAILED", "",objDebug, MAX_LEN, 3, nInfo) 			
			Exit Sub
        Case 2 
			Call  TrDebug_No_Date ("COMMIT " & strHost, "ERROR 2", objDebug, MAX_LEN, 1, nInfo)   
		    Call  TrDebug_No_Date ("LOADING INITIAL CONFIGURATION HAS FAILED", "",objDebug, MAX_LEN, 3, nInfo) 						
			Exit Sub
		Case Else
    End Select	
    Select Case strNodeID 
	    Case "Left"  
    	    strHost = strHostL
	    Case "Right"  
     	    strHost = strHostR
    End Select
	Call  TrDebug_No_Date ("COMMIT " & strHost, "OK", objDebug, MAX_LEN, 1, 1)   
 	crt.Screen.WaitForString "@" & strHost & "#"
	crt.Screen.Send "exit" & chr(13)
	crt.Screen.WaitForString "@" & strHost & ">"
	crt.Screen.Send "exit" & chr(13)
	Call TrDebug_No_Date ("JOB DONE ", "", objDebug, MAX_LEN, 3, 1)	
	If IsObject(objDebug) Then objDebug.close : End If
	Set objFSO = Nothing
	Set objEnvar = Nothing
'	crt.quit		
End Sub
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
'-----------------------------------------------------------------------------------------
'      Function Displays a Message with Continue and No Button. Returns True if Continue
'-----------------------------------------------------------------------------------------
 Function IE_Node_Input(vIE_Scale, strTitle, ByRef strNodeId, nDebug)
    Set g_objIE = Nothing
    Set objShell = Nothing
	Dim vLine(20,3), nLine
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_CONT = False
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		crt.Sleep 200
	Loop While g_objIE.Busy
'    With g_objIE.document.parentwindow.screen
'		intX = .availwidth
'        intY = .availheight
'    End With
    nLine = 5
	LineH = 13
	nRatioX = intX/1920
    nRatioY = intY/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	ColumnW1 = Round(50 * nRatioX,0)
	nTab = Int(CellW/4)
	nLeft = Round(20 * nRatioX,0)
	BottomH = Round(10 * nRatioY,0)
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	
 '  If nDebug = 1 Then MsgBox "intX=" & intX & "   intY=" & intY & "   RatioX=" & nRatioX & "  RatioY=" & nRatioY & "   Cell Width=" & cellW & "  Cell Hight=" & cellH End If
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	strHTMLBody = "<br>"
	'----------------------------------------------------------
	'    TEXT ON INFORMATION PANNEL
	'----------------------------------------------------------
	vLine(0,0) = "CHOOSE WHICH NODE YOU'D LIKE TO SETUP "            : vLine(0,1) = "Bold" : vLine(0,2) = HttpTextColor2
	vLine(1,0) = "<br>"
	vLine(2,0) = "<br>"											               : vLine(2,1) = "normal" : vLine(2,2) = HttpTextColor1 
	vLine(3,0) = "<br>"
	vLine(4,0) = "<br>"
    vLine(5,0) = "<br>"	
	vLine(6,0) = """YES"" TO CONTINUE OR ""NO"" TO CANCEL SETUP PROCESS" : vLine(5,1) = "normal" : vLine(5,2) = HttpTextColor1 
	nLine = 7
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
			
	Next		
	'----------------------------------------------------------
	'    FORMS TO SELECT DUT NODE ID
	'----------------------------------------------------------
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: " & nTab & " px; top: " & (2 + 2) * LineH * 2 & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: none; width: " & CellW - nTab & "px;"">" & _
		"<tbody>" &_
			"<tr>" &_
				"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" & 2 * LineH & """ width=""" & ColumnW1 & """>" & _
					"<input type=radio name='Node_ID' style=""color: " & HttpTextColor2 & """ value='Device_L' Accesskey='L' >" &_
				"</td>" &_
				"<td style=""border-style: none; background-color: none;""class=""oa1"" align=""Left"" height=""" & 2 * LineH & """ width=""" & CellW - ColumnW1 - nTab & """>" & _
				    "<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 & "; font-weight: normal;""><u>L</u>eft Node</p>" &_ 
				"</td>" &_
         	"</tr>" &_
			"<tr>" &_
				"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" & 2 * LineH & """ width=""" & ColumnW1 & """>" & _
				    "<input type=radio name='Node_ID' style=""color: " & HttpTextColor2 & """ value='Device_L' Accesskey='L' >" &_
				"</td>" &_
				"<td style=""border-style: none; background-color: none;""class=""oa1"" align=""Left"" height=""" & 2 * LineH & """ width=""" & CellW - ColumnW1 - nTab & """>" & _
					"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 & "; font-weight: normal;""><u>R</u>ight Node</p>" &_
				"</td>" &_
         	"</tr>" &_
    	"</tbody></table>"

    strHTMLBody = strHTMLBody &_
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nLeft & "px; bottom: " & BottomH & "px' name='Continue' AccessKey='Y' onclick=document.all('ButtonHandler').value='YES';><u>Y</u>ES</button>" & _
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; right: " & nLeft & "px; bottom: " & BottomH & "px' name='NO' AccessKey='N' onclick=document.all('ButtonHandler').value='NO';><u>N</u>O</button>" & _
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.Visible = True
	Do
		crt.Sleep 100
	Loop While g_objIE.Busy
	'----------------------------------------
	'	Set Initial form values
	'----------------------------------------
	g_objIE.Document.All("Node_ID")(0).Select
    g_objIE.Document.All("Node_ID")(0).Checked = false
    g_objIE.Document.All("Node_ID")(0).Click
	
	Set objShell = CreateObject("WScript.Shell")
	objShell.AppActivate g_objIE.document.Title
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "NO"
				IE_Node_Input = False
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
			Case "YES"
			    If g_objIE.Document.All("Node_ID")(0).Checked Then strNodeId = "Left"
				If g_objIE.Document.All("Node_ID")(1).Checked  Then strNodeId = "Right"
				IE_Node_Input = True
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
		End Select
		crt.Sleep 500
		Loop
End Function
'----------------------------------------------------------
'   Function Set_IE_obj (byRef objIE)
'----------------------------------------------------------
Function Set_IE_obj (byRef objIE)
	Dim nCount
	Set_IE_obj = False
	nCount = 0
	Do 
		On Error Resume Next
		Err.Clear
		Set objIE = CreateObject("InternetExplorer.Application")
		Select Case Err.Number
			Case &H800704A6 
				crt.sleep 1000
				nCount = nCount + 1
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				If nCount > 4 Then
					On Error goto 0
					Exit Function
				End If
			Case 0 
				Set_IE_obj = True
				On Error goto 0
				Exit Function
			Case Else 
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				On Error goto 0
				Exit Function
		End Select
	On Error goto 0
	Loop
End Function 
'-------------------------------------------------------------
'    Function GetScreenResolution(vIE_Scale, intX,intY)
'-------------------------------------------------------------
Function GetScreenResolution(ByRef vIE_Scale, nDebug)
Dim g_objIE, f_objShell, intX, intY, intXreal, intYreal	
Redim vIE_Scale(2,3)
	nInd = 0
	Call Set_IE_obj (g_objIE)
	With g_objIE
		.Visible = False
		.Offline = True	
		.navigate "about:blank"
		Do
			crt.sleep 200
		Loop While g_objIE.Busy	
		.Document.Body.innerHTML = "<p>tEST</p>"
		.MenuBar = False
		.StatusBar = False
		.AddressBar = False
		.Toolbar = False		
		.Document.body.scroll = "no"
		.Document.body.Style.overflow = "hidden"
		.Document.body.Style.border = "None " & HttpBdColor1
		.Height = 100
		.Width = 100
		OffsetX = .Width - .Document.body.clientWidth
		OffsetY = .Height - .Document.body.clientHeight
		.FullScreen = True
		.navigate "about:blank"	
		 intXreal = .width
		 intYreal = .height
		.Quit
	End With
	If intXreal => 1440 Then intX = 1920 else intX = intXreal
	If intYreal => 900 Then intY = 1080  else intY = intYreal
	vIE_Scale(0,0) = intX : vIE_Scale(0,1) = OffsetX : vIE_Scale(0,2) = intXreal 
	vIE_Scale(1,0) = intY : vIE_Scale(1,1) = OffsetY : vIE_Scale(1,2) = intYreal
	Set g_objIE = Nothing
End Function
