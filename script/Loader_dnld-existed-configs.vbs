#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'		JUNIPER MEF CONFIG DOWNLOAD SCRIPT
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 75
Dim nResult
Dim strLine
Dim nOverwrite
Dim strMonthMaxFileName, strFileString, strSkip, strFileButton, strFileInventory, strFileSession, strVersion
Dim strDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke
Dim strDeviceID, strAccountID
Dim nDebug, nInfo
Dim nIndex, nInd, nCount
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile
Dim vSession(30), vSettings
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand
Dim Platform
Dim objTab_L
Dim vWaitForCommit, vModels, vWaitForFtp,vLoadComplete, vCfgInventory
Dim vSessionCRT, bConnect, vLookForCfg, bSuccess
vWaitForftp = Array("No route to host","Connected to","Connection refused")
vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
vModels = Array("mx240","mx480","mx960","acx5096","acx5048","acx1100","acx1000","acx2100","acx2200","mx80","mx104")
vLoadComplete = Array("error", "invalid","#")
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
	If crt.Arguments.Count < 5 Then
			MsgBox "ERROR: Wrong number of arguments" & chr(13) &_
			"ARG1: Configuration Name" & chr(13) &_
			"ARG2: Configuration Version" & chr(13) &_
			"ARG3: Full Path to settings.dat file" & chr(13) &_
			"ARG4: Full Path for Loader Work Folder" & chr(13) &_
			"ARG5: Node Session 1" & chr(13) &_
			"ARGn: Node Session n"
		crt.quit
		Exit Sub
	End If
	strCfg = crt.Arguments(0)
	strVer = crt.Arguments(1)
	strFileSettings = crt.Arguments(2)
	strDirectoryWork = crt.Arguments(3)
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
	'          LOAD CONFIGURATION ATTRIBUTES
	'--------------------------------------------------------------------------------
	Dim vCfgAttributes
    strCfgFile = strDirectoryConfig & "\CfgList.txt"	
	Call GetFileLineCountByGroup(strCfgFile, vCfgInventory,"Inventory","","",0)
	vWaitForftp(1) = "Connected to " & strFTP_ip
	'--------------------------------------------------------------------------------
	'          GET LIST OF TERMINAL SESSIONS
	'--------------------------------------------------------------------------------
	Dim nSessionID
	Redim vSessionCRT(1)
	nSessionID = 0
	nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Sessions","","",0)
	For nInd = 0 to nInventory - 1
		If InStr(Split(vLines(nInd),"=")(0),SECURECRT_SESSION) > 0 Then
			Redim Preserve vSessionCRT(nSessionID + 1)
			vSessionCRT(nSessionID) = Split(vLines(nInd),"=")(1)
			nSessionID = nSessionID + 1
        End If 
	Next
    '--------------------------------
	' BEGIN MAIN CYCLE
	'--------------------------------
	' Call WriteStringToFile(strDirectoryConfig & "\DownloadedCfgList.txt","[Bulk_Dnld_" & Year(Date) & "_" & Month(Date()) & "_" & Day(Date()) & "_" & Time(), nDebug)
    Dim strHostL, strLogin, strSessionL, strFolder
	bFoundConfig = False
	For nSession = 4 to crt.Arguments.Count - 1
	    bSuccess = False
		bConnect = False
	    Do
			'--------------------------------------------------------------------------------
			'          GET NAME OF THE TELNET SESSIONS
			'--------------------------------------------------------------------------------
			For nInd = 0 to UBound(vSessionCRT) - 1
				If Split(vSessionCRT(nInd),",")(2) = crt.Arguments(nSession) Then 
					strFolder = Split(vSessionCRT(nInd), ",")(1) & "/"
					strSessionL = Split(vSessionCRT(nInd), ",")(2)
					strHostL = Split(vSessionCRT(nInd), ",")(3) 
					strLogin = Split(vSessionCRT(nInd), ",")(4)
					Exit For
				End If
			Next
			'------------------------------------------------------------------
			'	LOG MAIN VARIABLES
			'------------------------------------------------------------------
			Call TrDebug_No_Date ("TelnetScript: strSessionL, strHostL: " & strFolder & strSessionL & ", " & strHostL,"", objDebug, MAX_LEN, 1, nDebug)						
			'------------------------------------------------------------------
			'	START MAIN PROGRAM
			'------------------------------------------------------------------
			strConfigFileL = strSessionL & "-" & strCfg & ".conf"
			'--------------------------------------------------------------------------------
			'  Start SSH session to Node
			'--------------------------------------------------------------------------------
			Call TrDebug_No_Date ("START DOWNLOADING CONFIGURATION FILES FROM " & strSessionL,"", objDebug, MAX_LEN, 3, nInfo)						
			On Error Resume Next
			Err.Clear
			Set objTab_L = crt.Session.ConnectInTab("/S " & strFolder & strSessionL)
			If Err.Number <> 0 Then 
				Call  TrDebug_No_Date ("CAN'T CONNECT TO " & strSessionL & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, nInfo)
				bConnect = False
				bSuccess = False
				Exit Do
			End If
			On Error Goto 0
			bConnect = True
			objTab_L.Caption = strSessionL
			objTab_L.Screen.Synchronous = True
			'--------------------------------------------------------------------------------
			'  Get actual host name of the box
			'--------------------------------------------------------------------------------
			objTab_L.Screen.Send chr(13)
			strLine = objTab_L.Screen.ReadString (">")
			If InStr(strLine,"@") Then strHostL = Split(strLine,"@")(1)
			objTab_L.Screen.Send chr(13)
			nResult = objTab_L.Screen.WaitForString ("@" & strHostL & ">",5)
			If nResult = 0  Then
				If IsObject(objDebug) Then Call  TrDebug_No_Date (strSessionL & ": CAN'T GET RESPONSE FROM NODE", "ERROR", objDebug, MAX_LEN, 1, 1) End If
				objTab_L.Session.Disconnect
				bSuccess = False
				Exit Do
			End If
			objTab_L.Screen.WaitForString "@" & strHostL & "> "
			'---------------------------------------------
			'   CHECK IF CONFIGURATION EXISTS
			'---------------------------------------------
			vLookForCfg = Array(strConfigFileL,"@" & strHostL & ">")
			objTab_L.Screen.Send "file list |match " & strConfigFileL & chr(13)	
			objTab_L.Screen.WaitForString strConfigFileL ' <-- This is first occurrence of the config name from the  command it self.   
			nResult = objTab_L.Screen.WaitForStrings (vLookForCfg, 5)
			Select Case nResult
				Case 0
					Call TrDebug_No_Date ("CONFIGURATION " & strConfigFileL , "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
		'           objTab_L.Session.Disconnect			
		'			crt.quit
		            bSuccess = False
					Exit Do
				Case 1 
					Call TrDebug_No_Date ("CONFIGURATION " & strConfigFileL , "FOUND", objDebug, MAX_LEN, 1, nInfo)
					objTab_L.Screen.WaitForString "@" & strHostL & "> "
				Case 2 
					Call TrDebug_No_Date ("CONFIGURATION " & strConfigFileL , "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
                    bSuccess = False					
					Exit Do
			End Select	
			'--------------------------------------------------------------------------------------------
			'  Get actual configuration Date when connecting to the first node in the row
			'--------------------------------------------------------------------------------------------
			If Not bFoundConfig Then 
				objTab_L.Screen.Send "file show " & strConfigFileL & " |match ""Last change""" & chr(13)			
				strLine = objTab_L.Screen.ReadString (">")
				If InStr(strLine,"## Last change") Then 
				    strLine = Split(strLine,"## ")(1)
					If UBound(Split(strLine," ")) > 3 Then 
						strVersion = Split(strLine," ")(2)
					Else 
						strVersion = Year(Date) & "-" & Month(Date()) & "-" & Day(Date())
					End If
				Else 
				    strVersion = Year(Date) & "-" & Month(Date()) & "-" & Day(Date())
				End If 
				Call CreateNewCfg(strCfg, strVersion, strDirectoryConfig, vCfgInventory, vCfgAttributes, nDebug)
				Call WriteStringToFile(strDirectoryConfig & "\DownloadedCfgList.txt",strCfg, nDebug)
				bFoundConfig = True
			End If
			strDestFolder = "./" & strCfg & "/" & strVersion
			'---------------------------------------------
			'   START FTP SESSION
			'---------------------------------------------
			objTab_L.Screen.Send "ftp "  & strFTP_ip & chr(13)
			nResult = objTab_L.Screen.WaitForStrings (vWaitForFtp, 5)
			 Select Case nResult
				Case 0
					Call  TrDebug_No_Date (strSessionL & " Connecting to FTP server", "TIME OUT", objDebug, MAX_LEN, 1, nInfo)
					Call TrDebug_No_Date ("CONFIGURATION UPLOAD " & UCase(strSessionL) & " FAILED " , "", objDebug, MAX_LEN, 3, 1)					
		'           objTab_L.Session.Disconnect			
		'			crt.quit
		            bSuccess = False
					Exit Do
				Case 1 
					Call  TrDebug_No_Date (UCase(strSessionL) & " HAS NO ROUTE TO FTP SERVER", "", objDebug, MAX_LEN, 1, nInfo)
					Call  TrDebug_No_Date ("1. Check FTP Server IP-address under CFG Loader Settings ", "", objDebug, MAX_LEN, 1, nInfo)
					Call  TrDebug_No_Date ("2. Make sure the node has a route to FTP server IP address ","", objDebug, MAX_LEN, 1, nInfo)
					Call TrDebug_No_Date ("CONFIGURATION UPLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)		
					objTab_L.Session.Disconnect
					bSuccess = False
					Exit Do
				Case 2 
					Call  TrDebug_No_Date ( "CONNECTING TO FTP", "OK", objDebug, MAX_LEN, 1, nInfo)   
				Case 3
					Call  TrDebug_No_Date ("FTP CONNECTION REFUSED", "", objDebug, MAX_LEN, 1, nInfo)
					Call  TrDebug_No_Date ("1. Make sure that FTP Server is running ", "", objDebug, MAX_LEN, 1, nInfo)
					Call  TrDebug_No_Date ("2. Make sure connection is not blocked by firewall","", objDebug, MAX_LEN, 1, nInfo)
					Call TrDebug_No_Date ("CONFIGURATION UPLOAD " & UCase(strSessionL) & " FAILED ", "", objDebug, MAX_LEN, 3, 1)		
					objTab_L.Session.Disconnect
					bSuccess = False
					Exit Do
			End Select	
			objTab_L.Screen.WaitForString "Name (" & strFTP_ip & ":" & strLogin & "):"
			objTab_L.Screen.Send strFTP_name & chr(13)
			objTab_L.Screen.WaitForString "Password:"
			objTab_L.Screen.Send strFTP_pass & chr(13)
			objTab_L.Screen.WaitForString "ftp>"
			objTab_L.Screen.Send "binary" & chr(13)
			objTab_L.Screen.WaitForString "ftp>"
			'---------------------------------------------
			'   FTP TRANSFER main config
			'---------------------------------------------
			objTab_L.Screen.Send "cd " & strDestFolder & chr(13)
			If Not objTab_L.Screen.WaitForString ("is current directory", 10) Then
				Call  TrDebug_No_Date ("FTP FOLDER " & strDestFolder, "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
				objTab_L.Session.Disconnect
				bSuccess = False
				exit Do		
			End If
			objTab_L.Screen.Send "put " & strConfigFileL & " " & strConfigFileL & chr(13)
			If Not objTab_L.Screen.WaitForString ("Successfully transferred", 10) Then
				Call  TrDebug_No_Date ("FTP Downloading configuration file:", "FAILED", objDebug, MAX_LEN, 1, 1)
				objTab_L.Session.Disconnect
				bSuccess = False
				exit do
			Else
				Call  TrDebug_No_Date ("FTP Downloading configuration file: ", "OK", objDebug, MAX_LEN, 1, 1)  
				bSuccess = True
			End If
			objTab_L.Screen.WaitForString "ftp>"	
			objTab_L.Screen.Send "quit" & chr(13)
			objTab_L.Screen.WaitForString "@" & strHostL & ">"
			Exit Do
		Loop
        If bConnect Then
			objTab_L.Session.Disconnect	
		End If
		If bSuccess Then 
		    Call TrDebug_No_Date ("DOWNLOAD FROM " & UCase(strSessionL), "SUCCESS", objDebug, MAX_LEN, 1, 1)
            '----------------------------------------------------
            '   UPDATE LIST OF NODES IN THE CFG ATTRIBUTES LIST
            '----------------------------------------------------
            NodeCfgExists = False
			For nInd = 0 to UBound(vCfgAttributes) - 1
			   If InStr(vCfgAttributes(nInd),strSessionL) > 0 Then NodeCfgExists = True
            Next
			If Not NodeCfgExists Then 
			    Redim Preserve vCfgAttributes(UBound(vCfgAttributes) + 1)
			    vCfgAttributes(UBound(vCfgAttributes)-1) = "Session " & UBound(vCfgAttributes) - 1 & " = " & strSessionL
				Call AddFileLineInGroup(strDirectoryConfig & "\CfgList.txt", strCfg, vCfgAttributes(UBound(vCfgAttributes)-1),0)
			End If
		Else 
		    Call TrDebug_No_Date ("DOWNLOAD FROM " & UCase(strSessionL), "FAILED", objDebug, MAX_LEN, 1, 1)		
		End If 
    Next
	' Call TrDebug_No_Date ("JOB DONE ", "", objDebug, MAX_LEN, 3, 1)	
	If IsObject(objDebug) Then objDebug.close : End If
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
'   Function TrDebug_No_Date (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function TrDebug_No_Date (strTitle, strString, objDebug, nChar, nFormat, nDebug)
Dim strLine
strLine = ""
If nDebug <> 1 Then Exit Function End If
If IsObject(objDebug) Then 
	Select Case nFormat
		Case 0
			strLine = ""
			strLine = strLine & "  " & strTitle
			strLine = strLIne & strString
			objDebug.WriteLine strLine
			
		Case 1
			strLine = ""
			strLine = strLine & "  " & strTitle
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
'---------------------------------------------------------------------------
'   Function AddFileLineInGroup(strFileName, strGroup1, strParam1,nDebug_)
'---------------------------------------------------------------------------
 Function AddFileLineInGroup(strFileName, strGroup1, strParam1,nDebug)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	AddFileLineInGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	Redim vFileLines(nIndex)
	Call TrDebug_No_Date ("AddFileLineInGroup: String """ & strParam1 & """ under Group [" & strGroup1 & "] WILL BE ADDED", "", objDebug, MAX_LEN, 1, nDebug)					
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
    Do While objDataFileName.AtEndOfStream <> True
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
		Select Case Left(strLine,1)
			Case "["
				Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case Else
							If nGroupSelector = 1 Then 
								nIndex = nIndex - 1
								Do While Len(vFileLines(nIndex)) <= 1
									nIndex = nIndex - 1
								Loop
								nIndex = nIndex + 1
								Redim Preserve vFileLines(nIndex + 1) : vFileLines(nIndex) = strParam1 : nIndex = nIndex + 1
								Redim Preserve vFileLines(nIndex + 1) : vFileLines(nIndex) = " " : nIndex = nIndex + 1
								Call TrDebug_No_Date ("AddFileLineInGroup: String """ & strParam1 & """", "WAS ADDED", objDebug, MAX_LEN, 1, nDebug)					
							End If 
							nGroupSelector = 0
				End Select
				Redim Preserve vFileLines(nIndex + 1)
				vFileLines(nIndex) = strLine
				nIndex = nIndex + 1
			Case Else
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					nIndex = nIndex + 1
		End Select
	Loop
	If nGroupSelector = 1 Then 
		nIndex = nIndex - 1
		Do While Len(vFileLines(nIndex)) <= 1
			nIndex = nIndex - 1
		Loop
		nIndex = nIndex + 1
		Redim Preserve vFileLines(nIndex + 1) : vFileLines(nIndex) = strParam1 : nIndex = nIndex + 1
		Redim Preserve vFileLines(nIndex + 1) : vFileLines(nIndex) = " " : nIndex = nIndex + 1
		Call TrDebug_No_Date ("Adding record to Cfg_List """ & strParam1 & """", "OK", objDebug, MAX_LEN, 1, nDebug)					
	End If 
	objDataFileName.Close
	Call WriteArrayToFile(strFileName,vFileLines, UBound(vFileLines),1,nDebug)
    AddFileLineInGroup = True
End Function
 '#######################################################################
 ' Creates File if it doesn't exists
 ' nMode = 2  Then Append
 ' nMode = 1  Then Rewire all File content
 ' Function WriteArrayToFile - Returns number of lines int the text file
 '#######################################################################
 Function WriteArrayToFile(strFile,vFileLine, nFileLine,nMode,nDebug)
    Dim i, nCount
	Dim strLine
	Dim objDataFileName, objFSO	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.CreateTextFile(strFile)
		If Err.Number = 0 Then 
			objDataFileName.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	
	Select Case nMode
		Case 1 
			Set objDataFileName = objFSO.OpenTextFile(strFile,2,True)
		Case 2 	
			Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	End Select 
	i = 0
	On Error Resume Next
	Err.Clear
	Do While i < nFileLine
		objDataFileName.WriteLine vFileLine(i)
		If Err.Number <> 0 Then 
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T WRITE TO FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			Exit Do 			
		End If
		i = i + 1
	Loop
	On Error Goto 0
	If i = nFileLine Then WriteArrayToFile = True End If
	objDataFileName.close
	Set objFSO = Nothing
End Function
'----------------------------------------------------------------------------
'    Function CreateNewCfg(ByRef strCfg, nCfg, ByRef strVersion, strDirectoryConfig, ByRef vCfgInventory, ByRef vCfgAttributes, nDebug)
'----------------------------------------------------------------------------
Function CreateNewCfg(strCfg, ByRef strVersion, strDirectoryConfig, ByRef vCfgInventory, ByRef vCfgAttributes, nDebug)
    Dim objFSO, vFileLine, nLine, strLine, nVersion, vLine(1), vCfgTempList, nCfg
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	If GetExactObjectLineNumber(vCfgInventory, UBound(vCfgInventory),strCfg) = 0 Then 
	    '----------------------------------------------------
		'   CREATE NEW CFG RECORD AND VERSION
		'----------------------------------------------------
	    nCfg = UBound(vCfgInventory)
		' write new cfg name to CfgList file to the END of the list
		Call WriteStrToFile(strDirectoryConfig & "\CfgList.txt", strCfg, vCfgInventory(nCfg - 1), 3, 0)
		' write new cfg name to CfgInventory Array
		Redim Preserve vCfgInventory(nCfg + 1)
		vCfgInventory(nCfg) = strCfg
		' Create new Version Number
		strVersion = strVersion & "-" & "v.01"
		' write new cfg Version to CfgAttribute Array
		Redim vCfgAttributes(1)
		vCfgAttributes(0) = "Version = " & strVersion
		Redim vFileLine(3)
		vFileLine(0) = " "
		vFileLine(1) = "[" & strCfg & "]"
		vFileLine(2) = "Version = " & strVersion
		Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vFileLine, UBound(vFileLine),2,0)
	Else 
	    '------------------------------------------------------
		'   CREATE NEW VERSION FOR EXISTED CFG
		'------------------------------------------------------
        Call GetFileLineCountByGroup(strDirectoryConfig & "\CfgList.txt", vCfgAttributes,strCfg,"","",0)
		nLine = GetObjectLineNumber(vCfgAttributes, UBound(vCfgAttributes),"Version")
		If UBound(Split(vCfgAttributes(nLine - 1),"v.")) > 0 Then 
		    nVersion = CInt(Split(vCfgAttributes(nLine - 1),"v.")(UBound(Split(vCfgAttributes(nLine - 1),"v.")))) + 1
			If nVersion > 9 Then 
			    strVersion = strVersion & "-" & "v." & nVersion
			    vCfgAttributes(nLine - 1) = vCfgAttributes(nLine - 1) & "," & strVersion
			Else 
			    strVersion = strVersion & "-" & "v.0" & nVersion
			    vCfgAttributes(nLine - 1) = vCfgAttributes(nLine - 1) & "," & strVersion
			End If
		Else 
		   strVersion = strVersion & "-" & "v.01"
           vCfgAttributes(nLine - 1) = "Version = " & strVersion
		End If 
		Call DeleteFileGroup(strDirectoryConfig & "\CfgList.txt", strCfg, 0)
		vLine(0) = "[" & strCfg & "]"
		Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vLine, UBound(vLine),2,0)
		Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vCfgAttributes, UBound(vCfgAttributes),2,0)
	End If 
	'----------------------------------------------------
	'   CREATE CONFIGURATION AND VERSION FOLDERS
	'----------------------------------------------------
	If Not objFSO.FolderExists(strDirectoryConfig & "\" & strCfg) Then 
	    objFSO.CreateFolder  strDirectoryConfig & "\" & strCfg
	End If 
	If Not objFSO.FolderExists(strDirectoryConfig & "\" & strCfg & "\" & strVersion) Then 
	    objFSO.CreateFolder  strDirectoryConfig & "\" & strCfg & "\" & strVersion
	End If 	
	Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------
' Function DeleteFileGroup - Returns number of lines int the text file
'---------------------------------------------------------------------------
 Function DeleteFileGroup(strFileName, strGroup1, nDebug_)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	DeleteFileGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	Redim vFileLines(nIndex)
	If nDebug_ = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": DeleteFileGroup: -------------- NOW READ TO ARRAY -----------" End If  
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
    Do While objDataFileName.AtEndOfStream <> True
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
		Select Case Left(strLine,1)
			Case "["
				Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case Else
							nGroupSelector = 0
							Redim Preserve vFileLines(nIndex + 1) : vFileLines(nIndex) = strLine : nIndex = nIndex + 1
				End Select
			Case Else	
				If nGroupSelector = 0 Then
					Redim Preserve vFileLines(nIndex + 1) : vFileLines(nIndex) = strLine : nIndex = nIndex + 1
					If nDebug_ = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: vFileLines(" & nIndex & "): "  & vFileLines(nIndex) End If  
				End If
		End Select
	Loop
	objDataFileName.Close
	Call WriteArrayToFile(strFileName,vFileLines, UBound(vFileLines),1,nDebug)
    DeleteFileGroup = True
End Function
'---------------------------------------------------------------------------------------
' 	nMode = 2  Then Insert Above
'   nMode = 3  Then Insert Below
' 	nMode = 1  Then Change
'	Inserts or change Line in Text File at String Number "LineNumber" (count form 1)
'   Function WriteStrToFile(strDirectoryTmp & "\" & strFileLocalSessionTmp, nTime, LineNumber, CHANGE)
'---------------------------------------------------------------------------------------
Function WriteStrToFile(strFile, strNewLine, Line, nMode, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine, LineNumber
	Dim objFSO, objFile
	Const FOR_WRITING = 1
	WriteStrToFile = False
'	If LineNumber > 10000 Then objDebug.WriteLine "WriteStrToFile: ERROR: CAN'T OPERATE FILES WITH MORE THEN 10000 STRINGS" End If  
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objFile = objFSO.OpenTextFile(strFile,2,True)
		objFile.WriteLine "Empty"
		If Err.Number = 0 Then 
			objFile.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLIne is number of lines counted like 1,2,...,n
	If Not IsNumeric(Line) Then 
	    LineNumber = GetExactObjectLineNumber(vFileLine, UBound(vFileLine),Line)
	Else 
	    LineNumber = Int(Line)
	End If 
	If nMode = 2 and LineNumber > nFileLine Then nMode = 12 End If
	If nMode = 3 and LineNumber > nFileLine Then nMode = 12 End If	
	If nMode = 1 and LineNumber > nFileLine Then nMode = 12 End If	
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": WriteStrToFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	Select Case nMode
			Case 1 																		' - CHANGE REQUESTED LINENUMBER
					vFileLine(LineNumber - 1) = strNewLine
					If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 2
					Redim vvFileLine(nFileLine + 1)
					For i = 0 to LineNumber - 2
						vvFileLine(i) = vFileLine(i)
					Next
					vvFileLine(LineNumber - 1) = strNewLine
					For i = LineNumber to nFileLine
						vvFileLine(i) = vFileLine(i-1)
					Next
					nFileLine = nFileLine + 1
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 3 ' - Insert After
					Redim vvFileLine(nFileLine + 1)
					For i = 0 to LineNumber - 1
						vvFileLine(i) = vFileLine(i)
					Next
					vvFileLine(LineNumber) = strNewLine
					For i = LineNumber + 1 to nFileLine
						vvFileLine(i) = vFileLine(i-1)
					Next
					nFileLine = nFileLine + 1
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 12
					Redim vvFileLine(LineNumber)
					For i = 0 to nFileLine - 1
						vvFileLine(i) = vFileLine(i)
					Next
					For i = nFileLine to LineNumber - 2
						vvFileLine(i) = " "
					Next
					vvFileLine(LineNumber - 1) = strNewLine
					nFileLine = LineNumber
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
	End Select
End Function
'-------------------------------------------------------------------------
' Function WriteStringToFile - Returns number of lines int the text file
'-------------------------------------------------------------------------
 Function WriteStringToFile(strFile,strLine, nDebug)
	Dim objDataFileName, objFSO	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.CreateTextFile(strFile)
		If Err.Number = 0 Then 
			objDataFileName.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteStringToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteStringToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteStringToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	objDataFileName.WriteLine strLine
	objDataFileName.close
	Set objFSO = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------
' Function returns the number of the line from 1 to N which contains string strObject. Returns 0 if nothing found
'------------------------------------------------------------------------------------------------------------------
Function GetExactObjectLineNumber( byRef vArray, nArrayLen, byRef strObjectName)
Dim nInd
	nInd = 0
	GetExactObjectLineNumber = 0
	Do While nInd < nArrayLen
	If vArray(nInd) = strObjectName Then 
		GetExactObjectLineNumber = nInd + 1
		Exit Do
	End If
	nInd = nInd + 1
    Loop
End Function
'------------------------------------------------------------------------------------------------------------------
' Function returns the number of the line from 1 to N which contains string strObject. Returns 0 if nothing found
'------------------------------------------------------------------------------------------------------------------
Function GetObjectLineNumber( byRef vArray, nArrayLen, byRef strObjectName)
Dim nInd
	nInd = 0
	GetObjectLineNumber = 0
	Do While nInd < nArrayLen
	If InStr(vArray(nInd), strObjectName) <> 0	Then 
		GetObjectLineNumber = nInd + 1
		Exit Do
	End If
	nInd = nInd + 1
    Loop
End Function
