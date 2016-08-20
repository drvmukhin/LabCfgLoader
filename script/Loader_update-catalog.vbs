#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'		JUNIPER MEF CONFIG DOWNLOAD SCRIPT
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 75
' Define global array which stores parameters of all my objects per class
Dim vObjects
' Define global array which keeps properties of all my Classes' 
Dim vClass
' Define global array for JunosSW objects
Dim objMain, objMinor

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
Dim vWaitForCommit, vModels, vWaitForShip,vLoadComplete, vCfgInventory
Dim vSessionCRT, bConnect, vLookForCfg, bSuccess
Dim MainUpdateFlag, MinorUpdateFlag, MsgSuccess

vWaitForShip = Array("ship","]$")
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
	Const DEBUG_FILE = "debug-terminal"
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
			"ARG1: Catalog File Name" & chr(13) &_
			"ARG2: Settings File Name" & chr(13) &_			
			"ARG3: Full Path for Loader Work Folder" & chr(13) &_
		    crt.quit
		Exit Sub
	End If
	strFileSettings = crt.Arguments(1)
	strDirectoryWork = crt.Arguments(2)
	If crt.Arguments.Count = 5 Then
	    MainUpdateFlag = crt.Arguments(3)
		MinorUpdateFlag = crt.Arguments(4)
	Else 
	    MainUpdateFlag = "All"
		MinorUpdateFlag = "All"
    End If
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
strLaunch = strDirectoryWork & "\bin\tail.exe -f " & strDirectoryWork & "\log\" & DEBUG_FILE & ".log"
If Not GetWinAppPID(strPID, strParentPID, DEBUG_FILE, "tail.exe",nDebug) Then 
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
	'   LOAD CATALOG FOR JUNOS S/W
	'--------------------------------------------------------------------------------
    Dim strFileDeviceCatalog, vObjIndex
    strCatalogFile = crt.Arguments(0) 
	Redim vClass(1,1)
	Redim vObjects(1,1,1)
	Redim vObjIndex(1)
	' vClass 
    Call GetMyClass(strCatalogFile, vObjIndex, nDebug)
	ClassName = "JunosSW"
    Call SetMyObject(objMain,"JunosSW",nDebug)	
	Call SetMyObject(objMinor,"Release",nDebug)	
    '--------------------------------
	' BEGIN MAIN CYCLE
	'--------------------------------
    Dim strHostL, strLogin, strSessionL, strFolder, Folder1, Folder2,strTag,strDenyTag,objMainName, MaxMinor
	Dim vTag, vDenyTag
	bFoundConfig = False
	For nSession = 0 to UBound(objMain,1) - 1
	    bSuccess = False
		bConnect = False
	    Do
			If MainUpdateFlag <> "All" and CInt(MainUpdateFlag) <> nSession Then 
			    bSuccess = True
			    MsgSuccess = "SKIP"
			    Exit Do
			End If				
			'--------------------------------------------------------------------------------
			'          GET NAME OF THE TELNET SESSIONS
			'--------------------------------------------------------------------------------
			strSessionL = objMain(nSession,pIndex(0,"SecureCRT_Session"))
			Folder1 = objMain(nSession,pIndex(0,"Folder1"))
			Folder2 = objMain(nSession,pIndex(0,"Folder2"))
			strTag = objMain(nSession,pIndex(0,"Release_Tag"))
			strDenyTag = objMain(nSession,pIndex(0,"Release_Deny_Tag"))
            If strDenyTag = "" or strDenyTag = " " Then strDenyTag = "None"
			If strTag = "" or strTag = " " Then strTag = "None"			
			objMainName = objMain(nSession,pIndex(0,"Name"))
			MaxMinor = objMain(nSession,pIndex(0,"Amount_of_Minors"))
			vTag = Split(strTag,",")
			vDenyTag = Split(strDenyTag,",")
			'------------------------------------------------------------------
			'	Write main variables to log file
			'------------------------------------------------------------------
			Call TrDebug_No_Date ("TelnetScript: " & objMain(nSession,pIndex(0,"Name")) & " : " & strSessionL,"", objDebug, MAX_LEN, 1, nDebug)						
			'--------------------------------------------------------------------------------
			'  Start SSH session to Node
			'--------------------------------------------------------------------------------
			Call TrDebug_No_Date ("START LOADING CATALOGUES FROM " & strSessionL & " FOR " & objMain(nSession,pIndex(0,"Name")),"", objDebug, MAX_LEN, 3, nInfo)						
			On Error Resume Next
			Err.Clear
			Set objTab_L = crt.Session.ConnectInTab("/S " & strSessionL)
			If Err.Number <> 0 Then 
				Call  TrDebug_No_Date ("CAN'T CONNECT TO " & strSessionL & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, nInfo)
				bConnect = False
				bSuccess = False
				Exit Do
			End If
			On Error Goto 0
			bConnect = True
			objTab_L.Caption = Split(strSessionL,"/")(1)
			objTab_L.Screen.Synchronous = True
			'--------------------------------------------------------------------------------
			'  Read Structure of folders on server
			'--------------------------------------------------------------------------------
			' strLine = objTab_L.Screen.ReadString ("]$")
			objTab_L.Screen.Send chr(13)
			objTab_L.Screen.WaitForString ">"
			objTab_L.Screen.Send "bash" & chr(13)	
			objTab_L.Screen.WaitForString ("]$")
			For nRelease = 0 to UBound(Split(objMain(nSession,pIndex(0,"Main List")),","))
		        Do 
				    If MinorUpdateFlag <> "All" and CInt(MinorUpdateFlag) <> nRelease Then Exit Do
					' strMinorList = "Minor List = "
					strMinorList = ""
					strRelease = Split(objMain(nSession,pIndex(0,"Main List")),",")(nRelease)
					Call TrDebug_No_Date("FETCHING for list of images for: " & strRelease & " branch", "" , objDebug, MAX_LEN, 1, nInfo)	
					objTab_L.Screen.Send "cd " & Folder1 & chr(13)	
					objTab_L.Screen.WaitForString ("]$")				
					objTab_L.Screen.Send "cd " & strRelease & "/" & Folder2 & chr(13)	
					objTab_L.Screen.WaitForString ("]$")
					objTab_L.Screen.Send "ls -l" & chr(13)	
					strLine = objTab_L.Screen.ReadString ("[")
					objTab_L.Screen.WaitForString ("]$")
					' Get last five releases
					vLine = Split(strLine,chr(13))
					nLine = 0
					For i = UBound(vLine) to 0 Step -1
						vLine(i) = RTrim(vLine(i)) 
						vLine(i) = RTrim(Split(vLine(i),"->")(0)) 
						vLine(i) = Split(vLine(i)," ")(UBound(Split(vLine(i)," ")))
						Select Case vLine(i)
							Case "current"
								nLine = nLine + 1 : If nLine = int(MaxMinor) Then Exit For
								strMinorList = strMinorList & vLine(i) & ","
							Case Else
								If Len(vLine(i)) > 3  and IsNumeric(Left(vLine(i),1)) and ( strTag = "None" or InStrings(vLine(i),vTag)) and ( Not InStrings(vLine(i),vDenyTag)) Then 
								   strMinorList = strMinorList & vLine(i) & ","
								   nLine = nLine + 1 : If nLine = int(MaxMinor) Then Exit For
								End If 
						End Select
					Next
					' 
					' Validate Minor Release by checking if ship folder exist. If not Exclude it from minor list
					'
					if Len(strMinorList) > 0 Then strMinorList = Left(strMinorList,Len(strMinorList)-1)
					vMinorList = Split(strMinorList,",")
					strMinorList = "Minor List = "
					For Each strMinor in vMinorList
						objTab_L.Screen.Send "ls " & strMinor & "/" & chr(13)
						nResult = objTab_L.Screen.WaitForStrings (vWaitForShip, 20)
						Select Case nResult
							Case 1 
							  strMinorList = strMinorList & strMinor & ","
							  objTab_L.Screen.WaitForString ("]$")
							Case Else
						 End Select				
					Next
					' Delete final coma sign
					if Len(strMinorList) > 0 Then strMinorList = Left(strMinorList,Len(strMinorList)-1)
					' Get ID number of the objMinor by objMinor Name
					objMinorName = objMainName & "-" & strRelease
					Call TrDebug_No_Date("UPDATING  Images list for: " & objMinorName, "" , objDebug, MAX_LEN, 1, nInfo)				
					For nMinor = 0 to UBound(objMinor,1) - 1
					   If objMinorName = objMinor(nMinor,pIndex(1,"Name")) Then Exit For
					Next
					nCount = 0
					' Find [Release_ ] group for the given minor object ID
					For Each strObj in vObjIndex
						If InStr(strObj,"Release_") > 0 Then 
							If nCount = nMinor Then 
							   Call TrDebug_No_Date("GroupName FOUND: nCount = " & nCount & "  " & strObj, "" , objDebug, MAX_LEN, 1, nDebug)				
							   Exit For
							End If
						   nCount = nCount + 1
						End If
					Next 
					Call TrDebug_No_Date("FOUND " & Left(strMinorList,80) & "...", "" , objDebug, MAX_LEN, 1, nInfo)				
					If Not ReplaceFileLineInGroup(strCatalogFile, strObj, "Minor List =", strMinorList,nDebug) Then 
					   MsgBox "Failed to update catalogue file"
					End If
					Exit Do
				Loop
			Next
			bSuccess = True
			MsgSuccess = "SUCCESS"
			Exit Do
		Loop
        If bConnect Then
			objTab_L.Session.Disconnect	
		End If
		If bSuccess Then 
		    Call TrDebug_No_Date (objMain(nSession,pIndex(0,"Name")) , MsgSuccess, objDebug, MAX_LEN, 1, 1)
            '----------------------------------------------------
            '   UPDATE LIST OF MINOR RELEASES
            '----------------------------------------------------
		Else 
		    Call TrDebug_No_Date (objMain(nSession,pIndex(0,"Name")), "FAILED", objDebug, MAX_LEN, 1, 1)		
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
Dim strLine, i
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
'   Function GetWinAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetWinAppPID(ByRef strPID, ByRef strParentPID, strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetWinAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("GetWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("GetWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: CMD: " & process.CommandLine, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: ParentPID:" &  Process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug) 			
			Select Case Lcase(strCommandLine)
			    Case "null", "none", ""
					If pUser = strUser then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			    Case Else
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			End Select
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
'-------------------------------------------------------------------------
' Function AppendStringToFile - Returns number of lines int the text file
'-------------------------------------------------------------------------
 Function AppendStringToFile(strFile,strLine, nDebug)
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
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": AppendStringToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": AppendStringToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			AppendStringToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	objDataFileName.WriteLine strLine
	objDataFileName.close
	Set objFSO = Nothing
End Function
'-------------------------------------------------------------
' Function GetVariable("Param" & i, vCat, 2, nCat)
'-------------------------------------------------------------
Function GetVariable(strVar, vArray, nDim, Dim1, Dim2, nDebug)
	Dim vFileLines, i, nResult
	GetVariable = ""
	nSize = UBound(vArray,nDim)
	Redim vFileLines(nSize)
	Select Case nDim
		Case 1
			For i = 0 to nSize - 1
				vFileLines(i) = vArray(i)
			Next
		Case 2
			For i = 0 to nSize - 1
				vFileLines(i) = vArray(Dim1, i)
			Next
		Case 3
			For i = 0 to nSize - 1
				vFileLines(i) = vArray(Dim1, Dim2,i)
			Next
	End Select
	Do
		For i = 0 to 10
			nResult = GetObjectLineNumber( vFileLines, nSize, strVar & Space(i) & "=")
			If nResult > 0 Then Exit For End If
		Next
		If nResult = 0 Then 
			GetVariable = "NULL" 
			Call TrDebug_No_Date ("GetVariable: CAN'T FIND VARIABLE: " & strVar, "Dim1=" & Dim1 & " Dim2=" & Dim2, objDebug, MAX_LEN, 1, nDebug)
			Exit Do 
		End If
		If nResult > 0 and InStr(vFileLines(nResult - 1),"=") = 0 Then 
			GetVariable = "NULL" 
			Call TrDebug_No_Date ("GetVariable: ERROR: WRONG DEFINITION OF THE VARIABLE " & strVar, "", objDebug, MAX_LEN, 1, 1)
			Exit Do
		End If
			GetVariable = LTrim(RTrim(Split(vFileLines(nResult - 1),"=")(1)))
			Exit Do
	Loop
End Function 
'------------------------------------------------------------------------------
' Function GetMyClass
'------------------------------------------------------------------------------
Function GetMyClass(strFileName, ByRef vObjIndex, nDebug)
    Dim nIndex, i
	Dim strLine 
	Dim nGroupSelector, nClass
	Dim vFileLines, nFileLines, vClassIndex
	Const MAX_PARAM = 1
	
	GetMyClass = 0
	nGroupSelector = 0
	nMaxObjects = 0
	Call GetFileLineCountByGroup(strFileName , vObjIndex,"Object_Index","","",0)
	nClass = GetFileLineCountByGroup(strFileName , vClassIndex,"Classes","","",0)
	Redim vClass(nClass, MAX_PARAM)
	nFileLines = GetFileLineCountSelect(strFileName, vFileLines,"#","NULL","NULL",0)
	'-----------------------------------------------------
	'	COUNT Named Objects in Each Class
	'-----------------------------------------------------
	For n = 0 to nClass - 1
		nObj = 0
		For i = 0 to nFileLines - 1
			If InStr(vFileLines(i), "[" & vClassIndex(n) & "_") > 0 Then nObj = nObj + 1 End If
		Next
		vClass(n,0) = nObj
		Call TrDebug_No_Date ("GetMyClass: TOTAL OBJECTS IN CLASS: " & vClassIndex(n), nObj, objDebug, MAX_LEN, 1, nDebug)
	Next
	'-----------------------------------------------------
	'	FIND THE MAXIMUM NUMBER OF ALL OBJECTS
	'-----------------------------------------------------
	For n = 0 to nClass - 1
		nMaxObjects = MaxQ(nMaxObjects, vClass(n,0)) 		
	Next
	'-----------------------------------------------------
	'	LOAD CLASSES PROPERIES FROM CATALOG
	'-----------------------------------------------------
		nGroupSelector = 0
		For n = 0 to nClass - 1
			For nIndex = 0 to nFileLines - 1
				strLine = LTrim(vFileLines(nIndex))
				Select Case Left(strLine,1)
					Case "#"
					Case ""
					Case "["
						Select Case strLine
							Case "[Class_" & vClassIndex(n) & "]"
								Call TrDebug_No_Date ("GetMyClass: LOAD PROPERTIES FOR [Class_" & vClassIndex(n) & "]", "", objDebug, MAX_LEN, 3, nDebug)
								nParam = 1
								nGroupSelector = 1
							Case Else
								nGroupSelector = 0
						End Select
					Case Else	
						If nGroupSelector = 1 Then 
							Call TrDebug_No_Date ("GetMyClass:" & strLine, "", objDebug, MAX_LEN, 1, nDebug)					
							If nParam + 1 > UBound(vClass,2) Then Redim Preserve vClass(nClass,nParam + 1)
							vClass(n,nParam) = strLine
							nParam = nParam + 1 
						End If
				End Select
			Next
		Next
	'-----------------------------------------------------
	'	DEFINE vObjects Arrey
	'-----------------------------------------------------
	Redim vObjects(nClass, nMaxObjects, UBound(vClass,2))		
	'-----------------------------------------------------
	'	LOAD ALL CURRENT USER OBJECTS PROPERIES
	'-----------------------------------------------------
	For n = 0 to nClass - 1
		Do 
			If vClass(n,0) = 0 Then Exit Do End If
			Redim vIndex(0) : nCount = 0
			'-----------------------------------------------------
			'	CREATE OBJECT INDEX FOR THE CLASS
			'-----------------------------------------------------
			For nInd = 0 to UBound(vObjIndex)
				If InStr(vObjIndex(nInd),vClassIndex(n))<> 0 Then 
					nCount = nCount + 1
					Redim Preserve vIndex(nCount)
					vIndex(nCount - 1) = vObjIndex(nInd)
				End If
			Next
			'-----------------------------------------------------
			'	LOAD OBJECT PARAMETERS FROM CATALOG
			'-----------------------------------------------------
			For i = 0 to vClass(n,0) - 1
				nGroupSelector = 0
				For nIndex = 0 to nFileLines - 1
					strLine = LTrim(vFileLines(nIndex))
					Select Case Left(strLine,1)
						Case "#"
						Case ""
						Case "["
							Select Case strLine
								Case "[" & vIndex(i) & "]"
									Call TrDebug_No_Date ("GetMyClass: LOAD DATA FOR: " & strLine, "", objDebug, MAX_LEN, 3, nDebug)
									nParam = 0
									nGroupSelector = 1
								Case Else
									' If nGroupSelector = 1 Then Exit For
									nGroupSelector = 0
							End Select
						Case Else	
							If nGroupSelector = 1 Then 
								Call TrDebug_No_Date ("GetMyClass:" & strLine, "", objDebug, MAX_LEN, 1, nDebug)					
								vObjects(n,i,nParam) = strLine
								nParam = nParam + 1 
							End If
					End Select
				Next
			Next
			Exit Do
		Loop
	Next
	GetMyClass = nClass
End Function
'----------------------------------------------------------
'    Function pIndex(Byref vClass,strClassID,ParamName)
'----------------------------------------------------------
Function pIndex(strClassID,ParamName)
  Dim nIndex, ClassID, strLine
    pIndex = -1
    If Not IsNumeric(strClassID) Then 
	    For nIndex = 0 to UBound(vClass,1) - 1
	       If InStr(vClass(nIndex,1),strClassID) <> 0 Then  ClassID = nIndex
	    Next
	Else 
	   ClassID = Int(strClassID)
	End If 
	Call TrDebug_No_Date ("pIndex: ClassName: " & strClassID & " ClassID = " & ClassID, "", objDebug, MAX_LEN, 1, nDebug)					
    nIndex = 1
	For i = 0 to UBound(vClass,2)
	    strLine = vClass(ClassID,i)
		If InStr(strLine, "Param") <> 0 Then 
		   If InStr(strLine, ParamName) <> 0 Then pIndex = nIndex - 1 : Exit For : End If
		   nIndex = nIndex + 1
		End If
	Next
End Function
'----------------------------------------------------------
'    Function SetMyObject(ByRef objDevices, ByRef vObjects, ByRef vClass, strClassID)
'----------------------------------------------------------
Function SetMyObject(ByRef objDevices, strClassID, nDebug)
  Dim nIndex, ClassID, nObj, i
    '-------------------------------------
	'   GET CLASS ID 
	'-------------------------------------
    If Not IsNumeric(strClassID) Then 
	    For nIndex = 0 to UBound(vClass,1) - 1
	       If InStr(vClass(nIndex,1),strClassID) <> 0 Then  ClassID = nIndex
	    Next
	Else 
	   ClassID = Int(strClassID)
	End If 
	Call TrDebug_No_Date ("SetMyObject: ClassName: " & strClassID & " ClassID = " & ClassID, "", objDebug, MAX_LEN, 1, nDebug)	
	'---------------------------------------
	'   COUNT Params in Given Class
	'---------------------------------------
	nParam = 0
	For i = 0 to UBound(vClass,2)-1
		If InStr(vClass(ClassID,i),"Param") > 0 Then 
			nParam = nParam + 1 
		End If
	Next
	Call TrDebug_No_Date ("SetMyObject: Found " & nParam & " Properties for class " & strClassID,"", objDebug, MAX_LEN, 1, nDebug)
	'---------------------------------------
	'   COPY PROPERTIES VALUES TO THE OBJECT
	'---------------------------------------
	Redim objDevices(vClass(ClassID,0),nParam)
    For nObj = 0 to vClass(ClassID,0) - 1
	    For i = 1 to nParam
		    Param_Name = GetVariable("Param" & i, vClass, 2, ClassID, 0, nDebug)
			Param_Value = GetVariable(Param_Name, vObjects, 3, ClassID, nObj, nDebug)
		    objDevices(nObj,i-1) = Param_Value
        Next
	Next
End Function
'-----------------------------------------------------
'   GetAmountOfProperties(vClass)
'-----------------------------------------------------
Function GetAmountOfProperties(vClass, strClassID)
  Dim nParam, i, nIndex, ClassID
    '-------------------------------------
	'   GET CLASS ID 
	'-------------------------------------
    If Not IsNumeric(strClassID) Then 
	    For nIndex = 0 to UBound(vClass,1) - 1
	       If InStr(vClass(nIndex,1),strClassID) <> 0 Then  ClassID = nIndex
	    Next
	Else 
	   ClassID = Int(strClassID)
	End If   
	nParam = 0
	For i = 0 to UBound(vClass,2)-1
		If InStr(vClass(ClassID,i),"Param") > 0 Then 
			nParam = nParam + 1 
		End If
	Next
	GetAmountOfProperties = nParam
End Function 
'###################################################################################
'  Function MinQ - Returs the Minimum of two numeric values                  
'###################################################################################
Function MinQ( nA, nB)
   If nA < nB Then 
     MinQ = nA 
   Else 
     MinQ = nB
   End If
End Function
'###################################################################################
'  Function MinQ - Returs the Minimum of two numeric values                  
'###################################################################################
Function MaxQ( nA, nB) 
   If nA > nB Then 
     MaxQ = nA 
   Else 
     MaxQ = nB
   End If
End Function
'#######################################################################
' Function ReplaceFileLineInGroup - Returns number of lines int the text file
'#######################################################################
 Function ReplaceFileLineInGroup(strFileName, strGroup1, strParamOld, strParam1,nDebug)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	ReplaceFileLineInGroup = False
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	Redim vFileLines(nIndex)
	Call TrDebug_No_Date ("ReplaceFileLineInGroup: String """ & strParam1 & """ under Group [" & strGroup1 & "] WILL BE ADDED", "", objDebug, MAX_LEN, 1, nDebug)					
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
				End Select
				Redim Preserve vFileLines(nIndex + 1)
				vFileLines(nIndex) = strLine
				nIndex = nIndex + 1
			Case Else
			        Select Case nGroupSelector
					    Case 1
					        If InStr(strLine, strParamOld) > 0 Then strLine = strParam1
							ReplaceFileLineInGroup = True
					    Case 0
					End Select
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					nIndex = nIndex + 1
		End Select
	Loop
	objDataFileName.Close
	Call WriteArrayToFile(strFileName,vFileLines, UBound(vFileLines),1,nDebug)
    ReplaceFileLineInGroup = True
End Function
'----------------------------------------------------
'  Function InStrings(String1, String2)
'----------------------------------------------------
Function InStrings(String1, vString2)
Dim strLine
    InStrings = False
	For each strLine in vString2
	    If strLine = "" Then Exit For
	    If Instr(String1,strLine) Then 
	        Instrings = True
			Exit For
	    End If
	Next
End Function