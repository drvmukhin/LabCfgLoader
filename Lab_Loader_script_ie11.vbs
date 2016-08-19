'----------------------------------------------------------------------------------
'            VASILY LAB CONFIGURATION LOADER
'----------------------------------------------------------------------------------
Const LDR_SCRIPT_NAME = "Lab_Loader_"
Const ForAppending = 8
Const ForWriting = 2
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
Const PAUSE_TIME = 		20
Const MAX_TIMER = 		20
Const MAX_TIME_DELTA = 	180
Const MAX_LEN = 		134
Const nABC =			2
Const PARENTS = 		"1"
Const DEBUG_FILE = "debug-loader"
' Define global array which stores parameters of all my objects per class
Dim vObjects, vObjIndex
' Define global array which keeps properties of all my Classes' 
Dim vClass
' Define global array for JunosSW objects
Dim objMain, objMinor 
Dim D0 
Dim nResult, nTail, nIndex, nCountWeek, nCountMonth
Dim strLine
Dim strFileSessionTmp 
Dim strDirectoryWork, strCRT_InstallFolder, strDirectoryTmp, strCRTexe, SecureCRT, nWindowState, strCRT_ConfigFolder, strCRT_SessionFolder
Dim nDebug, nLine, nDebugCRT
Dim intX, intY
Dim nSession, nSessionTmp, nInventory
Dim vSession, vSessionTmp, vMsg(20), vSessionCRT, vSessionEnable
Dim vLine
Dim vvMsg(20,3)
Dim strFolder
Dim objFSO, objEnvar, objDebug, objShell
Dim vIE_Scale
Dim UserConfigFile
Dim objShellApp
Dim strPID
'----------------------------------
' NEW VARIABLES
'----------------------------------
Dim vFlavors, vNodes(2,5), vTemplates(4), vSettings
Dim strConfigFileL, strConfigFileR, strVersion
Dim Platform, DUT_Platform
Dim VBScript_DNLD_Config, VBScript_Upload_Config, VBScript_FWF_Apply, VBScript_FTP_User, VBScript_Set_Node, VBScript_BLK_DNLD_Config, VBScript_UPDATE_Catalog, VBScript_UPDATE_Junos
Dim strTempOrigFolder, strTempDestFolder, vXLSheetPrefix(4)
Dim objXLS, objWrkBk, objXLSeet
Dim vDelim, vParamNames,vPlatforms, objWMIService, IE_Window_Title
Dim objFolder, colFiles, IPConfigSet, strEditor, SecureCRT_Installed, FileZilla_Installed
    Const SECURECRT_FOLDER = "SecureCRT Folder"
    Const WORK_FOLDER = "Work Folder"
    Const CONFIGS_FOLDER = "Configuration Files Folder"
    Const CONFIGS_PARAM  = "MEF Service Parameters"
    Const CONFIGS_GLOBAL  = "CONFIGS_GLOBAL"
    Const CONFIGS_RE0  = "CONFIGS_RE0"
    Const CONFIGS_RE1  = "CONFIGS_RE1"
    Const Node_Left_IP  = "Left Node IP"
    Const Node_Right_IP  = "Right Node IP"
	Const HIDE_CRT = "Hide Terminal Session"
    Const FTP_IP  = "FTP IP"
    Const FTP_User  = "FTP User"
    Const FTP_Password  = "FTP Password"
	Const LAN_ADAPTER = "Network Adapter"
	Const PLATFORM_NAME = "Platform Name"
	Const PLATFORM_INDEX = "Node Name Prefix"
	Const Template = "XLS TEMPLATE"
	Const Orig_Folder = "Original TCG Templates"
	Const Dest_Folder = "Exported TCG Templates"
	Const WorkBookPrefix = "WorkBookPrefix"
	Const SECURECRT_L_SESSION = "Left Node Session"
	Const SECURECRT_R_SESSION = "Right Node Session"
	Const SECURECRT_SESSION = "Node Session"
	Const SAVE_AS = "Save New Configuration As..."
    Const LAST_BLK_FOLDER = "Blk Config File Folder"
	Const MAIN_TITLE = "Juniper Networks Lab Configuration Loader"
    Const CRT_REG_INSTALL = "HKEY_LOCAL_MACHINE\SOFTWARE\VanDyke\SecureCRT\Install\Main Directory"
    Const CRT_REG_CONFIG = "HKEY_CURRENT_USER\Software\VanDyke\SecureCRT\Config Path"
	Const LAB_CFG_LDR_REG = "HKEY_CURRENT_USER\Software\JnprLabCfgLdr\Main Directory"
	Const FTP_REG_INSTALL = "HKEY_CURRENT_USER\Software\FileZilla Server\Install_Dir"
	Const NOTEPAD_PP = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe\"
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	
ReDim vParamNames(30)
    vParamNames(0) = HIDE_CRT
    vParamNames(1) = LAN_ADAPTER
    vParamNames(2) = FTP_IP
    vParamNames(3) = FTP_User
    vParamNames(4) = FTP_Password
	vParamNames(5) = SECURECRT_FOLDER
    vParamNames(6) = WORK_FOLDER
    vParamNames(7) = CONFIGS_FOLDER
    vParamNames(8) = Orig_Folder
	vParamNames(9) = Dest_Folder
	vParamNames(10) = "Not Used"
	vParamNames(11) = "Not Used" 
    vParamNames(12) = CONFIGS_PARAM
	vParamNames(13) = PLATFORM_NAME
    vParamNames(14) = PLATFORM_INDEX
    vParamNames(25) = LAST_BLK_FOLDER	
	
	
vDelim = Array("=",",",":")
ReDim vSession (10)
Redim vSettings(30)
For nInd = 0 to UBound(vParamNames)
	vSettings(nInd) = vParamNames(nInd) & "=Unknown"
Next
vSession(1) = " "
strFileSessionTmp = "sessions_tmp.dat"
strDirectoryWork = "C:\Users\vmukhin\Documents\Products\Juniper\_SOLUTION_TEAM_ALL_\A&A\Fortius\MEF-Certification\ACX5K\MefCfgLoader"
strDirectoryBackUp = ""
strDirectoryConfig = ""
strCRT_InstallFolder = "C:\Program Files\VanDyke Software\SecureCRT"
SecureCRT = "SecureCRT.exe"
strCRTexe = "\SecureCRT.exe"""
strFileParam = "configurations.dat"
strFileSettings = "settings.dat"
strConfigFileL = ""
strConfigFileR = ""
Platform = "acx"
DUT_Platform = "Unknown"
D0 = DateSerial(2015,1,1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = WScript.CreateObject("WScript.Shell")
Set objShell = WScript.CreateObject("WScript.Shell")
nDebugCRT = 0
nDebug = 1
CurrentDate = Date()
CurrentTime = Time()
D0 = DateSerial(2015,1,1)

Main()
  
If IsObject(objDebug) Then objDebug.Close : End If
Set objFSO = Nothing
set objEnvar = Nothing
Set objShell = Nothing

Sub Main()
	'-----------------------------------------------------------------
	'  GET SCREEN RESOLUTION
	'-----------------------------------------------------------------
		Call GetScreenResolution(vIE_Scale, 0)
'		If nDebugCRT = 1 Then strCRTexe = strCRTexe & " /POS 10 11" Else strCRTexe = strCRTexe & " /POS 10 1080" End If
	'-------------------------------------------------------------------------------------------
	'  CHECK LCL FOLDER
	'-------------------------------------------------------------------------------------------
	nResult = 0
	On Error Resume Next
		Err.Clear
		strDirectoryWork = objShell.RegRead(LAB_CFG_LDR_REG)
		if Err.Number <> 0 Then 
			strDirectoryWork = "Not Set"
		End If
	On Error Goto 0
Do
' MsgBox "Step:" & nResult 
	Select Case nResult
		Case 0 ' - Check if WorkFolder Exists
				If Not objFSO.FolderExists(strDirectoryWork) Then 
				    nResult = 1 
					vvMsg(0,0) = "Lab Configuration Loader"		 				: vvMsg(0,1) = "Bold" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "Set a Local Folder"							: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1 
					vvMsg(2,0) = " "											: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor1 
					nLine = 3
				Else 
				    nResult = 2 
				End If	    
		Case 1  ' - WorkFolder Doesn't exist or script file was not found in folder
	            if 	strDirectoryWork = "Not Set" Then strDirectoryWork = "C:\"
				if Not IE_DialogFolder (vIE_Scale, "Working Folder of the Loader", strDirectoryWork, vvMsg, nLine, 0) then
					exit sub
				Else 
					nResult = 3
				End If
		Case 2  ' - Check if file exist in folder
        		Set objFolder = objFSO.GetFolder(strDirectoryWork)
        		Set colFiles = objFolder.Files
				For Each objFile in colFiles
					strFile = objFile.Name
					If InStr(LCase(strFile),LCase(LDR_SCRIPT_NAME)) Then nResult = 5	End If 
				Next
				If nResult <> 5 Then
					vvMsg(0,0) = "Lab Configuration Loader"		 				: vvMsg(0,1) = "Bold" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "Can't find Lab_Loader script in the folder"	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2 				
					vvMsg(2,0) = "Set a Working Folder where script located"    : vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor1 
					vvMsg(3,0) = " "											: vvMsg(3,1) = "normal" : vvMsg(3,2) = HttpTextColor1 
					nLine = 4
					nResult = 1
				End If
			    Set objFolder = Nothing
				Set colFiles = Nothing
		Case 3  ' - After new folder selected update Working folder in the script file
				On Error Resume Next
				Err.Clear
				objShell.RegWrite LAB_CFG_LDR_REG, strDirectoryWork, "REG_SZ"
				if Err.Number <> 0 Then 
					MsgBox "Error: Can't Right to Windows Registry" & chr(13) & Err.Description
					Exit Sub
				Else 
					On Error Goto 0
					Set objFolder = Nothing
					Set colFiles = Nothing
					Exit Do
				End If
		Case 4
		Case 5 ' Working folder was successfully set
			Exit Do
    End Select		   
Loop
'-------------------------------------------------------------------------------------------
'  OPEN LOG FILE
'-------------------------------------------------------------------------------------------
	If Not objFSO.FolderExists(strDirectoryWork & "\Log") Then 
			objFSO.CreateFolder(strDirectoryWork & "\Log") 
	End If
'-----------------------------------------------------------------
'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	
'-----------------------------------------------------------------
'  	CHECK IF START SCRIPT IS ALREADY RUNNING AND OPEN LOG FILE
'-----------------------------------------------------------------
	On Error Resume Next
	Set objDebug = objFSO.OpenTextFile(strDirectoryWork & "\Log\" & DEBUG_FILE & ".log",ForWriting,True)
	Select Case Err.Number
		Case 0
		Case 70
		    Do 
				Err.Clear
				'----------------------------------------------------
				'  GET MAIN FORM PID
				'----------------------------------------------------
				strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & MAIN_TITLE & " - " & IE_Window_Title & """"
				Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, 0)
			    strPID = ""
				For Each strLine in vCmdOut
				   If InStr(strLine,"iexplore.exe") then strPID = Split(strLine,""",""")(1)
				Next
				If strPID <> "" Then 
				     FocusToParentWindow(strPID)
					Exit Sub
				Else 
					vvMsg(0,0) = "IT SEAMS THAT ONE INSTANCE OF CONFIGURATION LOADER IS ALREADY RUNNING"	: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "Exit . . ."           									: vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2 
					Call IE_MSG(vIE_Scale, "Error",vvMsg,2, "Null")
					Exit Sub
				End If
			Loop
		Case Else 
			vvMsg(0,0) = "SOMETHING GOING WRONG. CAN'T LAUNCH THE SCRIPT " 						: vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Exit . . ."									: vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor1
 			Call IE_MSG(vIE_Scale, "Error",vvMsg,2,"Null")
			Exit Sub
	End Select
	On Error goto 0
	'-----------------------------------------------------------------
	'  CHOOSE A DEFAULT TEXT EDITOR
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		strEditor =  """" & objShell.RegRead(NOTEPAD_PP) & """"
		if Err.Number <> 0 Then 
			strEditor = "notepad.exe"
		End If
	On Error Goto 0
	'-------------------------------------------------------------------------------------------
	'  	CHECK FILEZILLA SERVER INSTALLED
	'-------------------------------------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		FileZilla_Installed = True
		strFTP_Folder = objShell.RegRead(FTP_REG_INSTALL)
		if Err.Number <> 0 Then 
			vvMsg(0,0) = "WARNING: CAN'T FIND FileZilla Server Folder"	           : vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Make sure that FileZilla Server Installed correctly"     : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1			
			vvMsg(2,0) = "If You are using other FTP Server " & _
                		 "make sure that FTP Login/Password are " & _
						 "the same as under MEF Loader Settings. " & _
						 "Folder with DUT Configuration files should " & _
						 "be used as FTP User Homedirectory"                      	: vvMsg(2,1) = "normal" : vvMsg(1,2) = HttpTextColor1						
			vvMsg(3,0) = "" : vvMsg(4,0) = "" :vvMsg(5,0) = "" :vvMsg(6,0) = "" :vvMsg(7,0) = "" :
			Call IE_MSG(vIE_Scale, "Error",vvMsg,8,"Null")
			FileZilla_Installed = False
			strFTP_Folder = "C:"
		End If
		If Right(strFTP_Folder,1) = "\" Then strFTP_Folder = Left(strFTP_Folder,Len(strFTP_Folder)-1)
	On Error Goto 0
	'-----------------------------------------------------------------
	'  CHECK SECURECRT IS INSTALLED ON THE SYSTEM
	'-----------------------------------------------------------------
	On Error Resume Next
	    SecureCRT_Installed = True
		Err.Clear
		strCRT_InstallFolder = objShell.RegRead(CRT_REG_INSTALL)
		if Err.Number <> 0 Then 
			vvMsg(0,0) = "WARNING: CAN'T FIND SecureCRT Install Folder"	: vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Make sure that Secure CRT Application was " & _
			              "installed on your system correctly"	        : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2			
			vvMsg(2,0) = ""         									: vvMsg(2,1) = "bold" : vvMsg(2,2) = HttpTextColor1
			Call IE_MSG(vIE_Scale, "Error",vvMsg,2,"Null")
            SecureCRT_Installed = False
			strCRT_InstallFolder = "C:"
		End If
		If Right(strCRT_InstallFolder,1) = "\" Then strCRT_InstallFolder = Left(strCRT_InstallFolder,Len(strCRT_InstallFolder)-1)
		strCRT_ConfigFolder = objShell.RegRead(CRT_REG_CONFIG)
		if Err.Number <> 0 Then 
			vvMsg(0,0) = "WARNING: CAN'T FIND SecureCRT Session Folder"	              : vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Once SecureCRT application was installed run it, " & _
			             "so that default session configuration will be created"	  : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
			vvMsg(2,0) = "Exit . . ."									              : vvMsg(2,1) = "bold" : vvMsg(1,2) = HttpTextColor1
			Call IE_MSG(vIE_Scale, "Error",vvMsg,3,"Null")
            SecureCRT_Installed = False
			strCRT_ConfigFolder = "C:"
		End If
		If Right(strCRT_ConfigFolder,1) = "\" Then strCRT_ConfigFolder = Left(strCRT_ConfigFolder,Len(strCRT_ConfigFolder)-1)
		strCRT_SessionFolder = strCRT_ConfigFolder & "\Sessions"
	On Error Goto 0
	
'-----------------------------------------------------
'  	LOAD INITIAL CONFIGURATION FROM SETTINGS FILE
'-----------------------------------------------------
	If objFSO.FileExists(strDirectoryWork & "\config\" & strFileSettings) Then 
		nSettings = GetFileLineCountByGroup(strDirectoryWork & "\config\" & strFileSettings, vLines,"Settings","","",0)
		strFileSettings= strDirectoryWork & "\config\" & strFileSettings
		For nInd = 0 to nSettings - 1 
		    ' vSettings(26) Reserved for temporary value when creating new configuration
		    ' vSettings(25) Used for the name of the file with list of the configurations for bulk load from nodes			
			Select Case Split(vLines(nInd),"=")(0)
					Case SECURECRT_FOLDER
								vSettings(5) = vLines(nInd)
								strCRT_InstallFolder = Split(vLines(nInd),"=")(1)
								vSettings(5) = vParamNames(5) & "=" & strCRT_InstallFolder
					Case WORK_FOLDER
								vSettings(6) = WORK_FOLDER & "=" & strDirectoryWork
'								strDirectoryWork = strDirectoryWork
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
					Case LAST_BLK_FOLDER
   					            vSettings(25) = vLines(nInd)					
					Case HIDE_CRT
								vSettings(0) = vLines(nInd)
							    Select Case Split(vLines(nInd),"=")(1)
							        Case "1"
                                        nWindowState = 2
								    Case Else
									    nWindowState = 1
							    End Select
					Case LAN_ADAPTER
								vSettings(1) = vLines(nInd)
								strEth =  Split(vLines(nInd),"=")(1)
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
	Else 
	    MsgBox "Error: Can't find settings file: " & chr(13) & strDirectoryWork & "\config\" & strFileSettings
		Exit Sub
	End If
	vSettings(15) = " =" & strCRT_SessionFolder
	vSettings(16) = " =" & strFTP_Folder
	strDirectoryTmp = strDirectoryWork
	strWinUtilsFolder = strDirectoryWork & "\Bin"
	strCRTexe = """" & strCRT_InstallFolder & strCRTexe
	Dim vsshFile, LineNumber
	If objFSO.FileExists(strCRT_ConfigFolder & "\SSH2.ini") Then
	    Call GetFileLineCountSelect(strCRT_ConfigFolder & "\SSH2.ini",vsshFile,"NULL","NULL","NULL",0)
        LineNumber = GetObjectLineNumber( vsshFile, UBound(vsshFile), "Host Key Database Location")		
		If LineNumber > 0 Then 
		   vSettings(17) = Split(vsshFile(LineNumber - 1),"=")(1)
		Else 
		   vSettings(17) = strCRT_ConfigFolder
		End If 
	End If 
'--------------------------------------------------------------------------------
'   CHECK FTP IP ADDRESS = LAN ADAPTER IP
'--------------------------------------------------------------------------------
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each IPConfig in IPConfigSet
	    If Split(vSettings(1),"=")(1) = IPConfig.Description Then 
			vSettings(2) = FTP_IP & "=" & IPConfig.IPAddress(0)
		End If
	Next
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
'--------------------------------------------------------------------------------
'          GET STATUS OF THE SESSIONS
'--------------------------------------------------------------------------------
	Call GetFileLineCountByGroup(strFileSettings, vSessionEnable,"Active_Sessions","","",0)
'--------------------------------------------------------------------------------
'          LOAD CONFIGURATION LIST
'--------------------------------------------------------------------------------
    Dim vCfgInventory, vCfg, vCfgList, strCfg, nCfg, strCfgFile
	strCfgFile = strDirectoryConfig & "\CfgList.txt"
	nCount = GetFileLineCountByGroup(strCfgFile, vCfgInventory,"Inventory","","",0)
	Redim vCfgList(nCount,UBound(vSessionCRT) + 2)
	nCfg = 0 
	For Each strCfg in vCfgInventory
	   If strCfg = "" Then Exit For
	   vCfgList(nCfg,0) = strCfg
	   nCount = GetFileLineCountByGroup(strCfgFile, vCfg, strCfg,"","",0)
	   nInd = 1
	   For Each LineItem in vCfg
	      If LineItem = "" Then Exit For
	      vCfgList(nCfg,nInd) = LineItem
		  nInd = nInd + 1
	   Next
	   nCfg = nCfg + 1
	Next
'--------------------------------------------------------------------------------
'          GET NAME OF THE TELNET SCRIPTS
'--------------------------------------------------------------------------------
	nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Scripts","","",0)
	For nIndex = 0 to nInventory - 1
		Select Case Split(vLines(nIndex),"=")(0)
			Case "UPLOAD"
				VBScript_Upload_Config = Split(vLines(nIndex),"=")(1)
			Case "DOWNLOAD"
			    VBScript_DNLD_Config = Split(vLines(nIndex),"=")(1)
			Case "DNLD_EXISTED"
			    VBScript_BLK_DNLD_Config = Split(vLines(nIndex),"=")(1)
			Case "UPDATE_CATALOG"
			    VBScript_UPDATE_Catalog = Split(vLines(nIndex),"=")(1)
			Case "UPDATE_JUNOS"
			    VBScript_UPDATE_Junos = Split(vLines(nIndex),"=")(1)
			Case "FWFILTER"
			    VBScript_FWF_Apply = Split(vLines(nIndex),"=")(1)
			Case "FTPUSER"
			    VBScript_FTP_User = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "SETNODE"
				VBScript_Set_Node = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
		End Select
	Next
'--------------------------------------------------------------------------------
'          GET NAME OF THE TELNET SCRIPTS
'--------------------------------------------------------------------------------
	nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Version","","",0)
	For nIndex = 0 to nInventory - 1
		Select Case Split(vLines(nIndex),"=")(0)
			Case "VERSION"
				strVersion = Split(vLines(nIndex),"=")(1)
		End Select
	Next
'-------------------------------------------------------------------------------------------
'  		LOAD TEMPLATES PARAMETERS 
'-------------------------------------------------------------------------------------------
	nCount = GetFileLineCountByGroup(strFileSettings, vLines,"Templates","","",1)
	nInd = 0 : nInd1 = 0
	For n = 0 to nCount - 1
		Select Case Split(vLines(n),"=")(0)
			Case Template
				vTemplates(nInd) = Split(vLines(n),"=")(1) : nInd = nInd + 1
				Call TrDebug("GetTemplates: ", vTemplates(nInd-1), objDebug, MAX_LEN, 1, 1)
			Case Orig_Folder
			    strTempOrigFolder = Split(vLines(n),"=")(1)
				vSettings(8) = vLines(n)
				Call TrDebug("GetTemplates: ", strTempOrigFolder, objDebug, MAX_LEN, 1, 1)
			Case Dest_Folder
			    strTempDestFolder = Split(vLines(n),"=")(1)
				vSettings(9) = vLines(n)
                Call TrDebug("GetTemplates: ", strTempDestFolder, objDebug, MAX_LEN, 1, 1)				
			Case WorkBookPrefix
			    vXLSheetPrefix(nInd1) = Split(vLines(n),"=")(1) : nInd1 = nInd1 + 1
				Call TrDebug("GetTemplates: ", vXLSheetPrefix(nInd1-1), objDebug, MAX_LEN, 1, 1)				
		End Select
	Next
'--------------------------------------------------------------------------------
'          GET LIST OF SUPPORTED PLATFORMS
'--------------------------------------------------------------------------------
	nPlatform = GetFileLineCountByGroup(strFileSettings, vPlatforms,"Supported_Platforms","","",0)
'-------------------------------------------------------------------------------------------
'        SET SERVICE PARAM FULL PATH
'-------------------------------------------------------------------------------------------
	strFileParam = strDirectoryWork & "\config\" & strFileParam
'--------------------------------------------------------------------------------
'   LOAD CATALOG FOR JUNOS S/W
'--------------------------------------------------------------------------------
    Dim strFileDeviceCatalog
    strFileDeviceCatalog = strDirectoryWork & "\config\class_catalog.dat"
	Redim vClass(1,1)
	Redim vObjects(1,1,1)
	Redim vObjIndex(1)
	' vClass 
    Call GetMyClass(strFileDeviceCatalog, vObjIndex, nDebug)
    Call SetMyObject(objMain,"JunosSW",nDebug)
	Call SetMyObject(objMinor,"Release",nDebug)
'	Call TrDebug ("CHECK objDevice Data: ","", objDebug, MAX_LEN, 3, 1)
'   Call GetDevicesStatus(objDevices, vClass, vAccount, vDevice, strSrvDirectory,False, 1)	
'    MsgBox vClass(1,0) & ", " & vClass(1,2) & ", " & vClass(1,3) & chr(13) & "pIndex: " & pIndex("Release","Name")
'    MsgBox GetVariable("ListNumber" & pIndex("Release","Minor List") + 1, vClass, 2, 1, 0, nDebug)
'    MsgBox objMain(1,pIndex(0,"ImageTemplate"))	
	
'-------------------------------------------------------------------------------------------
'  		LOAD TEMPORARY PARAMETERS FOR CURRENT/LAST SESSION
'-------------------------------------------------------------------------------------------
	If objFSO.FileExists(strDirectoryTmp & "\" & strFileSessionTmp) Then 
		nSessionTmp = GetFileLineCountSelect(strDirectoryTmp & "\" & strFileSessionTmp, vSessionTmp, "#", "", "", 0)
	Else
	    Redim vSessionTmp(7)
		vSessionTmp(0) = 0
		vSessionTmp(1) = 0
		vSessionTmp(2) = 0
		vSessionTmp(3) = 0
		vSessionTmp(4) = "N/A"
		vSessionTmp(5) = 0		
		vSessionTmp(6) = 0				
	End If

'#####################################################################################
'       MAIN PROGRAM
'#####################################################################################	
	Do
        nResult = IE_PromptForInput(vIE_Scale, vCfgList,vSessionCRT, vSessionTmp, vSessionEnable, vSettings, vCfgInventory, nDebug)
        Select Case nResult
		    Case False
			        ' Call FocusToParentWindow(strPID)
                    exit sub
			Case 1
			        '---------------------------------------
					'    UPDATE CfgList
					'---------------------------------------
					nCount = GetFileLineCountByGroup(strCfgFile, vCfgInventory,"Inventory","","",0)
					Redim vCfgList(nCount,UBound(vSessionCRT) + 2)
					nCfg = 0 
					For Each strCfg in vCfgInventory
					   If strCfg = "" Then Exit For
					   vCfgList(nCfg,0) = strCfg
					   nCount = GetFileLineCountByGroup(strCfgFile, vCfg, strCfg,"","",0)
					   nInd = 1
					   For Each LineItem in vCfg
						  If LineItem = "" Then Exit For
						  vCfgList(nCfg,nInd) = LineItem
						  nInd = nInd + 1
					   Next
					   nCfg = nCfg + 1
					Next
			Case Else 
                    exit Do
	    End Select 
	Loop
	objDebug.Close
End Sub
'##############################################################################
'      Function Displays a Message with OK Button. Returns True.
'##############################################################################
 Function IE_MSG (vIE_Scale, strTitle, ByRef vLine, ByVal nLine, objIEParent)
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim nDebug, cellW, CellH
	Dim g_objIE, objShell
    Set g_objIE = Nothing
    Set objShell = Nothing
	nDebug = 0
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_MSG = True
	Call IE_Hide(objIEParent)
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nRatioX = intX/1920
    nRatioY = intY/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((130 + nLine * 35) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
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
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center;font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
	Next		
	
    strHTMLBody = strHTMLBody &_
                "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
				"; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; left: " & Int((CellW - nButtonX)/2) & "px; bottom: 4px' name='OK' AccessKey='O' onclick=document.all('ButtonHandler').value='OK';><u>O</u>K</button>" & _
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
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title

	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strMyPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then strMyPID = Split(strLine,""",""")(1)
	     ' Call TrDebug("READ TASK PID:" , strLine, objDebug, MAX_LEN, 1, 1)
    Next
    If strMyPID = "" Then Call GetAppPID(strMyPID, "iexplore.exe")
	objShell.AppActivate strMyPID
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "OK"
				IE_MSG = True
				g_objIE.quit
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
		Call IE_UnHide(objIEParent)
End Function
'###################################################################################
' Function returns True if object/string exists in data file                 
'###################################################################################
Function MyObjectExist( byRef strFilePath, byRef strObjectName)
	    MyObjectExist = False
		Set objFileObject = objFSO.OpenTextFile(strFilePath)
		Do While objFileObject.AtEndOfStream <> True
            vLine = Split(objFileObject.ReadLine,",") 
	        If vLine(0) = strObjectName Then 
                MyObjectExist = True
			End If
        Loop
	    objFileObject.Close 
End Function
'###################################################################################
'  Authenticate User against its password. Requires an account data file as input                  
'###################################################################################
Function Authenticate( byRef strFilePath, byRef strObjectName, byRef passwd)
	    Authenticate = False
		Set objFileObject = objFSO.OpenTextFile(strFilePath)
		Do While objFileObject.AtEndOfStream <> True
            vLine = Split(objFileObject.ReadLine,",") 
	        If vLine(0) = strObjectName Then 
                If passwd = vLine(2) Then Authenticate = True End If
			End If
        Loop
	    objFileObject.Close 
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
 ' Function GetFileLineCount - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCount(strFileName, ByRef vFileLines, nDebug)
    Dim nIndex
	Dim strLine
	Dim objDataFileName
	
    strFileWeekStream = ""	
	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		If 	Left(strLine,1)<>"#" Then
			vFileLines(nIndex) = strLine
			If nDebug = 1 Then objDebug.WriteLine "GetFileLineCount: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
			nIndex = nIndex + 1
		End If
	Loop
	objDataFileName.Close
    GetFileLineCount = nIndex
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
			WScript.Sleep 200
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
' ----------------------------------------------------------------------------------------------
'   Function  TrDebug (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug (strTitle, strString, objDebug, nChar, nFormat, nDebug)
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
'-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
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
'---------------------------------------------------------------------------------------
'   Function FindAndReplaceStrInFile(strFile, strFind, strNewLine, nDebug)
'   Search for the First Line which contains "strFind" and Replaces whole Line with "strNewLine"
'---------------------------------------------------------------------------------------
Function FindAndReplaceStrInFile(strFile, strFind, strNewLine, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Const FOR_WRITING = 1

	FindAndReplaceStrInFile = False
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLine is number of lines counted like 1,2,...,n
	LineNumber = GetObjectLineNumber( vFileLine, nFileLine, strFind)
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": FindAndReplaceStrInFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	vFileLine(LineNumber - 1) = strNewLine
	If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then FindAndReplaceStrInFile = True End If
End Function
'---------------------------------------------------------------------------------------
'   Function FindStrInFile(strFile, strFind, strNewLine, nDebug)
'   Search for the First Line which contains "strFind" and Replaces whole Line with "strNewLine"
'---------------------------------------------------------------------------------------
Function FindStrInFile(strFile, strFind, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Const FOR_WRITING = 1
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLine is number of lines counted like 1,2,...,n
	LineNumber = GetObjectLineNumber( vFileLine, nFileLine, strFind)
	If LineNumber > 0 Then 
		FindStrInFile = LineNumber 
	Else 
		FindStrInFile = 0
	End If
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
				wscript.sleep 1000
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
 '------------------------------------------------------------------------------
 ' Function GetTestSeries
 '------------------------------------------------------------------------------
 Function GetTestSeries(strFileName, ByRef vSvc, ByRef vFlavors, nDebug)
    Dim nIndex, i, n, nParam
	Dim strLine
	Dim nGroupSelector, nService, nMaxFlavors, nFlavors
	Dim vFileLines, nFileLines, vService, vLines	
	GetTestSeries = 0
	nGroupSelector = 0
	nMaxFlavors = 0
	nService = GetFileLineCountByGroup(strFileName , vService,"Service","","",0)
	Redim vSvc(nService,2)
	nFileLines = GetFileLineCountSelect(strFileName, vFileLines,"#","NULL","NULL",0)
		
	'-----------------------------------------------------
	'	COUNT Tasks
	'-----------------------------------------r------------
	For n = 0 to nService - 1
		nFlavors = 0
		nFlavors = GetFileLineCountByGroup(strFileName , vLines,vService(n),"","",0)
		vSvc(n,0) = nFlavors
		vSvc(n,1) = vService(n)
		Call TrDebug ("GetTestSeries: TOTAL TASKS IN CATEGORY: " & vService(n), nFlavors, objDebug, MAX_LEN, 1, nDebug)
	Next
	'-----------------------------------------------------
	'	FIND THE MAXIMUM NUMBER OF ALL TASKS
	'-----------------------------------------------------
	For n = 0 to nService - 1
		nMaxFlavors = MaxQ(nMaxFlavors, vSvc(n,0)) 		
	Next
	Call TrDebug ("GetTestSeries: nMaxFlavor = " & nMaxFlavors, "", objDebug, MAX_LEN, 1, nDebug)
	'-----------------------------------------------------
	'	DEFINE vFlavors Array
	'-----------------------------------------------------
	Redim vFlavors(nService, nMaxFlavors,2)		
	'-----------------------------------------------------
	'	LOAD CATEGORIES PROPERIES
	'-----------------------------------------------------
		nGroupSelector = 0
		For n = 0 to nService - 1
			For nIndex = 0 to nFileLines - 1
				strLine = LTrim(vFileLines(nIndex))
				Select Case Left(strLine,1)
					Case "#"
					Case ""
					Case "["
						Select Case strLine
							Case "[" & vService(n) & "]"
								Call TrDebug ("GetTestSeries: LOAD PROPERTIES FOR [" & vService(n) & "]", "", objDebug, MAX_LEN, 3, nDebug)
								nParam = 0
								nGroupSelector = 1
							Case Else
								nGroupSelector = 0
						End Select
					Case Else	
						If nGroupSelector = 1 Then 
							Call TrDebug ("GetTestSeries:" & strLine, "", objDebug, MAX_LEN, 1, nDebug)					
							If nParam < nMaxFlavors Then 
								vFlavors(n,nParam,0) = Split(strLine,":")(0)
								vFlavors(n,nParam,1) = Split(strLine,":")(1)
								Call TrDebug ("GetTestSeries: Flavors(" & n &  "," & nParam & ",0) = "  & vFlavors(n,nParam,0), "", objDebug, MAX_LEN, 1, 1)	
								Call TrDebug ("GetTestSeries: Flavors(" & n &  "," & nParam & ",1) = "  & vFlavors(n,nParam,1), "", objDebug, MAX_LEN, 1, 1)	
							Else 
								Call TrDebug ("GetTestSeries: ERROR: Flavors(nParam) overflow > " & nMaxFlavors, "", objDebug, MAX_LEN, 1, 1)					
							End If
							nParam = nParam + 1 
						End If
				End Select
			Next
		Next
	GetTestSeries = nService
End Function
'------------------------------------------------
'    MAIN DIALOG FORM 
'------------------------------------------------
Function IE_PromptForInput(ByRef vIE_Scale, ByRef vCfgList, ByRef vSessionCRT, ByRef vSessionTmp,Byref vSessionEnable, byRef vSettings, byref vCfgInventory, nDebug)
	Dim g_objIE, g_objShell, objMonitor
	Dim vFileLine
	Dim nInd, Arg4, CFG_Downloaded, YES_NO
	Dim nRatioX, nRatioY, nFontSize_10, nFontSize_12, nButtonX, nButtonY, nA, nB
    Dim intX
    Dim intY
	Dim nCount
	Dim strLogin
	Dim IE_Menu_Bar
	Dim IE_Border
	Dim nLine, nService, nFlavor, nTask
	Dim vvMsg(8,3)
	Dim nMaxFlavors
	Dim objFile, objCfgFile
	Dim T_csr, T_ag1, T_ag2, T_ag3, csr_Notations, ag1_Notations, ag2_Notations, ag3_Notations, vTag
	Dim nCfg, nTag, nVersion, strCfg, CurrentCfg, strBulkList, vBulkList
	Const MAX_PARAM = 10
	Const MAX_BW_PROFILES = 30
	Const TCG_MONITOR = "TCG_monitor"
	Dim objCell
	Set g_objIE = Nothing
    Set g_objShell = Nothing
	Call TrDebug ("IE_PromptForInput: OPEN MAIN CONFIG LOADER FORM ", "", objDebug, MAX_LEN, 3, nDebug)				
	vTag = Array("*", "MBH", "Metro", "RSVP", "ZTD", "ZTP", "CE20")
	'-----------------------------------------------------
	'   SORT CRT SESSIONS INTO FOUR GROUPS
	'-----------------------------------------------------
	csr_Notations = Array("csr","an")
	ag1_Notations = Array("ag1","ag-1","ag.1","pe1","pe-1","pe.1","agn1","agn-1","agn.1")
	ag2_Notations = Array("ag2","ag-2","ag.2","pe2","pe-2","pe.2","agn2","agn-2","agn.2","asbr1","asbr-1","asbr.1","bng")
	ag3_Notations = Array("ag3","ag-3","ag.3","pe3","pe-3","pe.3","agn3","agn-3","agn.3","asbr2","asbr-2","asbr.2")
	Redim T_csr(1) : T_csr(0) = "N/A"
	Redim T_ag1(1) : T_ag1(0) = "N/A"
	Redim T_ag2(1) : T_ag2(0) = "N/A"
	Redim T_ag3(1) : T_ag3(0) = "N/A"
	Dim csrInd, ag1Ind, ag2Ind, ag3Ind, Tmax_Ind
	csrInd=0 : ag1Ind=0 : ag2Ind=0 : ag3Ind=0
	For nInd = 0 to UBound(vSessionCRT) - 1
	    Do
	        strSessionName = Split(vSessionCRT(nInd),",")(2)
			For Each LineItem in csr_Notations
			    If InStr(strSessionName, LineItem) > 0 Then 
			        Redim Preserve T_csr(csrInd + 1)
					T_csr(csrInd) = strSessionName
                    csrInd = csrInd + 1
					Exit Do
			    End If 
			Next
			For Each LineItem in ag1_Notations
			    If InStr(strSessionName, LineItem) > 0 Then 
			        Redim Preserve T_ag1(ag1Ind + 1)
					T_ag1(ag1Ind) = strSessionName
                    ag1Ind = ag1Ind + 1
					Exit Do
			    End If 
			Next
			For Each LineItem in ag2_Notations
			    If InStr(strSessionName, LineItem) > 0 Then 
			        Redim Preserve T_ag2(ag2Ind + 1)
					T_ag2(ag2Ind) = strSessionName
                    ag2Ind = ag2Ind + 1
					Exit Do
			    End If 
			Next
			' If no match found then place session into fourth row
			Redim Preserve T_ag3(ag3Ind + 1)
			T_ag3(ag3Ind) = strSessionName
			ag3Ind = ag3Ind + 1
	        Exit Do
		Loop
	Next
	Tmax_Ind = MaxQ(UBound(T_csr),Ubound(T_ag1))
	Tmax_Ind = MaxQ(Tmax_Ind,Ubound(T_ag2))
	Tmax_Ind = MaxQ(Tmax_Ind,Ubound(T_ag3))
	
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	'----------------------------------------
	' IE EXPLORER OBJECTS
	'----------------------------------------
    Set g_objShell = WScript.CreateObject("WScript.Shell")
    Call Set_IE_obj (g_objIE)
    g_objIE.Offline = True
    g_objIE.navigate "about:blank"
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	'----------------------------------------
	' MAIN VARIABLES OF THE GUI FORM
	'----------------------------------------
	If nRatioX > 1 Then nRatioX = 1 : nRatioY = 1 End If
	Select Case nRatioX
		Case 1
				DiagramFigure = strDirectoryWork & "\Data\TestBed005.png"
		Case 1600/1920
				DiagramFigure = strDirectoryWork & "\Data\TestBed002.png"
		Case else
				DiagramFigure = strDirectoryWork & "\Data\TestBed002.png"
				nRatioX = 1600/1920
				nRatioX = 900/1080
	End Select
	SettingsFigure = strDirectoryWork & "\data\settings-icon-7.png"
	BgFigure = strDirectoryWork & "\Data\bg_image_03.jpg"
	AttentionFigure = strDirectoryWork & "\Data\Attention_icon_30x30.png"
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	nBottom = Round(10 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellH = Round(24 * nRatioY,0)
	LoginTitleW = Round(700 * nRatioX,0)
	FullTitleW = LoginTitleW + Int(LoginTitleW/4)
	nLeft = Round(20 * nRatioX,0)
	nTab = Round(40 * nRatioX,0)
	CellW = LoginTitleW
	LoginTitleH = Round(40 * nRatioY,0)
	nSaveW = nLeft + nButtonX
	nScoreW = 3 * nSaveW
	nColumn = Int(LoginTitleW/3)	
	nNameW = Int((LoginTitleH - nColumn)/3)
	'------------------------------------------
	'	GET NUMBER OF TASKS LINES
	'------------------------------------------	
	nLine = 15
	WindowH = IE_Menu_Bar + 4 * LoginTitleH + cellH * (nLine) + nButtonY + nBottom
	WindowW = IE_Border + FullTitleW
	MenuH = WindowH - IE_Menu_Bar
	If WindowW < 300 then WindowW = 300 End If

	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)	
   	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "None "
'	g_objIE.Document.body.Style.background = "transparent url('" & BgFigure & "')"
	g_objIE.Document.body.Style.color = HttpTextColor1
    g_objIE.height = WindowH
    g_objIE.width = WindowW  
    g_objIE.document.Title = MAIN_TITLE
	g_objIE.Top = (intY - g_objIE.height)/2
	g_objIE.Left = (intX - g_objIE.width)/2
	g_objIE.Visible = False		
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title

    '-----------------------------------------------------------------
	' Create Background Table  		
	'-----------------------------------------------------------------
    htmlEmptyCell = _
        	"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/6) & """></td>"
	nLine = nLine + 2
   htmlMain = htmlMain &_	
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" height=""" & 4 * LoginTitleH + cellH * (nLine) + nButtonY + nBottom &_
		""" width=""" & FullTitleW & """ valign=""middle"" background=""" & bgFigure & """ background-repeat=""no-repeat""" &_ 
		"style="" position: absolute; top: 0px; left:0px;" &_
		"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" &_
			    "<tr>"&_
					htmlEmptyCell &_
				"</tr>" &_
				"</tbody>" &_
		"</table>"
	'------------------------------------------------------
	'   MENU BUTTONS 
	'------------------------------------------------------
    nMenuButtonX = Int(LoginTitleW/4)
	nMenuButtonY = nButtonY
	htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: " &  LoginTitleH & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: Transparent" &_
		"; height: " & LoginTitleH & "px; width: " & Int(LoginTitleW/4) & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & 2 * LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & 2 * nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='EPTY' onclick=document.all('ButtonHandler').value='EMPTY';></button>" & _	
					"</td>"&_
				"</tr>" &_
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"' name='LOAD' onclick=document.all('ButtonHandler').value='LOAD';><u>L</u>oad Config</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='Save_Tested' onclick=document.all('ButtonHandler').value='SAVE_TESTED';><u>S</u>ave Tested Config</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
    					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='Save_as' AccessKey='E' onclick=document.all('ButtonHandler').value='SAVE_AS';><u>S</u>ave as...</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
    					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='Blk_DownLoad' AccessKey='B' onclick=document.all('ButtonHandler').value='BLK_DOWNLOAD';><u>B</u>ulk Save</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px;' name='EDIT' onclick=document.all('ButtonHandler').value='EDIT';><u>E</u>dit Config</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
     					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='Delete_Cfg' onclick=document.all('ButtonHandler').value='DELETE_CFG';><u>D</u>elete Config</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
    					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='POPULATE_ORIG' AccessKey='P' onclick=document.all('ButtonHandler').value='POPULATE_ORIG';>TCG <u>E</u>xport Original</button>" & _
					"</td>"&_
				"</tr>" &_					
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
    					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='UPGRADE_SW' AccessKey='P' onclick=document.all('ButtonHandler').value='UPGRADE_SW';><u>U</u>pgrade Junos</button>" & _
					"</td>"&_
				"</tr>" &_					
				"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & 2 * LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & 3 * LoginTitleH & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='EPTY' onclick=document.all('ButtonHandler').value='EMPTY';></button>" & _	
					"</td>"&_
				"</tr></tbody></table>" &_
				"<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
		nButton = 7
    '-----------------------------------------------------------------
	' SET THE TITLE OF THE FORM   		
	'-----------------------------------------------------------------
	nLine = 0
	    htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: " & HttpBgColor5 & "; height: " & LoginTitleH & "px; width: " & FullTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td  style=""border-style: none; background-color: " & HttpBgColor5 & ";""" &_
			"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & FullTitleW - nTab & """>" & _
				"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
				";font-weight: normal;font-style: italic;"">"&_
				"&nbsp;&nbsp;Lab Configuration Loader <span style=""font-weight: bold;"">Ver." & strVersion & "</span></span></p>"&_
			"</td>" &_
			"<td background=""" & SettingsFigure & """ style=""background-repeat: no-repeat; background-position: 50% 50%; background-size: 40px 40px;"&_
				"border-style: none; background-color: " & HttpBgColor5 & ";""" &_
			    "valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & nTab & """>" & _
				"<button style='background-color: transparent; border-style: None; width:" &_
				"40;height:40;" &_
				"' name='SETTINGS' onclick=document.all('ButtonHandler').value='SETTINGS_';></button>" & _	
			"</td>" &_			
			"</tr></tbody></table>"
	
		'-----------------------------------------------------------------
		' DRAW CONFIGURATION TABLE
		'-----------------------------------------------------------------
		'-----------------------------------------------------------------
		' DRAW ROW WITH CONFIGURATION TITLE
		'-----------------------------------------------------------------	
		cTitle = "Choose LAB Configuration"
	    htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: " & Int(LoginTitleW/4) & "px; top: " & LoginTitleH & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: Transparent; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: None;""" &_
			"align=""center"" valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW & """>" & _
				"<p><span style="" font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
				";font-weight: normal;font-style: italic;"">" & cTitle & "</span></p>"&_
			"</td>" &_
		"</tr></tbody></table>"
		'-----------------------------------------------------------------
		' DRAW ROW WITH CONFIGURATIONS TO BE LOADED
		'-----------------------------------------------------------------
		strTitleCell = "<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: center; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">"

		htmlMain = htmlMain &_
			"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: " & Int(LoginTitleW/4) & "px; top: " & 2 * LoginTitleH + nLine * cellH & "px;" &_
			" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: Transparent"  &_
			"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					strTitleCell & "</p></td>"&_
					strTitleCell & "</p></td>"&_
					strTitleCell & "</p></td>"&_
					strTitleCell & "</p></td>"&_					 
				"</tr>"	&_		
				"<tr>" &_
					strTitleCell & "Year</p></td>"&_
					strTitleCell & "Search Tag</p></td>"&_
					strTitleCell & "Cfg Version</p></td>"&_
					strTitleCell & "Use Saved CFG</p></td>"&_					 
				"</tr>"
					'-----------------------------------------------------
					'  SELECT DATE/YEAR
					'-----------------------------------------------------
					htmlMain = htmlMain &_
					"<tr>"
						htmlMain = htmlMain &_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<select name='Input_Param_0' id='Input_Param_0'" &_
											"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
											";position: relative; left:" & nTab & "px; " &_
											"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
											"; background-color:" & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
											" onchange=document.all('ButtonHandler').value='Select_0';" &_
											"type=text > "
						htmlMain = htmlMain &_
											"<option value='All'>All</option>" 

					For nYear = 2012 to Year(Date())
						htmlMain = htmlMain &_
											"<option value='" & nYear & "'>" & nYear & "</option>" 
					Next
					htmlMain = htmlMain &_
    							"<option value='All'>"& Space_html(18)&"</option>" &_
							"</select>" &_
						"</td>"
					'-----------------------------------------------------
					'  SELECT SEARCH TAG
					'-----------------------------------------------------
					htmlMain = htmlMain &_
						"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & nNameW & """>" &_
							"<select name='Input_Param_1' id='Input_Param_1'" &_
											"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
											";position: relative; left:" & nTab & "px; " &_
											"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
											"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
											" onchange=document.all('ButtonHandler').value='Select_1';" &_
											"type=text > "
					For Each strTag in vTag
						If strTag = "" Then Exit For
						htmlMain = htmlMain &_
											"<option value='" & strTag & "'>" & strTag & "</option>" 
					Next
						htmlMain = htmlMain &_
											"<option value='*'>"& Space_html(18)& "</option>" 
					htmlMain = htmlMain &_
							"</select>" &_
						"</td>"
					'-----------------------------------------------------
					'  SELECT VERSION
					'-----------------------------------------------------
					htmlMain = htmlMain &_
						"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & nNameW & """>" &_
							"<select name='Input_Param_2' id='Input_Param_2'" &_
											"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
											";position: relative; left:" & nTab & "px; " &_
											"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
											"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
											" onchange=document.all('ButtonHandler').value='Select_2';" &_
											"type=text > "
					For nVersion = 0 to MAX_PARAM
						htmlMain = htmlMain &_
											"<option value='Latest'>"& Space_html(18)& "</option>" 
					Next
					htmlMain = htmlMain &_
							"</select>" &_
						"</td>"
					htmlMain = htmlMain &_
						"<td style="" border-style: None; background-color: Transparent;"" class=""oa2"" height=""" & LoginTitleH & """ width=""" & nTab & """ align=""middle"">" & _
							"<input type=checkbox name='ConfigLocation' style=""color: " & HttpTextColor2 & ";""" & _
							" onclick=document.all('ButtonHandler').value='CONFIG_SOURCE';" &_
							"value='Original'>" &_
						"</td>"&_
				"</tr>" &_
				"<tr>" &_
					strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>"&_					 
				"</tr>"								
	nLine = nLine + 4
	'-----------------------------------------------------
	'  END OF TABLE
	'-----------------------------------------------------
				htmlMain = htmlMain &_
						"</tbody></table>"
	'-----------------------------------------------------
	'  SELECT CFG NAME
	'-----------------------------------------------------

		strTitleCell = "<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: center; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">"
		strInputCell = 	"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<input name=BW_Param value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
						" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
					    "; background-color: Transparent; font-weight: Normal;"" AccessKey=i size=12 maxlength=15 " &_
						"type=text > "

		htmlMain = htmlMain &_
			"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: " & Int(LoginTitleW/4) & "px; top: " & 2 * LoginTitleH + nLine * cellH & "px;" &_
			" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: Transparent"  &_
			"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					 strTitleCell & "</p></td>"&_
				"</tr>" &_
    			"<tr>" &_
					 strTitleCell & "Configuration Name</p></td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/5) & """>" &_
						"<select name='cfg_name' id='cfg_name'" &_
						"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
						";position: relative; left:" & nTab & "px; " &_
						"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
						" onchange=document.all('ButtonHandler').value='SELECT_CFG' type=text > "
						For nInd = 0 to UBound(vCfgList,1)-1
							htmlMain = htmlMain &_
												"<option value=" & nInd & """></option>" 
						Next
						htmlMain = htmlMain &_
									"<option value=New"">New"& Space_html(128)& "</option>" &_
								"</select>" &_
					"</td>"	&_
				"</tr>" &_
				"<tr>" &_
					strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>"&_					 
				"</tr>"	&_	
			"</tbody></table>"
	nLine = nLine + 2   						

	'------------------------------------------------------
	'   EXIT BUTTON
	'------------------------------------------------------
	htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: " & Int(LoginTitleW/4) & "px; bottom: " & LoginTitleH & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: Transparent; background-color: Transparent " &_
		"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
					"</td>"&_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
					"</td>"&_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
					"</td>"&_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
					"</td>"&_
				"</tr></tbody></table>"
	'------------------------------------------------------
	'   BOTTOM INFO BAR 
	'------------------------------------------------------
	htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: None; background-color: " & HttpBgColor5 &_
		"; height: " & LoginTitleH & "px; width: " & FullTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/8) & """>" & _
						"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
						"; font-weight: normal;font-style: italic;"">Topology: </span></p>" &_
					"</td>" & _
					"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<input name=Current_config value='' style=""text-align: left; font-size: " & nFontSize_12 & ".0pt;" &_ 
						" border-style: none; font-family: 'Helvetica'; font-style: italic; color: " & HttpTextColor3 &_
					    "; background-color: Transparent; font-weight: Normal;"" AccessKey=i size=30 maxlength=48 " &_
						"type=text > " &_
					"</td>" & _
					"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/8) & """>" & _
						"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
						"; font-weight: normal;font-style: italic;"">" & "Last Loaded:</span></p>" &_
					"</td>" & _
					"<td style=""border-style: None; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & 5 * Int(LoginTitleW/8) & """>" & _
						"<input name=Current_config value='' style=""text-align: left; font-size: " & nFontSize_12 & ".0pt;" &_ 
						" border-style: none; font-family: 'Helvetica'; font-style: italic; color: " & HttpTextColor3 &_
					    "; background-color: Transparent; font-weight: Normal;"" AccessKey=i size=30 maxlength=48 " &_
						"type=text > " &_
					"</td>" & _
					"<td style=""border-style: None; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/8) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor3 & "; width:" &_
						2 * nButtonX & ";height:" & nButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px;' name='EXIT' onclick=document.all('ButtonHandler').value='Cancel';><u>E</u>xit</button>" & _	
					"</td>" & _
		"</tr></tbody></table>"
	'-----------------------------------------------------------------
	' NETWORK DIAGRAM
	'-----------------------------------------------------------------
    htmlEmptyCell = _
        	"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/6) & """></td>"
	nLine = nLine + 3

   htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & Int(LoginTitleW/6) & """ valign=""middle""" &_ 
		"style="" position: absolute; top: " & 2 * LoginTitleH + nLine * cellH & "px; left: " & nLeft + Int(LoginTitleW/4) & "px;" &_
		"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
			"<tbody>"
				For Each LineItem in T_csr
				    If LineItem = "" Then Exit For
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" &  Int(LoginTitleW/6) & """ align=""middle"">" &_
								"<input type=checkbox name='" & LineItem & "' style=""color: " & HttpTextColor2 & ";" & _
								"position: relative; left: 0px; top: 2px; """ &_									
								" onclick=document.all('ButtonHandler').value='ENABLE_SESSION#"& LineItem &"';" &_
								"value='Original'>" &_
								"<input name=UNI value='" & LineItem & "' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
								";position: relative; left:" & Int(nLeft/2) & "px; top: 2px; " &_			
								"; background-color: transparent; font-weight: Bold;"" AccessKey=i size=6 maxlength=10 " &_
								"type=text > " &_
							"</td>" &_
						"</tr>"
				Next
		htmlMain = htmlMain &_
				"</tbody>" &_
		"</table>"	
		htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width="""  & Int(LoginTitleW/6) &  """ valign=""middle""" &_ 
		"style="" position: absolute; top: " & 2 * LoginTitleH + nLine * cellH & "px; left: " & nLeft + Int(LoginTitleW/4) + Int(LoginTitleW/6) & "px;" &_
		"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
			"<tbody>"
				For Each LineItem in T_ag1
				    If LineItem = "" Then Exit For
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/6) & """ align=""middle"">" &_
								"<input type=checkbox name='" & LineItem & "' style=""color: " & HttpTextColor2 & ";" & _
								"position: relative; left: 0px; top: 2px; """ &_									
								" onclick=document.all('ButtonHandler').value='ENABLE_SESSION#"& LineItem &"';" &_
								"value='Original'>" &_
								"<input name=UNI value='" & LineItem & "' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
								";position: relative; left:" & Int(nLeft/2) & "px; top: 2px; " &_			
								"; background-color: transparent; font-weight: Bold;"" AccessKey=i size=6 maxlength=10 " &_
								"type=text > " &_
							"</td>" &_
						"</tr>"
				Next
		htmlMain = htmlMain &_
				"</tbody>" &_
		"</table>"
     htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1""  width=""" & Int(LoginTitleW/4) & """ valign=""middle""" &_ 
		"style="" position: absolute; top: " & 2 * LoginTitleH + nLine * cellH & "px; left: " & nLeft + Int(LoginTitleW/4) + 2 * Int(LoginTitleW/6) & "px;" &_
		"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
			"<tbody>"
				For Each LineItem in T_ag2
				    If LineItem = "" Then Exit For
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/6) & """ align=""middle"">" &_
								"<input type=checkbox name='" & LineItem & "' style=""color: " & HttpTextColor2 & ";" & _
								"position: relative; left: 0px; top: 2px; """ &_									
								" onclick=document.all('ButtonHandler').value='ENABLE_SESSION#"& LineItem &"';" &_
								"value='Original'>" &_
								"<input name=UNI value='" & LineItem & "' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
								";position: relative; left:" & Int(nLeft/2) & "px; top: 2px; " &_			
								"; background-color: transparent; font-weight: Bold;"" AccessKey=i size=6 maxlength=10 " &_
								"type=text > " &_
							"</td>" &_
						"</tr>"
				Next
		htmlMain = htmlMain &_
				"</tbody>" &_
		"</table>"
		htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & Int(LoginTitleW/4) & """ valign=""middle""" &_ 
		"style="" position: absolute; top: " & 2 * LoginTitleH + nLine * cellH & "px; left: " & nLeft + Int(LoginTitleW/4) + 3 * Int(LoginTitleW/6) & "px;" &_
		"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
			"<tbody>"
				For Each LineItem in T_ag3
				    If LineItem = "" Then Exit For
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/6) & """ align=""middle"">" &_
								"<input type=checkbox name='" & LineItem & "' style=""color: " & HttpTextColor2 & ";" & _
								"position: relative; left: 0px; top: 2px; """ &_									
								" onclick=document.all('ButtonHandler').value='ENABLE_SESSION#"& LineItem &"';" &_
								"value='Original'>" &_
								"<input name=UNI value='" & LineItem & "' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
								";position: relative; left:" & Int(nLeft/2) & "px; top: 2px; " &_			
								"; background-color: transparent; font-weight: Bold;"" AccessKey=i size=6 maxlength=10 " &_
								"type=text > " &_
							"</td>" &_
						"</tr>"
				Next
		htmlMain = htmlMain &_
    		"</tbody>" &_
		"</table>"	
	'-----------------------------------------------------------------
	' SETTINGS MENU
	'-----------------------------------------------------------------
	htmlMain = htmlMain & _
	 "<div id='divSettings' name='divSettings' style='color: " & HttpBgColor2 & " ;background-color:" & HttpBgColor5 & "; width: 200px; height: " & MenuH - 2 * LoginTitleH & "px; position: absolute; right: -200px; top: " & LoginTitleH & "px;'>" &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: relative; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: Transparent" &_
		"; height: " & LoginTitleH & "px; width: " & Int(LoginTitleW/4) & "px;"">" & _
			"<tbody>" & _
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
					"</td>"&_
				"</tr>" &_
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='SET_CFG_TEMPLATE' onclick=document.all('ButtonHandler').value='SET_CFG_TEMPLATE';>CFG Template Name</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/4) & """ width=""" & LoginTitleW & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='SET_CONNECTIVITY' onclick=document.all('ButtonHandler').value='SET_CONNECTIVITY';>LAB Connectivity</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
    					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='SET_FOLDER' AccessKey='E' onclick=document.all('ButtonHandler').value='SET_FOLDER';>Folder Settings</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
    					"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='SET_CRT' AccessKey='E' onclick=document.all('ButtonHandler').value='SET_CRT';>SecureCRT Sessions</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px;' name='SET_OTHER' onclick=document.all('ButtonHandler').value='SET_OTHER';>Other Settings</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor4 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px;' name='JUNOS_CATALOG' onclick=document.all('ButtonHandler').value='SET_JUNOS_CATALOG';>Junos Catalog</button>" & _	
					"</td>"&_
				"</tr>" &_				
		    "</tbody></table>" &_
	 "</div>"

	'-----------------------------------------------------------------
	' HTML Form Parameaters
	'-----------------------------------------------------------------
    g_objIE.Document.Body.innerHTML = htmlMain
    g_objIE.MenuBar = False
    g_objIE.StatusBar = False
    g_objIE.AddressBar = False
    g_objIE.Toolbar = False
    ' Wait for the "dialog" to be displayed before we attempt to set any
    ' of the dialog's default values.
	'-----------------------------------------------------------------
	'	SET DEFAULT PARAMETERS
	'-----------------------------------------------------------------
	CFG_Downloaded = False
	YES_NO = False
	nYear = Int(vSessionTmp(0))
	nTag = Int(vSessionTmp(1))
	nCfg = Int(vSessionTmp(2))
	nVersion = Int(vSessionTmp(3))
	CurrentCfg = vSessionTmp(4)
	If vSessionTmp(5) = "0" Then YES_NO = False Else YES_NO = True
	g_objIE.Document.All("ButtonHandler").Value = vSessionTmp(6)

	g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)
	g_objIE.Document.All("Current_config")(1).Value = CurrentCfg

	'--------------------------------------
	' WAIT UNTIL IE FORM LOADED
	'--------------------------------------
    Do
        WScript.Sleep 100
    Loop While g_objIE.Busy
	if g_objIE.Document.All("ConfigLocation").Checked then	
		SourceFolder = strDirectoryConfig & "\Tested"
		Arg4 = "tested"
	Else
		SourceFolder = strDirectoryConfig
		Arg4 = "original"
    end if
	'--------------------------------------
	' LOAD INITAIAL YEAR AND SEARCH TAG
	'--------------------------------------
    g_objIE.document.getElementById("Input_Param_0").selectedIndex = nYear
	g_objIE.document.getElementById("Input_Param_1").selectedIndex = nTag
	strYear = g_objIE.document.getElementById("Input_Param_0").options(nYear).Value
	strTag = g_objIE.document.getElementById("Input_Param_1").options(nTag).Value
	'--------------------------------------
	' LOAD CONFIGURATION LIST
	'--------------------------------------
	nOptions = UpdateCfgList(g_objIE, nCfg, strYear, strTag, vCfgList, "cfg_name")
    g_objIE.document.getElementById("cfg_name").SelectedIndex = nOptions	
	strCfg = g_objIE.document.getElementById("cfg_name").Options(nOptions).Text
	nCfg = g_objIE.document.getElementById("cfg_name").Options(nOptions).Value

'	strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
'	strConfigFileL = SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL
'	    g_objIE.document.getElementById("bw_profile").Options(0).Text = "N/A"
'	    g_objIE.document.getElementById("bw_profile").SelectedIndex = 0
'	    g_objIE.document.getElementById("bw_profile").Disabled = True
	'--------------------------------------
	' UPDATE CFG VERSIONS
	'--------------------------------------
	nVersion = UpdateCfgVer(g_objIE, nCfg, vCfgList, "Input_Param_2")
	strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).text							
	'--------------------------------------
	'  SESSION STATUS CHECK BOXES
	'--------------------------------------

'	Select Case vSessionTmp(6)
'	    Case "DOWNLOAD"
'			Call UpdateSessionStatus(g_objIE, nCfg, SAVE_AS, vCfgList,vSessionCRT, vSessionEnable)
'	    Case Else 
'			Call UpdateSessionStatus(g_objIE, nCfg, strCfg, vCfgList,vSessionCRT, vSessionEnable)
'	End Select
    Call UpdateSessionStatus(g_objIE, nCfg, strCfg, vCfgList,vSessionCRT, vSessionEnable)
    '--------------------------------------
	Call TrDebug("MAIN FORM:" , g_objIE.document.getElementById("Input_Param_1").Options(0).text, objDebug, MAX_LEN, 1, 1)
	Call TrDebug("MAIN FORM:" , g_objIE.document.getElementById("Input_Param_1").Options(1).text, objDebug, MAX_LEN, 1, 1)
	'--------------------------------------
	' LOAD LAST REMEMBERED PROFILE
	'--------------------------------------
	'nService = vSessionTmp(0)
	'nFlavor = vSessionTmp(1)
	'nTaskInd = vSessionTmp(2)
	'nTask = Split(vFlavors(nService,nFlavor,1),",")(nTaskInd)

	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	Call IE_Unhide(g_objIE)
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strPID = ""
	For Each strItem in vCmdOut
	   If InStr(strItem,"iexplore.exe") then strPID = Split(strItem,""",""")(1)
    Next
	g_objShell.AppActivate strPID	
	'----------------------------------------------------
	'  START MAIN CYCLE OF THE INPUT FORM
	'----------------------------------------------------
	' g_objIE.Document.All("ButtonHandler").Value = "CHECK"
	nPressSettings = 0
   Do
        ' If the user closes the IE window by Alt+F4 or clicking on the 'X'
        ' button, we'll detect that here, and exit the script if necessary.
        On Error Resume Next
			If g_objIE.width <> WindowW Then g_objIE.width = WindowW End If
			If g_objIE.height <> WindowH Then g_objIE.height = WindowH End If
			MenuH = WindowH - IE_Menu_Bar
			Err.Clear
            szNothing = g_objIE.Document.All("ButtonHandler").Value
            if Err.Number <> 0 then exit function
        On Error Goto 0    
        ' Check to see which buttons have been clicked, and address each one
        ' as it's clicked.
        Select Case Split(szNothing,"#")(0)
		    Case "ENABLE_SESSION"
			    g_objIE.Document.All("ButtonHandler").Value = "None"
				strSessionName = split(szNothing,"#")(1)
				nInd = GetObjectLineNumber(vSessionCRT,UBound(vSessionCRT),strSessionName)
				If g_objIE.Document.All(strSessionName).Checked Then 
					vSessionEnable(nInd - 1) = "Status " & nInd & "=Enabled"
				Else 
					vSessionEnable(nInd - 1) = "Status " & nInd & "=Disabled"
				End If
		    Case "SET_CFG_TEMPLATE"
                g_objIE.Document.All("ButtonHandler").Value = "None"
        	    For MenuX = 200 to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
				Next
				nPressSettings = nPressSettings + 1				
				Call IE_Hide(g_objIE)
        	    If IE_PromptForSettings(vIE_Scale, 1, vSettings, vSessionCRT, vPlatforms, nDebug) = -1 Then g_objIE.Document.All("ButtonHandler").Value = "Cancel"
                Call IE_Unhide(g_objIE)
			    g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)

			Case "SET_CONNECTIVITY"
                g_objIE.Document.All("ButtonHandler").Value = "None"
        	    For MenuX = 200 to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
				Next
				nPressSettings = nPressSettings + 1				
				Call IE_Hide(g_objIE)
        	    If IE_PromptForSettings(vIE_Scale, 2, vSettings, vSessionCRT, vPlatforms, nDebug) = -1 Then g_objIE.Document.All("ButtonHandler").Value = "Cancel"
                Call IE_Unhide(g_objIE)
			    g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)
			
			Case "SET_FOLDER"
                g_objIE.Document.All("ButtonHandler").Value = "None"
        	    For MenuX = 200 to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
				Next
				nPressSettings = nPressSettings + 1				
				Call IE_Hide(g_objIE)
        	    If IE_PromptForSettings(vIE_Scale, 3, vSettings, vSessionCRT, vPlatforms, nDebug) = -1 Then g_objIE.Document.All("ButtonHandler").Value = "Cancel"
                Call IE_Unhide(g_objIE)
			    g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)
			
			Case "SET_CRT"
                g_objIE.Document.All("ButtonHandler").Value = "None"
        	    For MenuX = 200 to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
				Next
				nPressSettings = nPressSettings + 1				
				Call IE_Hide(g_objIE)
        	    If IE_PromptForSettings(vIE_Scale, 4, vSettings, vSessionCRT, vPlatforms, nDebug) = -1 Then g_objIE.Document.All("ButtonHandler").Value = "Cancel"
                Call IE_Unhide(g_objIE)
			    g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)
			
			Case "SET_OTHER"
                g_objIE.Document.All("ButtonHandler").Value = "None"
        	    For MenuX = 200 to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
				Next
				nPressSettings = nPressSettings + 1				
				Call IE_Hide(g_objIE)
        	    If IE_PromptForSettings(vIE_Scale, 5, vSettings, vSessionCRT, vPlatforms, nDebug) = -1 Then g_objIE.Document.All("ButtonHandler").Value = "Cancel"
                Call IE_Unhide(g_objIE)
			    g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)
			Case "SET_JUNOS_CATALOG"
                g_objIE.Document.All("ButtonHandler").Value = "None"
        	    For MenuX = 200 to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
				Next
				nPressSettings = nPressSettings + 1				
				Call IE_Hide(g_objIE)
        	    If IE_PromptForSettings(vIE_Scale, 8, vSettings, vSessionCRT, vPlatforms, nDebug) = -1 Then g_objIE.Document.All("ButtonHandler").Value = "Cancel"
                Call IE_Unhide(g_objIE)
			    g_objIE.Document.All("Current_config")(0).Value = Split(vSettings(13),"=")(1)
		    Case "SETTINGS_"
				g_objIE.Document.All("ButtonHandler").Value = "None"
			    Select Case nPressSettings Mod 2
				    Case 0
							 For MenuX = 0 to 200 step 2
							    g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
								'WScript.Sleep 1
							 Next
							 
							 nPressSettings = nPressSettings + 1
					Case Else
							 For MenuX = 200 to 0 step - 2
							    g_objIE.document.getElementById("divSettings").style.right = (MenuX - 200) & "px"
								'WScript.Sleep 1
							 Next
						nPressSettings = nPressSettings + 1
				End Select
		    Case "CONFIG_SOURCE"
						if g_objIE.Document.All("ConfigLocation").Checked then
						    SourceFolder = strDirectoryConfig & "\Tested"
							Arg4 = "tested"
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf") Then
								vvMsg(0,0) = "CONFIGURATION: " &  vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "HASN'T BEEN TESTED YET."                    : vvMsg(1,1) = "normal" 	: vvMsg(1,2) =  HttpTextColor2
								vvMsg(2,0) = "USE ORIGINAL CONFIGURATION INSTEAD"         : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
								Call IE_MSG(vIE_Scale, "Can't find configuration",vvMsg, 3, g_objIE)
							    SourceFolder = strDirectoryConfig
							    Arg4 = "original"
								g_objIE.Document.All("ConfigLocation").Checked = False
							End If
						Else
							SourceFolder = strDirectoryConfig
							Arg4 = "original"
						end if
						g_objIE.Document.All("ButtonHandler").Value = "Nothing is selected"
			Case "Select_0", "Select_1"
			            g_objIE.Document.All("ButtonHandler").Value = "None"
						'--------------------------------------
						' UPDATE SEARCH YEAR
						'--------------------------------------
					    nYear = g_objIE.document.getElementById("Input_Param_0").selectedindex
					    strYear = g_objIE.document.getElementById("Input_Param_0").options(nYear).Value
						'--------------------------------------
						' UPDATE SEARCH TAG
						'--------------------------------------
					    nTag = g_objIE.document.getElementById("Input_Param_1").selectedindex
					    strTag = g_objIE.document.getElementById("Input_Param_1").options(nTag).Value
						'--------------------------------------
						' UPDATE CFG LIST
						'--------------------------------------
						nOptions = UpdateCfgList(g_objIE, nCfg, strYear, strTag, vCfgList, "cfg_name")
						g_objIE.document.getElementById("cfg_name").SelectedIndex = nOptions
						strCfg = g_objIE.document.getElementById("cfg_name").Options(nOptions).Text
						nCfg = g_objIE.document.getElementById("cfg_name").Options(nOptions).Value
						'--------------------------------------
						' UPDATE CFG VERSIONS
						'--------------------------------------
                        nVersion = UpdateCfgVer(g_objIE, nCfg, vCfgList, "Input_Param_2")
						strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).text
						'--------------------------------------
						'  SESSION STATUS CHECK BOXES
						'--------------------------------------
						Call UpdateSessionStatus(g_objIE, nCfg, strCfg, vCfgList,vSessionCRT, vSessionEnable)	
            Case "Select_2" 
			            g_objIE.Document.All("ButtonHandler").Value = "None"
						nVersion = g_objIE.document.getElementById("Input_Param_2").selectedIndex
						strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).text														
			Case "SELECT_CFG"
			            g_objIE.Document.All("ButtonHandler").Value = "None"
						nOptions_New = g_objIE.document.getElementById("cfg_name").selectedIndex
						If g_objIE.document.getElementById("cfg_name").Options(nOptions_New).Value = "N/A"	Then 
						    g_objIE.document.getElementById("cfg_name").selectedIndex = nOptions
						Else 
						    nOptions = nOptions_New
						End If
						strCfg = g_objIE.document.getElementById("cfg_name").Options(nOptions).Text
						nCfg = g_objIE.document.getElementById("cfg_name").Options(nOptions).Value
						'--------------------------------------
						' UPDATE CFG VERSIONS
						'--------------------------------------
                        nVersion = UpdateCfgVer(g_objIE, nCfg, vCfgList, "Input_Param_2")
						strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).text							
						'--------------------------------------
						'  SESSION STATUS CHECK BOXES
						'--------------------------------------
						Call UpdateSessionStatus(g_objIE, nCfg, strCfg, vCfgList,vSessionCRT, vSessionEnable)						
			Case "DELETE_CFG"
						g_objIE.Document.All("ButtonHandler").Value = "Do Nothing"
						Do
							strCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Text
							nOptions = g_objIE.document.getElementById("cfg_name").selectedIndex 
							nCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Value
							nVersion = g_objIE.document.getElementById("Input_Param_2").selectedIndex  
							strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).Text
							'------------------------------------------------------------
							'   CHECK IF MULTIPLE VERSION OF THE SAME CFG ARE AVAILABLE
							'------------------------------------------------------------
							bDelete = 0
							Call GetFileLineCountByGroup(strDirectoryConfig & "\CfgList.txt", vFileLine,strCfg,"","",0)
							If UBound(Split(vFileLine(0),"v.")) = 1 Then 
								vvMsg(0,0) = "DELETE ALL CONFIGURATIONS?"    : vvMsg(0,1) = "bold" 	: vvMsg(0,2) = HttpTextColor2
								vvMsg(1,0) = "Name: " & strCfg  	           : vvMsg(1,1) = "normal"  : vvMsg(1,2) = HttpTextColor1
								vButton = Array("Cancel", "Continue")
								If IE_CONT_MULT(vIE_Scale, "Delete Configuration?", vvMsg, 2, vButton, g_objIE, nDebug) > 0 Then bDelete = 1 
							Else 
								vvMsg(0,0) = "DELETING CONFIGURATION?"                : vvMsg(0,1) = "bold" 	: vvMsg(0,2) = HttpTextColor2
								vvMsg(1,0) = "Would You like to completely delete"    : vvMsg(1,1) = "normal" 	: vvMsg(1,2) = HttpTextColor1
								vvMsg(2,0) = "configuration or current version only?" : vvMsg(2,1) = "normal" 	: vvMsg(2,2) = HttpTextColor1									
								vvMsg(3,0) = "Name:    " & strCfg  	           : vvMsg(3,1) = "normal"  : vvMsg(3,2) = HttpTextColor1
								vvMsg(4,0) = "Version: " & strVersion  	       : vvMsg(4,1) = "normal"  : vvMsg(4,2) = HttpTextColor1								
								vButton = Array("Cancel", "All", "Version")
								Select Case IE_CONT_MULT(vIE_Scale, "Delete Configuration?", vvMsg, 5, vButton, g_objIE, nDebug)
								    Case 0 
							            bDelete = 0
									Case 1
							            bDelete = 1										
                                    Case 2
							            bDelete = 2										
								End Select 
							End If
							Select Case bDelete
							    Case 0 
								     Exit Do
							    Case 1 ' Delete All
								    Call FindAndReplaceExactStrInFile(strDirectoryConfig & "\CfgList.txt", strCfg, "", nDebug)
									Call DeleteFileGroup(strDirectoryConfig & "\CfgList.txt", strCfg, 0)
									If objFSO.FolderExists(strDirectoryConfig & "\" & strCfg) Then 
										objFSO.DeleteFolder  strDirectoryConfig & "\" & strCfg, true
									End If
								Case 2 ' Delete current version only
									' Call GetFileLineCountByGroup(strDirectoryConfig & "\CfgList.txt", vFileLine,strCfg,"","",0)
									VersionList = RTrim(LTrim(Split(vFileLine(0),"=")(1)))
									vVersion = Split(VersionList,",")
									vFileLine(0) = "Version = "
									For nInd = 0 to UBound(vVersion)
									   If vVersion(nInd) <> strVersion Then vFileLine(0) = vFileLine(0) & vVersion(nInd) & ","
									Next
									vFileLine(0) = Left(vFileLine(0),Len(vFileLine(0)) - 1)
									Call DeleteFileGroup(strDirectoryConfig & "\CfgList.txt", strCfg, 0)
									Call AppendStringToFile(strDirectoryConfig & "\CfgList.txt", "[" & strCfg & "]", 0)
									Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vFileLine, UBound(vFileLine),2,0)
									If objFSO.FolderExists(strDirectoryConfig & "\" & strCfg & "\" & strVersion) Then 
										objFSO.DeleteFolder  strDirectoryConfig & "\" & strCfg & "\" & strVersion, True
									End If 	
							End Select
                            g_objIE.Document.All("ButtonHandler").Value = "Reload after Download"																
							Exit Do
						Loop
			
			Case "SAVE_TESTED"
						strCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Text
			            nOptions = g_objIE.document.getElementById("cfg_name").selectedIndex 
						nCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Value
						nVersion = g_objIE.document.getElementById("Input_Param_2").selectedIndex  
						strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).Text
						g_objIE.Document.All("ButtonHandler").Value = "DOWNLOAD"
			Case "SAVE_AS"
						g_objIE.Document.All("ButtonHandler").Value = "Do Nothing"
						Do
							Call IE_Hide(g_objIE)
							nResult = IE_PromptForSettings(vIE_Scale, 6, vSettings, vSessionCRT, vPlatforms, nDebug)
							Select Case nResult
								Case 0, -1
									Call IE_Unhide(g_objIE)
									Exit Do
								Case Else
									Call IE_Unhide(g_objIE)
									strCfg = vSettings(26)
							End Select										
							Call CreateNewCfg(strCfg, nCfg, strVersion,strDirectoryConfig, vCfgInventory, vCfgList, nDebug)
							nYear = 0
							nTag = 0
							nVersion = 0
							CurrentCfg = strCfg
							YES_NO = True
							g_objIE.Document.All("ButtonHandler").Value = "DOWNLOAD"
							Exit Do
						Loop
			Case "DOWNLOAD" ' Save Tested Config
						g_objIE.Document.All("ButtonHandler").Value = "None"			
			            Do
    						If Not SecureCRT_Installed Then 
							   Exit Do
    						End If
							If CurrentCfg <> "Null" and CurrentCfg <> strCfg Then
						        If Not YES_NO Then
									vvMsg(0,0) = "ATTENTION!" 	: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
									vvMsg(1,0) = "The name of the loaded configuration is different from the name you use to save it now:"  : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1			
									vvMsg(2,0) = " "       	  : vvMsg(2,1) = "normal"   : vvMsg(2,2) = HttpTextColor1
									vvMsg(3,0) = "New Name: " & strCfg  	  : vvMsg(3,1) = "normal"   : vvMsg(3,2) = HttpTextColor1
									vvMsg(4,0) = "Old Name: " & CurrentCfg    : vvMsg(4,1) = "normal" : vvMsg(4,2) = HttpTextColor1			
									If Not IE_CONT(vIE_Scale, "Downloading Final Configuration?", vvMsg, 5, g_objIE, nDebug) Then Exit Do
								End If
							End If
							If Not YES_NO Then
								vvMsg(0,0) = "ATTENTION!" 	                                                     : vvMsg(0,1) = "bold" : vvMsg(0,2) =  HttpTextColor1
								vvMsg(1,0) = "Would you like to create a new version of the configuration?"      : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1			
								vvMsg(2,0) = " "                                                                 : vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor1											
								If IE_CONT(vIE_Scale, "Downloading Final Configuration?", vvMsg, 4, g_objIE, nDebug) Then
									'--------------------------------------------
									'   CREARE NEW VERSION OF THE CONFIGURATION
									'--------------------------------------------
									Call CreateNewCfg(strCfg, nCfg, strVersion, strDirectoryConfig,vCfgInventory, vCfgList, nDebug)
									'--------------------------------------
									' UPDATE CFG VERSIONS
									'--------------------------------------
									nVersion = UpdateCfgVer(g_objIE, nCfg, vCfgList, "Input_Param_2")
									strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).text
							    End if 
							End If
							'------------------------------------------------
							'   READ A SESSION LIST TO DOWNLOAD CONFIG FROM
							'------------------------------------------------
							SessionList = ""
							For nInd = 0 to Ubound(vSessionEnable) - 1
							    strFileSettings = strDirectoryWork & "\config\settings.dat"
							    Call FindAndReplaceStrInFile(strFileSettings, "Status " & nInd + 1, vSessionEnable(nInd), 0)
								If InStr(vSessionEnable(nInd),"Enabled") > 0 Then SessionList = SessionList & Split(vSessionCRT(nInd),",")(2) & " "
                            Next
							SessionList = RTrim(SessionList)
							vvMsg(0,0) = "SAVING CONFIGURATION:" 					        	: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
							vvMsg(1,0) = "Configuration Name: ..\" & strCfg             	: vvMsg(1,1) = "normal; font-size 4pt;" : vvMsg(1,2) = HttpTextColor1			
							vvMsg(2,0) = "Version:  " & strVersion                       	: vvMsg(2,1) = "normal"	: vvMsg(2,2) = HttpTextColor1
							vvMsg(3,0) = "Session List:  " & SessionList                   	: vvMsg(3,1) = "normal"	: vvMsg(3,2) = HttpTextColor1
							If IE_CONT(vIE_Scale, "Save configurations?", vvMsg, 4, g_objIE, nDebug) Then 
								'---------------------------------------------------------
								' RUN TELNET SCRIPT
								'---------------------------------------------------------
								strCmd = strCRTexe &_ 
									" /ARG " & strCfg &_
									" /ARG " & strVersion &_
									" /ARG " & strFileSettings &_	
									" /ARG " & strDirectoryWork
								For i = 0 to UBound(Split(SessionList," "))
								    strCmd = strCmd & " /ARG " & Split(SessionList," ")(i)
								Next
								strCmd = strCmd & " /SCRIPT " & strDirectoryWork & "\" & VBScript_DNLD_Config
                                Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)														
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strCfg, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strVersion, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strFileSettings, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork, "", objDebug, MAX_LEN, 1, 1)
								For i = 0 to UBound(Split(SessionList," "))
								    Call TrDebug ("IE_PromptForInput: " & " /ARG " & Split(SessionList," ")(i), "", objDebug, MAX_LEN, 1, 1)
								Next
								Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & strDirectoryWork & "\" & VBScript_DNLD_Config, "",objDebug, MAX_LEN, 1, 1)						
								g_objShell.run strCmd, nWindowState, True
								CurrentCfg = strCfg
								CurrentVer = strVer
							    CFG_Downloaded = True
								g_objIE.Document.All("ButtonHandler").Value = "Reload after Download"								
							End If	
						    Exit Do
						Loop
						YES_NO = False
			Case "BLK_DOWNLOAD" 
						g_objIE.Document.All("ButtonHandler").Value = "None"			
			            Do
    						If Not SecureCRT_Installed Then 
							   Exit Do
    						End If
							'----------------------------------------------------
							'   CREATE NEW CONFIGURATION NAME AND FOLDERS FOR IT
							'----------------------------------------------------							
							Call IE_Hide(g_objIE)
							nResult = IE_PromptForSettings(vIE_Scale, 7, vSettings, vSessionCRT, vPlatforms, nDebug)
							Select Case nResult
								Case 0, -1
									Call IE_Unhide(g_objIE)
									Exit Do
								Case Else
									Call IE_Unhide(g_objIE)
									strBulkList = Split(vSettings(25),"=")(1)
							End Select										
							If GetFileLineCountByGroup(strBulkList, vBulkList,"Bulk_Load","","",0) = 0 Then Exit Do
							nYear = 0
							nTag = 0
							nVersion = 0
							'------------------------------------------------
							'   READ A SESSION LIST TO DOWNLOAD CONFIG FROM
							'------------------------------------------------
							SessionList = ""
							For nInd = 0 to Ubound(vSessionEnable) - 1
								strFileSettings = strDirectoryWork & "\config\settings.dat"
								Call FindAndReplaceStrInFile(strFileSettings, "Status " & nInd + 1, vSessionEnable(nInd), 0)
								If InStr(vSessionEnable(nInd),"Enabled") > 0 Then SessionList = SessionList & Split(vSessionCRT(nInd),",")(2) & " "
							Next
							SessionList = RTrim(SessionList)
							vvMsg(0,0) = "SAVING CONFIGURATION:" 					        	: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
							vvMsg(1,0) = "Bulk Download: " & UBound(vBulkList) & " configurations"    	: vvMsg(1,1) = "normal; font-size 4pt;" : vvMsg(1,2) = HttpTextColor1			
							vvMsg(2,0) = "Session List:  " & SessionList                   	: vvMsg(2,1) = "normal"	: vvMsg(2,2) = HttpTextColor1
							If Not IE_CONT(vIE_Scale, "Save configurations?", vvMsg, 4, g_objIE, nDebug) Then Exit Do
							For each strCfg in vBulkList
							    If strCfg = "" Then Exit For
								CurrentCfg = strCfg
								nCfg = GetExactObjectLineNumber(vCfgInventory, UBound(vCfgInventory),strCfg)
								' Call CreateNewCfg(strCfg, nCfg, strVersion, strDirectoryConfig,vCfgInventory, vCfgList, nDebug)
								'---------------------------------------------------------
								' RUN TELNET SCRIPT
								'---------------------------------------------------------
								strCmd = strCRTexe &_ 
									" /ARG " & strCfg &_
									" /ARG " & strVersion &_
									" /ARG " & strFileSettings &_	
									" /ARG " & strDirectoryWork
								For i = 0 to UBound(Split(SessionList," "))
								    strCmd = strCmd & " /ARG " & Split(SessionList," ")(i)
								Next
								strCmd = strCmd & " /SCRIPT " & strDirectoryWork & "\" & VBScript_BLK_DNLD_Config
                                Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)														
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strCfg, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strVersion, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strFileSettings, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork, "", objDebug, MAX_LEN, 1, 1)
								For i = 0 to UBound(Split(SessionList," "))
								    Call TrDebug ("IE_PromptForInput: " & " /ARG " & Split(SessionList," ")(i), "", objDebug, MAX_LEN, 1, 1)
								Next
								Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & strDirectoryWork & "\" & VBScript_BLK_DNLD_Config, "",objDebug, MAX_LEN, 1, 1)						
								g_objShell.run strCmd, nWindowState, True
								CurrentCfg = strCfg
								CurrentVer = strVer
							    CFG_Downloaded = True
                            Next								
						    Exit Do
						Loop
						g_objIE.Document.All("ButtonHandler").Value = "Reload after Download"						
						YES_NO = False
			Case "LOAD"
		                g_objIE.Document.All("ButtonHandler").Value = "None"
						Do
				            If Not SecureCRT_Installed Then 
							   Exit Do
    						End If 
							If Not CFG_Downloaded and CurrentCfg <> "Null" Then 
								vvMsg(0,0) = "DO YOU WANT TO SAVE"  : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
								vvMsg(1,0) = "Current Configuration ? "  : vvMsg(1,1) = "normal" : vvMsg(1,2) =  HttpTextColor2
								vvMsg(2,0) = "Configuration: " & CurrentCfg   : vvMsg(2,1) = "bold"     : vvMsg(2,2) =  HttpTextColor2
								If IE_CONT(vIE_Scale, "Continue?", vvMsg,3, g_objIE, nDebug) Then 
									g_objIE.Document.All("ButtonHandler").Value = "DOWNLOAD"
									YES_NO = True
									Exit Do
								End If
							End If
							nOptions = g_objIE.document.getElementById("cfg_name").selectedIndex 
							nCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Value
							strCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Text
							nVersion = g_objIE.document.getElementById("Input_Param_2").selectedIndex  
							strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).Text
							SessionList = ""
							For nInd = 0 to Ubound(vSessionEnable) - 1
							    strFileSettings = strDirectoryWork & "\config\settings.dat"
							    Call FindAndReplaceStrInFile(strFileSettings, "Status " & nInd + 1, vSessionEnable(nInd), 0)
								If InStr(vSessionEnable(nInd),"Enabled") > 0 Then SessionList = SessionList & Split(vSessionCRT(nInd),",")(2) & " "
                            Next
							SessionList = RTrim(SessionList)
							vvMsg(0,0) = "LOAD CONFIGURATION:" 					        	: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
							vvMsg(1,0) = "Configuration Name: ..\" & strCfg             	: vvMsg(1,1) = "normal; font-size 7;" : vvMsg(1,2) = HttpTextColor2			
							vvMsg(2,0) = "Version:  " & strVersion                       	: vvMsg(2,1) = "normal"	: vvMsg(2,2) = HttpTextColor2
							vvMsg(3,0) = "Session List:  " & SessionList                   	: vvMsg(3,1) = "normal"	: vvMsg(3,2) = HttpTextColor2
							If IE_CONT(vIE_Scale, "Load new configurations?", vvMsg, 4, g_objIE, nDebug) Then 
								'---------------------------------------------------------
								' RUN TELNET SCRIPT
								'---------------------------------------------------------
								strCmd = strCRTexe &_ 
									" /ARG " & strCfg &_
									" /ARG " & strVersion &_
									" /ARG " & strFileSettings &_	
									" /ARG " & strDirectoryWork
								For i = 0 to UBound(Split(SessionList," "))
								    strCmd = strCmd & " /ARG " & Split(SessionList," ")(i)
								Next
								strCmd = strCmd & " /SCRIPT " & strDirectoryWork & "\" & VBScript_Upload_Config
                                Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)														
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strCfg, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strVersion, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strFileSettings, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork, "", objDebug, MAX_LEN, 1, 1)
								For i = 0 to UBound(Split(SessionList," "))
								    Call TrDebug ("IE_PromptForInput: " & " /ARG " & Split(SessionList," ")(i), "", objDebug, MAX_LEN, 1, 1)
								Next
								Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & strDirectoryWork & "\" & VBScript_Upload_Config, "",objDebug, MAX_LEN, 1, 1)						
								g_objShell.run strCmd, nWindowState
								CurrentCfg = strCfg
								CurrentVer = strVer
                                CFG_Downloaded = False							
							End If	
							Exit Do
						Loop
			Case "UPGRADE_SW"
			                g_objIE.Document.All("ButtonHandler").Value = "None"
							Do 
								Call IE_Hide(g_objIE)
								nResult = IE_PromptForSettings(vIE_Scale, 8, vSettings, vSessionCRT, vPlatforms, nDebug)
								Select Case nResult
									Case 1
                                        strDownloadOnly = "-i"
									Case 3
									    strDownloadOnly = "-d"
									Case Else 										
										Call IE_Unhide(g_objIE)
										Exit Do
								End Select
								SessionList = ""
								Call IE_Unhide(g_objIE)
                                ' Get list of Sessions to be Upgraded
								For nInd = 0 to Ubound(vSessionEnable) - 1
									If InStr(vSessionEnable(nInd),"Enabled") > 0 Then SessionList = SessionList & Split(vSessionCRT(nInd),",")(2) & " "
								Next
								SessionList = RTrim(SessionList)
								vvMsg(0,0) = "UPGRADE JUNOS IMAGE:"					: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
								vvMsg(1,0) = "Following nodes will be upgraded: "   : vvMsg(1,1) = "normal; font-size 7;" : vvMsg(1,2) = HttpTextColor2			
								vvMsg(2,0) = SessionList                   	        : vvMsg(2,1) = "normal"	: vvMsg(2,2) = HttpTextColor2
								vvMsg(3,0) = ""                   	                : vvMsg(3,1) = "normal"	: vvMsg(3,2) = HttpTextColor2
								If Not IE_CONT(vIE_Scale, "Load new configurations?", vvMsg, 3, g_objIE, nDebug) Then Exit Do
								'---------------------------------------------------------
								' ENTER LOGIN AND PASSWORD TO ACCESS CATALOG
								'---------------------------------------------------------
								strLogin = "vmukhin"
								strPassword = ""
								vvMsg(0,0) = "Login and Password for"			 : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
								vvMsg(1,0) = "Junos Image Catalogue: "           : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2
								vvMsg(2,0) = "Hint: Use your UNIX credentials"   : vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor2
								Call IE_PromptLoginPassword (objParentWin, vIE_Scale, vvMsg, 3, strLogin, strPassword, True, nDebug )
								'---------------------------------------------------------
								' RUN TELNET SCRIPT
								'---------------------------------------------------------
								For nInd = 0 to Ubound(vSessionEnable) - 1
									If InStr(vSessionEnable(nInd),"Enabled") > 0 Then 
										strCmd = strCRTexe &_ 
											" /ARG " & strDirectoryWork & "\config\class_catalog.dat" &_	
											" /ARG " & strDirectoryWork & "\config\settings.dat" &_
											" /ARG " & strDirectoryWork &_
											" /ARG " & nInd &_
											" /ARG " & strLogin &_
											" /ARG " & strPassword &_
											" /ARG " & strDownloadOnly
										strCmd = strCmd & " /SCRIPT " & strDirectoryWork & "\" & VBScript_UPDATE_Junos
										Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)														
										Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork & "\config\class_catalog.dat", "", objDebug, MAX_LEN, 1, 1)						
										Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork & "\config\settings.dat" , "", objDebug, MAX_LEN, 1, 1)						
										Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork, "", objDebug, MAX_LEN, 1, 1)						
										Call TrDebug ("IE_PromptForInput: " & " /ARG " & nInd, "", objDebug, MAX_LEN, 1, 1)
										Call TrDebug ("IE_PromptForInput: " & " /ARG " & strLogin, "", objDebug, MAX_LEN, 1, 1)
										Call TrDebug ("IE_PromptForInput: " & " /ARG " & strPassword, "", objDebug, MAX_LEN, 1, 1)
										Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & strDirectoryWork & "\" & VBScript_UPDATE_Junos, "",objDebug, MAX_LEN, 1, 1)						
								        g_objShell.run strCmd, nWindowState, True
									End If
								Next
							    Exit Do
						    Loop								
			Case "EDIT"
						Do
							nOptions = g_objIE.document.getElementById("cfg_name").selectedIndex 
							nCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Value
							strCfg = g_objIE.document.getElementById("cfg_name").options(nOptions).Text
							nVersion = g_objIE.document.getElementById("Input_Param_2").selectedIndex  
							strVersion = g_objIE.document.getElementById("Input_Param_2").options(nVersion).Text
							Tsys0 = DateDiff("n",D0,Date() & " " & Time()) 
							'---------------------------------------------------------
							' OPEN CONFIGURATION FILES WITH TEXT EDITOR
							'---------------------------------------------------------
							g_objShell.Run "Explorer.exe" & " " & strDirectoryConfig & "\" & strCfg & "\" & strVersion	
'							g_objShell.Run strEditor & " " & strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileL	
'							g_objShell.Run strEditor & " "  & strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileR
						Exit Do
						Loop
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Reload and Download"
					    vSessionTmp(0) = nYear
						vSessionTmp(1) = nTag
						vSessionTmp(2) = nCfg
						vSessionTmp(3) = nVersion
						vSessionTmp(4) = CurrentCfg
						vSessionTmp(5) = 1
						vSessionTmp(6) = "DOWNLOAD"
                        IE_PromptForInput = 1
						g_objIE.Quit
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						exit function						
			Case "Reload after Download"
					    vSessionTmp(0) = nYear
						vSessionTmp(1) = nTag
						vSessionTmp(2) = nCfg
						vSessionTmp(3) = nVersion
						vSessionTmp(4) = CurrentCfg
						vSessionTmp(5) = 0
						vSessionTmp(6) = "Do Nothing"
                        IE_PromptForInput = 1
						g_objIE.Quit
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						exit function						
			Case "Cancel"
					    vSessionTmp(0) = nYear
						vSessionTmp(1) = nTag
						vSessionTmp(2) = nCfg
						vSessionTmp(3) = nVersion
						vSessionTmp(4) = CurrentCfg
						vSessionTmp(5) = 0		
						vSessionTmp(6) = "Do Nothing"
						Call WriteArrayToFile(strDirectoryTmp & "\" & strFileSessionTmp,vSessionTmp,UBound(vSessionTmp),1,0)
'						Call WriteStrToFile(strDirectoryTmp & "\" & strFileSessionTmp, nService, 1, 1, 0)
'						Call WriteStrToFile(strDirectoryTmp & "\" & strFileSessionTmp, nFlavor, 2, 1, 0)
'						Call WriteStrToFile(strDirectoryTmp & "\" & strFileSessionTmp, nTaskInd, 3, 1, 0)
						IE_PromptForInput = 0
						g_objIE.Quit
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						exit function
			Case "POPULATE_ORIG"
				vvMsg(0,0) = "WOULD YOU LIKE TO POPULATE :" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
				vvMsg(1,0) = "ALL ORIGINAL CONFIGS" 			    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
				vvMsg(2,0) = "TO TCG XLS TEMPLATES? "           	: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor2			
				If IE_CONT(vIE_Scale, "DownLoad configurations?", vvMsg, 3, g_objIE, nDebug) Then 
					g_objIE.Document.All("ButtonHandler").Value = "None" ' <-- If You need to activate action use POPULATE_ALL instead of None
					SourceCfgFolder = strDirectoryConfig
				Else 
					g_objIE.Document.All("ButtonHandler").Value = "None"
				End If
			Case "POPULATE_DNLD"
				vvMsg(0,0) = "WOULD YOU LIKE TO POPULATE :" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
				vvMsg(1,0) = "ALL DOWNLOADED CONFIGS" 			    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
				vvMsg(2,0) = "TO TCG XLS TEMPLATES? "           	: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor2			
				If IE_CONT(vIE_Scale, "DownLoad configurations?", vvMsg, 3, g_objIE, nDebug) Then 
					SourceCfgFolder = strDirectoryConfig & "\Tested"
					g_objIE.Document.All("ButtonHandler").Value = "None" ' <-- If You need to activate action use POPULATE_ALL instead of None
				Else 
					g_objIE.Document.All("ButtonHandler").Value = "None"
				End If
			case "POPULATE_ALL"
				Do
					nNewService = nService
					nNewFlavor = nFlavor
					nNewTaskInd = nTaskIndex
					nNewTask = nTask
				    g_objIE.Document.All("ButtonHandler").Value = "None"
					On Error Resume Next
					Err.Clear
						set objXLS = CreateObject("Excel.Application")				
						if Err.Number <> 0 then
							vvMsg(0,0) = "CAN'T POPULATE CONFIGURATION FILES TO TCG TEMPLATE" 		         : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  "Red"
							vvMsg(1,0) = "Make sure that MS Excel Application is installed on the system "   : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
							vvMsg(2,0) = "Skip config export... "                                            : vvMsg(2,1) = "normal"      : vvMsg(2,2) = HttpTextColor2
                            vvMsg(3,0) = ""							
						   Call IE_MSG(vIE_Scale, "Error",vvMsg,4, g_objIE)
						   Exit Do
						End If 
					On Error Goto 0  
					
					nIndex = 0					
					Do
						On Error Resume Next
							Err.Clear
							Set objMonitor = objFSO.OpenTextFile(strDirectoryWork & "\Log\" & TCG_MONITOR & "_" & nIndex & ".log",ForWriting,True)
							Select Case Err.Number
								Case 0
								   Exit Do
								Case Else
									nIndex = nIndex + 1						
									If nIndex > 2 Then
									   Call TrDebug ("IE_PromptForInput: WAS UNABLE TO CREATE MONITOR LOG FILE FOR TCG EXPORT PROCEDURE","", objDebug, MAX_LEN, 3, 1)
									   Exit Do
									End If
									wscript.sleep 300
							End Select
						On Error goto 0					
					Loop
					strMonitorFile = strDirectoryWork & "\Log\" & TCG_MONITOR & "_" & nIndex & ".log"
					objXLS.visible = True
					Set objFile = CreateObject("Scripting.FileSystemObject")
						For nService = 0 to Ubound(vSvc,1) - 1
							Set objFolder = objFSO.GetFolder(strTempOrigFolder)
							Set colFiles = objFolder.Files
							For Each objFile in colFiles
								If InStr(LCase(objFile.Name) ,LCase(vSvc(nService,1))) and InStr(objFile.Name,"xlsx") Then strWorkBook = objFile.Name End If
							Next					
		'					For nInd = 0 to 3
		'					   If InStr(LCase(vTemplates(nInd)) ,LCase(vSvc(nService,1))) Then strWorkBook = vTemplates(nInd)
		'					Next
							Call TrDebug ("IE_PromptForInput: OPEN WorkBook: " & strTempOrigFolder & "\" & strWorkBook,"", objDebug, MAX_LEN, 3, 1)
							Call TrDebug ("EXPORT CFG: OPEN WorkBook: " & strTempOrigFolder & "\" & strWorkBook,"", objMonitor, MAX_LEN, 3, 1)												
							Set objWrkBk = objXLS.Workbooks.open(strTempOrigFolder & "\" & strWorkBook)
							For nFlavor = 0 to vSvc(nService,0) - 1
								StartRow = 0
								For nTaskIndex = 0 to Ubound(Split(vFlavors(nService, nFlavor, 1),","))
									strXLSheet = Split(vXLSheetPrefix(nService),",")(nFlavor) & " CFG " & Split(vFlavors(nService, nFlavor, 1),",")(nTaskIndex)
									Call TrDebug ("IE_PromptForInput: OPEN XLSheet: " & strXLSheet,"", objDebug, MAX_LEN, 1, 1)
									Set objXLSeet = objWrkBk.Worksheets(strXLSheet)
									Set objCell = objXLSeet.Cells.Find("3/ ",,,,1,2)
									StartRow = objCell.Row
									Set objCell = objXLSeet.Cells.Find("Copy / Paste ",objCell,,,1,2)
							'		Set objCell = objXLSeet.Cells(objCell.Row,objCell.Column + 1)
							'		Do While i<30
							'			Set objCell = objXLSeet.Cells(objCell.Row,objCell.Column + i)
							'			if objXLSheet.Cells.Value
							'		Loop
									StartCol = objCell.Column
							'		MsgBox StartRow & ", " & StartCol & ", " 							
									Call TrDebug ("IE_PromptForInput: LOOKING FOR Start Row: " & StartRow,"", objDebug, MAX_LEN, 1, 1)
									strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & Split(vFlavors(nService, nFlavor, 1),",")(nTaskIndex) & "-" & Platform & "-l.conf"
									strConfigFileR = vFlavors(nService, nFlavor,0) & "-" & Split(vFlavors(nService, nFlavor, 1),",")(nTaskIndex) & "-" & Platform & "-r.conf"
									'---------------------------------------------------------
									' CHECK IF CONFIGURATION FILE LEFT EXIST. COPY TO BACK FOLDER
									'---------------------------------------------------------
									Skip = 0
									If Not objFSO.FileExists(SourceCfgFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL) Then 
										Call TrDebug ("EXPORT CONFIGURATION: " & vSvc(nService,1) & "\" & strConfigFileL, "NOT FOUND" , objDebug, MAX_LEN, 1, 1)						
										Call TrDebug ("EXPORT CFG: " & vSvc(nService,1) & "\" & strConfigFileL, "NOT FOUND" , objMonitor, MAX_LEN, 1, 1)																
										Skip = 1
									Else 
										Call TrDebug ("EXPORT CONFIGURATION: " & vSvc(nService,1) & "\" & strConfigFileL, "FOUND" , objDebug, MAX_LEN, 1, 1)						
									End If
									'---------------------------------------------------------
									' CHECK IF CONFIGURATION FILE RIGHT EXIST. COPY TO BACK FOLDER
									'---------------------------------------------------------
									If Not objFSO.FileExists(SourceCfgFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR) Then 
										Call TrDebug ("EXPORT CONFIGURATION: " & vSvc(nService,1) & "\" & strConfigFileR, "NOT FOUND" , objDebug, MAX_LEN, 1, 1)
										Call TrDebug ("EXPORT CFG: " & vSvc(nService,1) & "\" & strConfigFileR, "NOT FOUND" , objMonitor, MAX_LEN, 1, 1)																
										Skip = 1
									Else
										Call TrDebug ("EXPORT CONFIGURATION: " & vSvc(nService,1) & "\" & strConfigFileR, "FOUND" , objDebug, MAX_LEN, 1, 1)						
									End If
									If Skip = 0 Then
										nIndex = StartRow + 3
									'---------------------------------------------------------
									' POPULATE CONFIG FOR LEFT NODE
									'---------------------------------------------------------
										nSession = GetFileLineCountSelect(SourceCfgFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL, vSession,"NULL","NULL","NULL",0)
										nCount = 0
										Do While nCount < nSession
											objXLSeet.Cells(nIndex,1) = vSession(nCount)
											nIndex = nIndex + 1 
											nCount = nCount + 1
										Loop
										Call TrDebug ("EXPORT CONFIGURATION: " & vSvc(nService,1) & "\" & strConfigFileL, "POPULATED" , objDebug, MAX_LEN, 1, 1)
										Call TrDebug ("EXPORT CFG: " & vSvc(nService,1) & "\" & strConfigFileL, "OK" , objMonitor, MAX_LEN, 1, 1)						
										Redim vSession(0)
									'---------------------------------------------------------
									' POPULATE CONFIG FOR RIGHT NODE
									'---------------------------------------------------------
										nSession = GetFileLineCountSelect(SourceCfgFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR, vSession,"NULL","NULL","NULL",0)
										nIndex = StartRow + 3
										nCount = 0
										Do While nCount < nSession
											objXLSeet.Cells(nIndex,StartCol) = vSession(nCount)
											nIndex = nIndex + 1 : nCount = nCount + 1
										Loop
										Redim vSession(0)
										Call TrDebug ("EXPORT CONFIGURATION: " & vSvc(nService,1) & "\" & strConfigFileR, "POPULATED" , objDebug, MAX_LEN, 1, 1)
										Call TrDebug ("EXPORT CFG: " & vSvc(nService,1) & "\" & strConfigFileR, "OK" , objMonitor, MAX_LEN, 1, 1)
									End If
									Set objXLSeet = Nothing
									StartRow = 0 
								Next
							Next
							objWrkBk.Application.DisplayAlerts = False
							objWrkBk.SaveAs(strTempDestFolder & "\" & strWorkBook)
							Call TrDebug ("EXPORT CFG: Saving WorkBook As: " & strTempDestFolder & "\" & strWorkBook, "OK" , objMonitor, MAX_LEN, 1, 1)
							' objWrkBk.Application.DisplayAlerts = True
							objWrkBk.close
							Set objWrkBk = Nothing
						Next
						objXLS.Quit
						vvMsg(0,0) = "CONFIGURATION FILES HAS BEEN EXPORTED SUCCESSFULLY " 		         : vvMsg(0,1) = "bold" 	: vvMsg(0,2) = HttpTextColor1
                        vvMsg(1,0) = ""							
  						Call IE_MSG(vIE_Scale, "Error",vvMsg,2, g_objIE)
						If IsObject(objMonitor) Then Close(objMonitor)
    					If objFSO.FileExists(strMonitorFile) Then
						    g_objShell.run "notepad.exe " &  strMonitorFile, 1, False
						End If 
					Exit Do
				Loop
				nService = nNewService
				nFlavor = nNewFlavor
				nTaskIndex = nNewTaskInd
				nTask = nNewTask
				set objXLS = Nothing
		End Select
		WScript.Sleep 300
    Loop
End Function
'------------------------------------------------
'    SETTINGS DIALOG FORM 
'------------------------------------------------
Function IE_PromptForSettings(ByRef vIE_Scale, MenuID, byRef vSettings, byRef vSessionCRT, byRef vPlatforms, nDebug)
	Dim g_objIE, g_objShell, objShellApp, objFSO
	Dim nInd
	Dim nRatioX, nRatioY, nFontSize_10, nFontSize_12, nButtonX, nButtonY, nA, nB, vOld_Settings, vOld_SessionCRT, vSessionCRT_to_file
    Dim intX
    Dim intY
	Dim nCount
	Dim strLogin
	Dim IE_Menu_Bar
	Dim IE_Border
	Dim nLine, nService, nFlavor, nTask, nPlatform
	Dim vvMsg(8,3)
	Dim objFile, objCfgFile
	Dim objWMIService, IPConfigSet, vTitle
	Const MAX_PARAM = 40
	Const MAX_BW_PROFILES = 30
	Const N_SELECT = 5
	Dim objFolder, objForm, colFiles, strFile, objDialog
	Call TrDebug ("IE_PromptForInput: OPEN MAIN CONFIG LOADER FORM ", "", objDebug, MAX_LEN, 3, nDebug)	
	Set objForm = CreateObject("Shell.Application")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Redim vOld_Settings(Ubound(vSettings))
	Redim vOld_SessionCRT(1)
	Redim vSessionCRT_to_file(1)
	Set g_objIE = Nothing
    Set g_objShell = Nothing
	vTitle = Array("Platform under test",_
	               "Connectivity Settings",_
				   "Folder Settings",_
				   "SecureCRT Sessions",_
				   "Advanced Settings",_
				   "Create New Configuration Name",_
				   "Browse for file with list of configurations",_
				   "Junos Images Catalog")
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	'----------------------------------------
	' IE EXPLORER OBJECTS
	'----------------------------------------
	Set g_objShell = WScript.CreateObject("WScript.Shell")
    Call Set_IE_obj (g_objIE)
    g_objIE.Offline = True
    g_objIE.navigate "about:blank"
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	'-------------------------------------------------------------------------------------------
	'  READ LOCAL IPCONFIG
	'-------------------------------------------------------------------------------------------
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	'----------------------------------------
	' MAIN VARIABLES OF THE GUI FORM
	'----------------------------------------
	If nRatioX > 1 Then nRatioX = 1 : nRatioY = 1 End If
	Select Case nRatioX
		Case 1
				DiagramFigure = strDirectoryWork & "\Data\TestBed001.png"
		Case 1600/1920
				DiagramFigure = strDirectoryWork & "\Data\TestBed002.png"
		Case else
				DiagramFigure = strDirectoryWork & "\Data\TestBed002.png"
				nRatioX = 1600/1920
				nRatioX = 900/1080
	End Select
	SettingsFigure = strDirectoryWork & "\Data\Settings-icon-6.png"
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	nBottom = Round(10 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellH = Round(24 * nRatioY,0)
	LoginTitleW = Round(800 * nRatioX,0)
	nLeft = Round(20 * nRatioX,0)
	nTab = Round(40 * nRatioX,0)
	CellW = LoginTitleW
	LoginTitleH = Round(40 * nRatioY,0)
	nSaveW = nLeft + nButtonX
	nScoreW = 3 * nSaveW
	nColumn = Int(LoginTitleW/3)	
	nNameW = Int((LoginTitleH - nColumn)/3)
	'------------------------------------------
	'	GET NUMBER OF TASKS LINES
	'------------------------------------------	
	Select Case MenuID
	    Case 0 
	      nLine = 26
	    Case 1
	      nLine = 5
	    Case 2
	      nLine = 6
	    Case 3
	      nLine = 8
	    Case 4 
	      nLine = 5 + Ubound(vSessionCRT)
	    Case 5 
	      nLine = 4
		Case 6
		  nLine = 4
		Case 7
		  nLine = 4
		Case 8
			ClassName = "JunosSW"
'			Call SetMyObject(objMain,"JunosSW",nDebug)
'			Call SetMyObject(objMinor,"Release",nDebug)
			nLine = 5 + int(UBound(objMain,1) * N_SELECT * 3/4)
'			MsgBox "pIndex: " & pIndex(1,"Minor List")
'			MsgBox GetVariable("ListNumber" & pIndex(1,"Minor List") + 1, vClass, 2, 1, 0, nDebug)
'           MsgBox objMain(1,pIndex(0,"ImageTemplate")) & ", " & UBound(objMain,1) & ", " & vClass(0,0)
	End Select
		
	WindowH = IE_Menu_Bar + 2 * LoginTitleH + cellH * (nLine) + nBottom
	WindowW = IE_Border + LoginTitleW
	If WindowW < 300 then WindowW = 300 End If

	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)	
   	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "None " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
    g_objIE.height = WindowH
    g_objIE.width = WindowW  
    g_objIE.document.Title = "Lab Configuration Loader Settings"
	g_objIE.Top = (intY - g_objIE.height)/2
	g_objIE.Left = (intX - g_objIE.width)/2
	g_objIE.Visible = True		
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title
    '-----------------------------------------------------------------
	' SET THE TITLE OF THE  FORM   		
	'-----------------------------------------------------------------
	nLine = 0
	    htmlMain = 	"<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>" &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: " & HttpBgColor5 & "; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""" &_
			"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW & """>" & _
				"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
				";font-weight: normal;font-style: italic;"">"&_
				"&nbsp;&nbsp;" & vTitle(MenuID - 1) & " <span style=""font-weight: bold;""></span></span></p>"&_
			"</td>" &_
		"</tr></tbody></table>"
	nLine = nLine +1
	Select Case MenuID
	    Case 1
				'-----------------------------------------------------------------
				' PLATFORM TITLE
				'-----------------------------------------------------------------
				cTitle = "Platform under test"
				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>" 
				nLine = nLine + 1
				'-----------------------------------------------------------------
				' PLATFORM PARAMETERS:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
				strType = "text"
				htmlMain = htmlMain &_
					"<tr>"&_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & Split(vSettings(13),"=")(0) & "</p>" &_
						"</td>"&_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
						"</td>" &_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
							"<select name='Platform_Name' id='Platform_Name'" &_
								"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
								"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
								" onchange=document.all('ButtonHandler').value='SelectPlatform';" &_
								"type=text > "
								For nPlatform = 0 to Ubound(vPlatforms) - 1
									htmlMain = htmlMain &_
														"<option value=" & nPlatform & """>" & Split(vPlatforms(nPlatform),",")(0) & "</option>" 
								Next
								htmlMain = htmlMain &_
							"<option value=" & Ubound(vPlatforms) & """>" & Space_html(24) & "</option>" &_
							"</select>" &_
						"</td>" &_
					"</tr>"
				htmlMain = htmlMain &_
					"<tr>"&_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & Split(vSettings(14),"=")(0) & "</p>" &_
						"</td>"&_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_
							"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
							";font-weight: normal;font-style: italic;"">"&_
							"&nbsp;&nbsp;Config. name: [Service Type]-[TC#]-<span style=""font-weight: bold;"">[Prefix]</span>-[L|R].conf</span></p>"&_
						"</td>" &_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
							"<input name=Platform_Index value='' style=""text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=15 " &_
							"type=" & strType & " > " &_
						"</td>" &_
					"</tr>"

				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 2
		Case 2
					'-----------------------------------------------------------------
					' CONNECTIVITY SETTINGS TITLE
					'-----------------------------------------------------------------
					cTitle = "Connectivity Settings"
					htmlMain = htmlMain &_
						"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
						"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
						"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
							"<tbody>"
				'-----------------------------------------------------------------
				' SETTINGS PARAMETERS:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
				For nSetting = 2 to 4
					nLine = nLine + 1
					strType = "text"
					If InStr(Split(vSettings(nSetting),"=")(0),"assword") Then strType = "password" End If
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & Split(vSettings(nSetting),"=")(0) & "</p>" &_
							"</td>"
					Select Case nSetting
						Case 2
								htmlMain = htmlMain &_
										"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
											"<select name='Adapter_Name' id='Adapter_Name'" &_
											"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
											"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
											"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
											" onchange=document.all('ButtonHandler').value='SelectAdapter';" &_
											"type=text > "
											nAdapter = 0
											For Each IPConfig in IPConfigSet
												htmlMain = htmlMain &	"<option value=" & IPConfig.IPAddress(0) & ">" & IPConfig.Description & "</option>" 
												nAdapter = nAdapter + 1	
											Next
												htmlMain = htmlMain &	"<option value=127.0.0.1>Loopback</option>"
												nAdapter = nAdapter + 1										
											htmlMain = htmlMain & "</select>" &_
										"</td>"
						Case Else
								htmlMain = htmlMain &_
										"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
										"</td>"
					End Select
					htmlMain = htmlMain &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=18 maxlength=18 " &_
								"type=" & strType & " > " &_
							"</td>" &_
						"</tr>" 
				Next
				ButtonDisabled = ""
				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 2
		Case 3	
				'-----------------------------------------------------------------
				' FOLDERS TITLE
				'-----------------------------------------------------------------
				cTitle = "Folder Settings"
				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>"
				'-----------------------------------------------------------------
				' FOLDERS PARAMETERS:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
				For nSetting = 5 to 9
					nLine = nLine + 1
					strType = "text"
					BgTextColor = HttpBgColor4 : ButtonDisabled = ""
					If nSetting = 5 Then ButtonDisabled = "disabled" : BgTextColor = HttpBgColor1  End If			
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & Split(vSettings(nSetting),"=")(0) & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & BgTextColor & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
								"type=" & strType & " > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
								"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & 2 * nButtonX &_
								";font-size: " & nFontSize_10 &".0pt;" &_
								";height:" & Int(nButtonY/2) &_
								"px' name='Edit_Folder'" & nSetting & " onclick=document.all('ButtonHandler').value='Folder_" & nSetting & "'; " & ButtonDisabled & ">Edit Folder</button>" & _	
							"</td>" &_
						"</tr>"
				Next
				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 2
		Case 4
				'-----------------------------------------------------------------
				' SECURECRT SESSIONS TITLE
				'-----------------------------------------------------------------
				cTitle = "SecureCRT Sessions"
				strTitleCell = 	"<td style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
								";"" align=""right"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """>" & _
									"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
									";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;&nbsp;&nbsp;"

				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>" &_
							"<tr>" & _
								"<td style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
								";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"</td>" &_
								strTitleCell & "Session Name" & "</p></td>" &_
								strTitleCell & "MNG IP" & "</p></td>" &_
								strTitleCell & "Platform Type" & "</p></td>" &_
								strTitleCell & "Session Login" & "</p></td>" &_
							"</tr>"
				'-----------------------------------------------------------------
				' SECURECRT SESSIONS PARAMETERS:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
			    strEmptyCell = "</td><td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">"
				For nSetting = 0 to UBound(vSessionCRT) - 1
					nLine = nLine + 1
					BgTextColor = HttpBgColor1
					strDisabled = "disabled"
					If nSetting = 0 Then 
					    strDisabled = ""  
						BgTextColor = HttpBgColor4
					End If			
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;Session&nbsp; " & nSetting + 1 & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name='Session_Param_2' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name='Session_Param_0' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name='Session_Param_3' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text> " &_
							"</td>" &_							
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name='Session_Param_4' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
						"</tr>"
				Next
				htmlMain = htmlMain &_				
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & HIDE_CRT & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
									"<input type=checkbox name='Hide_CRT' style=""color: " & HttpTextColor2 & ";""" & _
									" onclick=document.all('ButtonHandler').value='HIDE_CRT';" &_
									"value='Display'>" &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
							"</td>" &_
						"</tr>" &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;Session Folder</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name='Session_Param_1' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text> " &_
							strEmptyCell & strEmptyCell & strEmptyCell &_
						"</tr>"
						

				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 3
		Case 5
				'-----------------------------------------------------------------
				' LAB CONFIGURATION PARAMETERS TITLE
				'-----------------------------------------------------------------
				cTitle = "Advanced Settings"
				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>" 
				'-----------------------------------------------------------------
				' LAB CONFIGURATION PARAMETERS
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
				For nSetting = 12 to 12
					nLine = nLine + 1
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & Split(vSettings(nSetting),"=")(0) & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
								"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & 2 * nButtonX &_
								";font-size: " & nFontSize_10 &".0pt;" &_
								";height:" & Int(nButtonY/2) &_
								"px' name='Edit_PARAM' onclick=document.all('ButtonHandler').value='EDIT_PARAM';>Edit File</button>" & _	
							"</td>" &_
						"</tr>"
				Next
				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 2
		Case 6
				'-----------------------------------------------------------------
				' CREATE NEW LAB CONFIGURATION NAME
				'-----------------------------------------------------------------
				cTitle = "Create Ne Configuration Name"
				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>" 
				'-----------------------------------------------------------------
				' LAB CONFIGURATION PARAMETERS
				'-----------------------------------------------------------------
				For nSetting = 26 to 26
					nLine = nLine + 1
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;Enter Configuration Name</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
							"</td>" &_
						"</tr>"
				Next
				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 2
        Case 7
				'-----------------------------------------------------------------
				' TITLE
				'-----------------------------------------------------------------
				cTitle = "Folder Settings"
				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>"
				'-----------------------------------------------------------------
				' FOLDERS PARAMETERS:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
				For nSetting = 25 to 25
					nLine = nLine + 1
					strType = "text"
					BgTextColor = HttpBgColor4 : ButtonDisabled = ""
					If nSetting = 5 Then ButtonDisabled = "disabled" : BgTextColor = HttpBgColor1  End If			
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & Split(vSettings(nSetting),"=")(0) & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & BgTextColor & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
								"type=" & strType & " > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
								"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & 2 * nButtonX &_
								";font-size: " & nFontSize_10 &".0pt;" &_
								";height:" & Int(nButtonY/2) &_
								"px' name='Edit_Folder'" & nSetting & " onclick=document.all('ButtonHandler').value='Folder_" & nSetting & "'; " & ButtonDisabled & ">Edit Folder</button>" & _	
							"</td>" &_
						"</tr>"
				Next
				htmlMain = htmlMain &_
						"</tbody></table>"
				nLine = nLine + 2
		Case 8
				'-----------------------------------------------------------------
				' JUNOS CATALOG TITLE
				'-----------------------------------------------------------------
				strTitleCell = 	"<td style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
								";"" align=""right"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """>" & _
									"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
									";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;&nbsp;&nbsp;"

				htmlMain = htmlMain &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>" &_
							"<tr>" & _
								"<td style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
								";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"</td>" &_
								strTitleCell & "Platform" & "</p></td>" &_
								strTitleCell & "Type" & "</p></td>" &_
								strTitleCell & "Main Release" & "</p></td>" &_
								strTitleCell & "Minor Release" & "</p></td>" &_
							"</tr>"
				'-----------------------------------------------------------------
				' LIST OF JUNOS RELEASE:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
			    strEmptyCell = "</td><td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">"
				For nImage = 0 to UBound(objMain,1) - 1
					nLine = nLine + 1
					BgTextColor = HttpBgColor1
					strDisabled = "disabled"
					htmlMain = htmlMain &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(2 * LoginTitleW/16) & """align=""center"" valign=""top"">" & _
									"<button style='font-weight: Normal; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
									nButtonX & ";height:" & Int(nButtonY/2) & "; font-size: " & nFontSize_12 & ".0pt;" &_
									"px; ' name='ImageStatus' onclick=document.all('ButtonHandler').value='Clear_" & nImage& "';><u>C</u>lear</button>" &_										
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(2 * LoginTitleW/16) & """valign=""top"" align=""center"">" &_									
								"<input name='Image_Param_7' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(4 * LoginTitleW/16) & """valign=""top"" align=""center"">" &_									
								"<input name='Image_Param_9' value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_							
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(4 * LoginTitleW/16) & """valign=""top"" align=""center"">" &_									
									"<select name='Main_Release" & nImage & "' id='Main_Relese" & nImage & "'" &_
									"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
									"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
									"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
									" onchange=document.all('ButtonHandler').value='SelectMain_" & nImage & "';" &_
									"type=text > "
									vMain = Split(objMain(nImage,pIndex(0,"Main List")),",")
									For nMain = 0 to UBound(vMain)
										htmlMain = htmlMain &	"<option value=Main_" & nMain & ">" & vMain(nMain) & "</option>" 
									Next
									htmlMain = htmlMain & "<option value=Main_" & nMain + 1 & ">" & Space_html(30) & "</option>" 
									htmlMain = htmlMain & "</select>" &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(4 * LoginTitleW/16) & """valign=""top"" align=""center"">" &_									
									"<select name='Minor_Release" & nImage & "' id='Minor_Relese" & nImage & "'" &_
									"style=""border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
									"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
									"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='" & N_SELECT & "'" & _
									" onchange=document.all('ButtonHandler').value='SelectMinor_" & nImage & "';" &_
									"type=text > "
									' Split(objMain(0,pIndex(1,"Minor List")),",")(nMinor)
									' For nMinor = 0 to GetVariable("ListNumber" & pIndex(1,"Minor List") + 1, vClass, 2, 1, 0, nDebug)
									For nMinor = 0 to 100
										htmlMain = htmlMain &	"<option value=Minor_" & nMinor & " >" & Space_html(30) & "</option>" 
									Next
									htmlMain = htmlMain & "</select>" &_
							"</td>" &_							
						"</tr>"
				Next
				htmlMain = htmlMain &_
				        "<tr></tr>" &_
				        "<tr>" &_
				    		"<td style="" background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 & "; border-style: Solid;"" class=""oa2"" height=""" & cellH & """ align=""center"">" &_									
									"<input type=checkbox name='DNLD_IMG_ONLY' style=""color: " & HttpTextColor2 & ";""" & _
									" onclick=document.all('ButtonHandler').value='DNLD_IMG_ONLY';" &_
									"value='Display'>" &_
							"</td>" &_
								strTitleCell & "Download Only" & "</p></td>" &_
								strTitleCell & "</p></td>" &_
								strTitleCell & "</p></td>" &_
								strTitleCell  & "</p></td>" &_
						"</tr>" &_
			"</tbody></table>"
			nLine = nLine + 1		
	End Select
	'------------------------------------------------------
	'   EXIT BUTTON
	'------------------------------------------------------
	htmlMain = htmlMain &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: " & HttpBgColor2 &_
		"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" &_
			"<tbody>" & _
				"<tr>" &_
					"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						2 * nButtonX & ";height:" & nButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='SAVE' onclick=document.all('ButtonHandler').value='SAVE';><u>S</u>ave Settings</button>" & _	
					"</td>"

	Select Case MenuID
	    Case 8
	            htmlMain = htmlMain &_
                	"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						2 * nButtonX & ";height:" & nButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='UPDATE_CATALOG' onclick=document.all('ButtonHandler').value='UPDATE_CATALOG';><u>U</u>pdate Catalog</button>" & _	
					"</td>"
		Case Else
        	    htmlMain = htmlMain &_
					"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
					"</td>"
    End Select
	htmlMain = htmlMain &_
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						2 * nButtonX & ";height:" & nButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px;' name='EXIT' onclick=document.all('ButtonHandler').value='Cancel';><u>E</u>xit</button>" & _	
					"</td>" &_
					"</tr></tbody></table>"
	'-----------------------------------------------------------------
	' HTML Form Parameaters
	'-----------------------------------------------------------------
	g_objIE.Document.Body.innerHTML = htmlMain
    g_objIE.MenuBar = False
    g_objIE.StatusBar = False
    g_objIE.AddressBar = False
    g_objIE.Toolbar = False
    ' Wait for the "dialog" to be displayed before we attempt to set any
    ' of the dialog's default values.
	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strMyPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then strMyPID = Split(strLine,""",""")(1)
    Next
    If strMyPID = "" Then Call GetAppPID(strMyPID, "iexplore.exe")
	g_objShell.AppActivate strMyPID
	'-----------------------------------------------------------------
	'	SET DEFAULT PARAMETERS
	'-----------------------------------------------------------------
	Select Case MenuID
	    Case 0
		Case 1
			If Split(vSettings(13),"=")(1) <> "Unknown" Then nIndex = GetObjectLineNumber( vPlatforms, UBound(vPlatforms), Split(vSettings(13),"=")(1)) - 1 Else nIndex = 0 End If
			g_objIE.document.getElementById("Platform_Name").selectedIndex = nIndex
			g_objIE.Document.All("Platform_Index").Value = Split(vSettings(14),"=")(1)
        Case 2
			For nInd = 0 to nAdapter - 1
			    If Split(vSettings(1),"=")(1) = g_objIE.document.getElementById("Adapter_Name").Options(nInd).text Then 
				   g_objIE.document.getElementById("Adapter_Name").selectedindex = nInd
				End If 
			Next
			g_objIE.Document.All("Settings_Param_2").Value = g_objIE.document.getElementById("Adapter_Name").Value
			For nInd = 3 to 4
				g_objIE.Document.All("Settings_Param_" & nInd).Value = Split(vSettings(nInd),"=")(1)
			Next
        Case 3
			For nInd = 5 to 9
				g_objIE.Document.All("Settings_Param_" & nInd).Value = Split(vSettings(nInd),"=")(1)
			Next
        Case 4	
		    If Split(vSettings(0),"=")(1) = "1" Then g_objIE.Document.All("Hide_CRT").Checked = True
			For i = 0 to UBound(vSessionCRT) - 1
				g_objIE.Document.All("Session_Param_0")(i).Value = Split(vSessionCRT(i),",")(0)
				g_objIE.Document.All("Session_Param_1").Value = Split(vSessionCRT(0),",")(1)
                g_objIE.Document.All("Session_Param_2")(i).Value = Split(vSessionCRT(i),",")(2)
                g_objIE.Document.All("Session_Param_3")(i).Value = Split(vSessionCRT(i),",")(3)				
                g_objIE.Document.All("Session_Param_4")(i).Value = Split(vSessionCRT(i),",")(4)
			Next
		Case 5
		    g_objIE.Document.All("Settings_Param_12").Value = Split(vSettings(12),"=")(1) 			
		Case 7
			strFolder = Split(vSettings(25),"=")(1)
			Select Case IsFile(strFolder)
				Case 0 ' Empty path
					strFolder = "C:\Users\vmukhin\Documents\LAB_CONFIGS"
				Case 1 ' File 
					strFolder = objFSO.GetParentFolderName(strFolder)
				Case 2  ' Folder
					' Do nothing
				Case Else 
					strFolder = "C:\Users\vmukhin\Documents\LAB_CONFIGS"
			End Select
		    g_objIE.Document.All("Settings_Param_25").Value = strFolder
		Case 8 
		    For nImage = 0 to UBound(objMain,1) - 1
			    g_objIE.Document.All("Image_Param_7")(nImage).Value = objMain(nImage,pIndex(0,"Platform"))
				g_objIE.Document.All("Image_Param_9")(nImage).Value = objMain(nImage,pIndex(0,"Display_Name"))
				
				strMain = objMain(nImage,pIndex(0,"Main"))
				strMinor = objMain(nImage,pIndex(0,"Minor"))
				strStatus = objMain(nImage,pIndex(0,"Status"))
				' Found line number to highlighting'
				nOptions = 0
				For nInd = 0 to 99
				    If RTrim(Ltrim(g_objIE.document.getElementById("Main_Release" & nImage).Options(nInd).Text)) = "" Then Exit For
					If RTrim(Ltrim(g_objIE.document.getElementById("Main_Release" & nImage).Options(nInd).Text)) = strMain Then nOptions = nInd : Exit For End If
				Next
				g_objIE.document.getElementById("Main_Release" & nImage).SelectedIndex = nOptions
				strMinorName = objMain(nImage,pIndex(0,"Name")) & "-" & g_objIE.document.getElementById("Main_Release" & nImage).Options(nOptions).Text
				For nMinor = 0 to UBound(objMinor,1) - 1
				   If strMinorName = objMinor(nMinor,pIndex(1,"Name")) Then Exit For
				Next
                vMinor = Split(objMinor(nMinor,pIndex(1,"Minor List")),",")
				nInd = 0
				nOptions = 0
				For Each LineItem in vMinor
				   g_objIE.document.getElementById("Minor_Release" & nImage).Options(nInd).Text = LineItem
				   if strMinor = LineItem Then nOptions = nInd
				   nInd = nInd + 1
				Next
				if strStatus = "Active" Then g_objIE.document.getElementById("Minor_Release" & nImage).SelectedIndex = nOptions
			Next		
	End Select
    Do
        WScript.Sleep 100
    Loop While g_objIE.Busy
    
   ' g_objShell.AppActivate g_objIE.document.Title	

	Do
        ' If the user closes the IE window by Alt+F4 or clicking on the 'X'
        ' button, we'll detect that here, and exit the script if necessary.
        On Error Resume Next
			If g_objIE.width <> WindowW Then g_objIE.width = WindowW End If
			If g_objIE.height <> WindowH Then g_objIE.height = WindowH End If
			Err.Clear
            szNothing = g_objIE.Document.All("ButtonHandler").Value
            if Err.Number <> 0 then exit function
        On Error Goto 0    
        ' Check to see which buttons have been clicked, and address each one
        ' as it's clicked.
        Select Case szNothing
		    Case "SelectMain_0", "SelectMain_1","SelectMain_2","SelectMain_3","SelectMain_4","SelectMain_5","SelectMain_6","SelectMain_7","SelectMain_8","SelectMain_9","SelectMain_10"
			            g_objIE.Document.All("ButtonHandler").Value = "Nothing Selected"   
						nImage = Int(Split(szNothing,"_")(1))
						nOptions = g_objIE.document.getElementById("Main_Release" & nImage).SelectedIndex
						strMinorName = objMain(nImage,pIndex(0,"Name")) & "-" & g_objIE.document.getElementById("Main_Release" & nImage).Options(nOptions).Text
						For nMinor = 0 to UBound(objMinor,1) - 1
						   If strMinorName = objMinor(nMinor,pIndex(1,"Name")) Then Exit For
						Next
						vMinor = Split(objMinor(nMinor,pIndex(1,"Minor List")),",")
						nInd = 0
						For nInd = 0 to 99
						    If UBound(vMinor) >= nInd Then
						       g_objIE.document.getElementById("Minor_Release" & nImage).Options(nInd).Text = vMinor(nInd)
						    Else 
							   g_objIE.document.getElementById("Minor_Release" & nImage).Options(nInd).Text = Space(30)
							End If
						Next
						
			Case "SelectPlatform"
			            nPlatform = g_objIE.document.getElementById("Platform_Name").selectedIndex
			            g_objIE.Document.All("Platform_Index").Value = Split(vPlatforms(nPlatform),",")(1)
						g_objIE.Document.All("ButtonHandler").Value = "Nothing Selected"
			Case "SelectAdapter"
			            ' nAdapter = g_objIE.document.getElementById("Adapter_Name").selectedIndex
			            g_objIE.Document.All("Settings_Param_2").Value = g_objIE.document.getElementById("Adapter_Name").Value
						g_objIE.Document.All("ButtonHandler").Value = "Nothing Selected"   
			Case "Cancel"
						g_objIE.Quit
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings = 0
						'Call FocusToParentWindow(strPID)
						exit function
			Case "Exit_and_Reload_Settings"
						g_objIE.Quit
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings = 1000 + Int(MenuID)
						'Call FocusToParentWindow(strPID)
						exit function
			
			Case "Exit_After_Save"
						g_objIE.Quit
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings = 1
						'Call FocusToParentWindow(strPID)
						exit function
			Case "Exit_After_Save_3"
						g_objIE.Quit
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings = 3
						'Call FocusToParentWindow(strPID)
						exit function
			Case "Exit_And_Close_Wscript"
						g_objIE.Quit
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objIE = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings = -1
						exit function
			Case "Folder_5"
						strFolder = g_objIE.Document.All("Settings_Param_5").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_5").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Folder_6"
						strFolder = g_objIE.Document.All("Settings_Param_6").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_6").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Folder_7"
						strFolder = g_objIE.Document.All("Settings_Param_7").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_7").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Folder_8"
						strFolder = g_objIE.Document.All("Settings_Param_8").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_8").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Folder_9"
						strFolder = g_objIE.Document.All("Settings_Param_9").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_9").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Folder_25"
						strFolder = BrowseForFile()
						g_objIE.Document.All("Settings_Param_25").Value = strFolder													
						g_objIE.Document.All("ButtonHandler").Value = "None"		
		    Case "UPDATE_CATALOG"
						g_objIE.Document.All("ButtonHandler").Value = "None"		
						'---------------------------------------------------------
						' RUN TELNET SCRIPT
						'---------------------------------------------------------
						strCmd = strCRTexe &_ 
							" /ARG " & strDirectoryWork & "\config\class_catalog.dat" &_	
                            " /ARG " & strDirectoryWork & "\config\settings.dat" &_								
							" /ARG " & strDirectoryWork
						strCmd = strCmd & " /SCRIPT " & strDirectoryWork & "\" & VBScript_UPDATE_Catalog
						Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)														
						Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork & "\config\class_catalog.dat", "", objDebug, MAX_LEN, 1, 1)						
						Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork & "\config\settings.dat", "", objDebug, MAX_LEN, 1, 1)						
						Call TrDebug ("IE_PromptForInput: " & " /ARG " & strDirectoryWork, "", objDebug, MAX_LEN, 1, 1)
						Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & strDirectoryWork & "\" & VBScript_DNLD_Config, "",objDebug, MAX_LEN, 1, 1)						
						g_objShell.run strCmd, nWindowState, True
						'Reload Classes and objects properties
						'g_objIE.Document.All("ButtonHandler").Value = "Exit_and_Reload_Settings"
						Call GetMyClass(strDirectoryWork & "\config\class_catalog.dat", vObjIndex, nDebug)
						Call SetMyObject(objMain,"JunosSW",nDebug)
						Call SetMyObject(objMinor,"Release",nDebug)
						' Refresh forms
					    For nImage = 0 to UBound(objMain,1) - 1
							g_objIE.Document.All("Image_Param_7")(nImage).Value = objMain(nImage,pIndex(0,"Platform"))
							g_objIE.Document.All("Image_Param_9")(nImage).Value = objMain(nImage,pIndex(0,"Display_Name"))
							nOptions = g_objIE.document.getElementById("Main_Release" & nImage).SelectedIndex
							strMinorName = objMain(nImage,pIndex(0,"Name")) & "-" & g_objIE.document.getElementById("Main_Release" & nImage).Options(nOptions).Text
							For nMinor = 0 to UBound(objMinor,1) - 1
							   If strMinorName = objMinor(nMinor,pIndex(1,"Name")) Then Exit For
							Next
							vMinor = Split(objMinor(nMinor,pIndex(1,"Minor List")),",")
							nInd = 0
							For Each strMinor in vMinor
							   g_objIE.document.getElementById("Minor_Release" & nImage).Options(nInd).Text = strMinor
							   nInd = nInd + 1
							Next
						Next		
            Case "Clear_0", "Clear_1", "Clear_2", "Clear_3", "Clear_4", "Clear_5", "Clear_6", "Clear_7","Clear_8"  
			            nImage = Int(Split(szNothing,"_")(1))
					    g_objIE.document.getElementById("Minor_Release" & nImage).SelectedIndex = -1
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "SAVE" 
						g_objIE.Document.All("ButtonHandler").Value = "None"
						For i = 0 to Ubound(vSettings)
						    vOld_Settings(i) = vSettings(i)
						Next
                        Redim vOld_SessionCRT(UBound(vSessionCRT))						
						For i = 0 to Ubound(vSessionCRT) - 1
						    vOld_SessionCRT(i) = vSessionCRT(i)
						Next
						Do
					        '----------------------------------------------------
							'   DO CONSISTENCY AND FORMAT CHECK OF THE SETTINGS
							'----------------------------------------------------
						    Select Case MenuID
							    Case 7
								        If Not objFSO.FileExists(g_objIE.Document.All("Settings_Param_25").Value) Then 
											vvMsg(0,0) = "WRONG FILE NAME OR FILE DOESN'T EXISTS"	   : vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor2
											vvMsg(1,0) = "Try again "           		   : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1 											
											Call IE_MSG(vIE_Scale, "Error",vvMsg,2,"Null")
										    Exit Do										
										End If 
								Case 6  
								        If InStr(g_objIE.Document.All("Settings_Param_26").Value," ") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"/") > 0 Or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"\") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"""") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"'") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"#") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"$") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"*") > 0 or _
										InStr(g_objIE.Document.All("Settings_Param_26").Value,"?") > 0 Then 
											vvMsg(0,0) = "IVALID CONFIGURATION NAME!"	   : vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor2
											vvMsg(1,0) = "<space>, #,\,/,$,*,? "           : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1 
                                            vvMsg(2,0) = "symbols are not allowed "        : vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor1 											
											vvMsg(3,0) = "Try again "           		   : vvMsg(3,1) = "normal" : vvMsg(3,2) = HttpTextColor1 											
											Call IE_MSG(vIE_Scale, "Error",vvMsg,4,"Null")
										    Exit Do
										End If 			
								Case 3  ' CHEK FOLDERS
										'----------------------------------------------------
										'   SAVE NEW WorkFolder to Registry
										'----------------------------------------------------
										strFolder = g_objIE.Document.All("Settings_Param_6").Value
										If Right(strFolder,1) = "\" Then 
											strFolder = Left(strFolder,Len(strFolder)-1)
											g_objIE.Document.All("Settings_Param_6").Value = strFolder
										End If
										If strFolder <> Split(vOld_Settings(6),"=")(1) Then 
											If Not Continue("You are about to change work folder of the MEF CFG Loader Script!?", "Continue?") Then 
											   g_objIE.Document.All("Settings_Param_6").Value = Split(vOld_Settings(6),"=")(1)
											   Exit Do
											End If
											If Not objFSO.FolderExists(strFolder) Then 
												MsgBox "Error: Can't find Work Folder: " & chr(13) & strFolder 
												g_objIE.Document.All("Settings_Param_6").Value = Split(vOld_Settings(6),"=")(1)
												Exit Do
											End If	    
											Set objFolder = objFSO.GetFolder(strFolder)
											Set colFiles = objFolder.Files
											nResult = 0
											For Each objFile in colFiles
												strFile = objFile.Name
												If InStr(LCase(strFile),LCase(LDR_SCRIPT_NAME)) Then nResult = 1 End If 
											Next
											If nResult = 0 Then 
												MsgBox "Error: Folder doesn't contain a Cfg Loader script: " & chr(13) & strFolder & chr(13) & "Check work folder path"
												g_objIE.Document.All("Settings_Param_6").Value = Split(vOld_Settings(6),"=")(1)
												Exit Do
											End If
											On Error Resume Next
											Err.Clear
											g_objShell.RegWrite LAB_CFG_LDR_REG, strFolder, "REG_SZ"
											if Err.Number <> 0 Then 
												MsgBox "Error: Can't Right to Windows Registry" & chr(13) & Err.Description
												g_objIE.Document.All("Settings_Param_6").Value = Split(vOld_Settings(6),"=")(1)
											Else
												g_objIE.Document.All("ButtonHandler").Value = "Exit_And_Close_Wscript"   
											End If 
											On Error Goto 0
											Exit Do
										End If
										'-------------------------------------
										'   GET WORK FOLDER AND SETTINGS FILE
										'-------------------------------------
										strDirectoryWork = g_objIE.Document.All("Settings_Param_6").Value
										strFileSettings = strDirectoryWork & "\config\settings.dat"
										'-------------------------------------
										'   CHECK OTHER FOLDERS
										'-------------------------------------
										For nInd = 6 to 9
											strFolder = g_objIE.Document.All("Settings_Param_" & nInd).Value
											If Right(strFolder,1) = "\" Then 
												strFolder = Left(strFolder,Len(strFolder)-1)
												g_objIE.Document.All("Settings_Param_" & nInd).Value = strFolder
											End If
											If Not objFSO.FolderExists(g_objIE.Document.All("Settings_Param_" & nInd).Value) Then
												MsgBox "Folder doesn't exist: " & chr(13) &  g_objIE.Document.All("Settings_Param_" & nInd).Value
												Exit Do
											End If
										Next
									
                                Case 4  '   CHECK IP ADDRESS FORMAT
                                    For i = 0 to UBound(vSessionCRT) - 1								
										If Not CheckAddrFormat(g_objIE.Document.All("Session_Param_0")(i).Value,False) Then
											MsgBox "Wrong IP address format Session (" & i & "): " & g_objIE.Document.All("Session_Param_0")(i).Value &_
													chr(13) & "Use the following format for management IP address: "  &  "A.B.C.D"
											Exit Do
										End If						
									Next
							End Select ' End of consistency check
							'--------------------------------------------------------------------
							'  PREPARE UPDATED vSETTINGS, vSessionCRT, Objects AND WRITE SETTINGS TO FILE
							'--------------------------------------------------------------------
							g_objIE.Document.All("ButtonHandler").Value = "Exit_After_Save"
							Select Case MenuID
							    Case 7 ' Enter name of the file with list of configuration names to download
								        vSettings(25) = vParamNames(25) & Space(30 - Len(vParamNames(25))) & "= " & g_objIE.Document.All("Settings_Param_25").Value
							    Case 6
								       vSettings(26) = g_objIE.Document.All("Settings_Param_26").Value
									   g_objIE.Document.All("ButtonHandler").Value = "Exit_After_Save"
									   Exit Do
						        Case 0 
								Case 1
										nPlatform = g_objIE.document.getElementById("Platform_Name").selectedIndex
										vSettings(13) = vParamNames(13) & Space(30 - Len(vParamNames(13))) & "= " & g_objIE.document.getElementById("Platform_Name").Options(nPlatform).text
										vSettings(14) = vParamNames(14) & Space(30 - Len(vParamNames(14))) & "= " & g_objIE.Document.All("Platform_Index").Value								
								Case 2
								        For nInd = 1 to 4
										    Select Case nInd
												Case 1
												    nAdapter = g_objIE.document.getElementById("Adapter_Name").selectedIndex
													vSettings(nInd) = vParamNames(nInd) & Space(30 - Len(vParamNames(nInd))) & "= " & g_objIE.document.getElementById("Adapter_Name").Options(nAdapter).Text
												Case 2,3,4
													vSettings(nInd) = vParamNames(nInd) & Space(30 - Len(vParamNames(nInd))) & "= " & g_objIE.Document.All("Settings_Param_" & nInd).Value
												End Select
										Next
								Case 3
     							        For nInd = 5 to 9
										    vSettings(nInd) = vParamNames(nInd) & Space(30 - Len(vParamNames(nInd))) & "= " & g_objIE.Document.All("Settings_Param_" & nInd).Value
										Next
                                Case 4
										If g_objIE.Document.All("Hide_CRT").Checked Then 
											 vSettings(0) = vParamNames(0) & Space(30 - Len(vParamNames(0))) & "= 1"
										Else 
											 vSettings(0) = vParamNames(0) & Space(30 - Len(vParamNames(0))) & "= 0"										
										End If
										' Normalize folder session name across all CRT sessions
										Redim vSessionCRT_to_file(UBound(vSessionCRT))
										For nInd = 0 to UBound(vSessionCRT) - 1
										    'UPDATE SESSION ARREY WITH NEW DATA
										    vSessionCRT(nInd) = _
											g_objIE.Document.All("Session_Param_0")(nInd).Value & "," &_
											g_objIE.Document.All("Session_Param_1").Value       & "," &_
											g_objIE.Document.All("Session_Param_2")(nInd).Value & "," &_
											g_objIE.Document.All("Session_Param_3")(nInd).Value & "," &_
											g_objIE.Document.All("Session_Param_4")(nInd).Value
                                            'CREATE LIST OF NEW SESSION RECORDS TO BE PLACED INTO SESSIONGS FILE
										    vSessionCRT_to_file(nInd) = SECURECRT_SESSION & " " & nInd + 1 & Space(30 - Len(SECURECRT_SESSION & " " & nInd + 1)) & "= " &_
											g_objIE.Document.All("Session_Param_0")(nInd).Value & ", " &_
											g_objIE.Document.All("Session_Param_1").Value & ", " &_
											g_objIE.Document.All("Session_Param_2")(nInd).Value & ", " &_
											g_objIE.Document.All("Session_Param_3")(nInd).Value & ", " &_
											g_objIE.Document.All("Session_Param_4")(nInd).Value
										Next
                                Case 5
 								        vSettings(12) = vParamNames(12) & Space(30 - Len(vParamNames(12))) & "= " & g_objIE.Document.All("Settings_Param_12").Value
								Case 8
								        For nImage = 0 to UBound(objMain,1) - 1
											nOptions = g_objIE.document.getElementById("Main_Release" & nImage).SelectedIndex
											objMain(nImage,pIndex(0,"Main")) = g_objIE.document.getElementById("Main_Release" & nImage).Options(nOptions).Text
											nOptions = g_objIE.document.getElementById("Minor_Release" & nImage).SelectedIndex
											If nOptions => 0 Then 
												objMain(nImage,pIndex(0,"Minor")) = g_objIE.document.getElementById("Minor_Release" & nImage).Options(nOptions).Text
                                                objMain(nImage,pIndex(0,"Status")) = "Active"												
											Else 
											    objMain(nImage,pIndex(0,"Minor")) = "None"
												objMain(nImage,pIndex(0,"Status")) = "Inactive"
											End If
										Next
                         				If g_objIE.Document.All("DNLD_IMG_ONLY").Checked Then 
											g_objIE.Document.All("ButtonHandler").Value = "Exit_After_Save_3"
										End If										
                            End Select    
							'------------------------------------------------------------------
							'  WRITE NEW SETTINGS TO FILE
							'------------------------------------------------------------------
							Select Case MenuID
							        Case 0
									Case 1
										For nInd = 13 to 14
											Call FindAndReplaceStrInFile(strFileSettings, vParamNames(nInd), vSettings(nInd), 0)
											vSettings(nInd) = NormalizeStr(vSettings(nInd),vDelim)
										Next 									
									Case 2
										For nInd = 1 to 4
											Call FindAndReplaceStrInFile(strFileSettings, vParamNames(nInd), vSettings(nInd), 0)
											vSettings(nInd) = NormalizeStr(vSettings(nInd),vDelim)
										Next 									
									Case 3
										For nInd = 5 to 9
											Call FindAndReplaceStrInFile(strFileSettings, vParamNames(nInd), vSettings(nInd), 0)
											vSettings(nInd) = NormalizeStr(vSettings(nInd),vDelim)
										Next 									
									Case 4
											Call FindAndReplaceStrInFile(strFileSettings, vParamNames(0), vSettings(0), 0)
											vSettings(0) = NormalizeStr(vSettings(0),vDelim)
											For nInd = 0 to UBound(vSessionCRT) - 1
											    If vSessionCRT(nInd) <> vOld_SessionCRT(nInd) Then 
												    Call FindAndReplaceStrInFile(strFileSettings, SECURECRT_SESSION & " " & nInd + 1, vSessionCRT_to_File(nInd), 0)
												End If
											Next
									Case 5
											Call FindAndReplaceStrInFile(strFileSettings, vParamNames(12), vSettings(12), 0)
											vSettings(12) = NormalizeStr(vSettings(12),vDelim)
									Case 7 
											Call FindAndReplaceStrInFile(strFileSettings, vParamNames(25), vSettings(25), 0)
											vSettings(25) = NormalizeStr(vSettings(25),vDelim)	
                                    Case 8 
                        				For nImage = 0 to UBound(objMain,1) - 1
										    nCount = 0
											For Each objName in vObjIndex
											    If InStr(objName,"JunosSW") > 0 Then 
											       If nCount = nImage Then Exit For Else nCount = nCount + 1
												End If
											Next
											Call ReplaceFileLineInGroup(strDirectoryWork & "\config\class_catalog.dat", objName, "Main", "Main = " & objMain(nImage,pIndex(0,"Main")) ,nDebug)
											Call ReplaceFileLineInGroup(strDirectoryWork & "\config\class_catalog.dat", objName, "Minor", "Minor = " & objMain(nImage,pIndex(0,"Minor")),nDebug)
											Call ReplaceFileLineInGroup(strDirectoryWork & "\config\class_catalog.dat", objName, "Status", "Status = " & objMain(nImage,pIndex(0,"Status")),nDebug)
										Next
							End Select
							'-------------------------------------------------------
							'   SET SETTINGS VARIABLES CONSISTENT WITH DATA IN FILE
							'-------------------------------------------------------
							For nInd = 0 to 14
								Select Case Split(vSettings(nInd),"=")(0)
									Case HIDE_CRT
										Select Case Split(vSettings(nInd),"=")(1)
											Case "1"
												nWindowState = 2
											Case Else
												nWindowState = 1
										End Select									
									Case SECURECRT_FOLDER
										strCRT_InstallFolder = Split(vSettings(nInd),"=")(1)
									Case WORK_FOLDER
										strDirectoryWork = Split(vSettings(nInd),"=")(1)
									Case CONFIGS_FOLDER
										strDirectoryConfig =  Split(vSettings(nInd),"=")(1)
									Case CONFIGS_PARAM
										strFileParam = strDirectoryWork & "\config\" & Split(vSettings(nInd),"=")(1)
									Case FTP_IP
										strFTP_ip =  Split(vSettings(nInd),"=")(1)
									Case LAN_ADAPTER
										strEth =  Split(vSettings(nInd),"=")(1)
									Case FTP_User
										strFTP_name =  Split(vSettings(nInd),"=")(1)
									Case FTP_Password
										strFTP_pass =  Split(vSettings(nInd),"=")(1)
									Case Orig_Folder
										strTempOrigFolder = Split(vSettings(nInd),"=")(1)
									Case Dest_Folder
										strTempDestFolder = Split(vSettings(nInd),"=")(1)
									Case PLATFORM_NAME
										DUT_Platform = Split(vSettings(nInd),"=")(1)
									Case PLATFORM_INDEX
										Platform = Split(vSettings(nInd),"=")(1)									
								End Select
							Next
    						' g_objIE.Document.All("ButtonHandler").Value = "None"
							'------------------------------------------------------------------
							'  WRITE NEW SETTINGS TO FILE
							'------------------------------------------------------------------
							Select Case MenuID
							    Case 4 ' SAVE SECURE CRT SESSIONS SETTINGS
									If SecureCRT_Installed Then Call Create_CRT_Sessions(vSettings, vSessionCRT, vOld_SessionCRT)
								Case 2 ' RECREATE FTP USER AND HOME FOLDERS
									If FileZilla_Installed Then 
										If vSettings(3) <> vOld_Settings(3) or vSettings(4) <> vOld_Settings(4) or vSettings(7) <> vOld_Settings(7) Then 
											Set objShellApp = CreateObject("Shell.Application")
											objShellApp.ShellExecute "wscript", """" & VBScript_FTP_User & """" &_
																				" """ & strDirectoryConfig & """ " &_
																				strFTP_name & " " &_
																				Split(vOld_Settings(3),"=")(1) & " " &_
																				Split(vSettings(4),"=")(1), "", "runas", 1
											Set objShellApp = Nothing
										End If
									End If
							End Select
							Exit Do
					    Loop
			Case "EDIT_PARAM"
						strFileParam = strDirectoryWork & "\config\" & g_objIE.Document.All("Settings_Param_12").Value
						objEnvar.Run strEditor & " " & strFileParam
						' MsgBox "notepad.exe " & strFileParam
						vSettings(12) = CONFIGS_PARAM & "=" & g_objIE.Document.All("Settings_Param_12").Value
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "SET_NODE"
                        g_objIE.Document.All("ButtonHandler").Value = "None"
        End Select
        
		WScript.Sleep 300
    Loop
End Function

'##############################################################################
'      Function PROMPT FOR FOLDER NAME
'##############################################################################
 Function IE_DialogFolder (vIE_Scale, strTitle, strFolder, vLine, ByVal nLine, nDebug)
	Dim objForm, g_objIE, objShell
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
    Set g_objIE = Nothing
    Set objShell = Nothing
	Dim IE_Menu_Bar
	Dim  IE_Border
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_DialogFolder = -1	
	
	Call Set_IE_obj (g_objIE)
	Set objForm = CreateObject("Shell.Application")

	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy

'    With g_objIE.document.parentwindow.screen
'		intX = .availwidth
'        intY = .availheight
'    End With

	nRatioX = intX/1920
    nRatioY = intY/1080
	LineH   = Round (12 * nRatioY)	
	nHeader = Round(10 * nRatioY,0)
	nTab = Round(20 * nRatioX,0)

	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	WindowW = IE_Border + Round(550 * nRatioX,0)
	WindowH = IE_Menu_Bar + 2 * (5 + nLine) * LineH + nButtonY
	
'   If nDebug = 1 Then MsgBox "intX=" & intX & "   intY=" & intY & "   RatioX=" & nRatioX & "  RatioY=" & nRatioY & "   Cell Width=" & cellW & "  Cell Hight=" & cellH End If

	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left =(intX - WindowW)/2
	strHTMLBody = "<br>"
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
			"<b><p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p></b>" 

	Next
	'---------------------------------------------------
	' SET INPUT FILED AND BUTTUN FOR LCL DIRECTORY
	'---------------------------------------------------
	strHTMLBody = strHTMLBody &_	
				"<input name='UserInput' size='80' maxlength='128' style=""position: absolute; Left: " & nTab & "px; top: " &_
				2 * ( nLine + 2) * LineH & "px; font-size: " & nFontSize_12 &_
				".0pt; border-style: None; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
				"; background-color: " & HttpBgColor2 & "; font-weight: bold;"">"				


	'---------------------------------------------------
	' SET OK and CANCEL
	'---------------------------------------------------
    strHTMLBody = strHTMLBody &_
			    "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nTab & "px; bottom: 4px' name='OK' AccessKey='O' onclick=document.all('ButtonHandler').value='OK';><u>O</u>K</button>" & _
                "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; right:" & nTab & "px; bottom: 4px' name='Cancel' AccessKey='C' onclick=document.all('ButtonHandler').value='Cancel';><u>C</u>ancel</button>" & _
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
	strHTMLBody = strHTMLBody &_
                "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
				";width:" & 2 * nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & Int(WindowW/2) - nButtonX & "px; bottom: 4px'" &_ 
				"name='Local' AccessKey='L' onclick=document.all('ButtonHandler').value='Local';><u>S</u>elect Folder</button>"


			
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
		WScript.Sleep 100
	Loop While g_objIE.Busy
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.AppActivate g_objIE.document.Title
	g_objIE.Document.All("UserInput").Focus
	g_objIE.Document.All("UserInput").Value = strFolder

	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case strNothing
			Case "Cancel"
				' The user clicked Cancel. Exit the loop
				g_objIE.quit
				Set g_objIE = Nothing
				IE_DialogFolder = False			
				Exit Do
			Case "OK"
				IE_DialogFolder = True
				strFolder = g_objIE.Document.All("UserInput").Value
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
			Case "Local"
				If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  "GetPromptFolder: Local Button Pushed" End If  
				strFolder = g_objIE.Document.All("UserInput").Value
				Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
				If Not objFolder Is Nothing Then
					strFolder = objFolder.self.path
					g_objIE.Document.All("UserInput").Value = strFolder
				End If
				g_objIE.Document.All("ButtonHandler").Value = "None"
		End Select
	Wscript.Sleep 200
Loop
    Set g_objIE = Nothing
    Set objShell = Nothing
	Set objFolder = Nothing
End Function
'---------------------------------------------------------------------------
'   Function Space_html(n)
'---------------------------------------------------------------------------
Function Space_html(n)
Dim Str, i
Str = ""
for i = 1 to n
  Str = Str & "&nbsp;" 
next
Space_html = Str
End Function
'---------------------------------------------------------------------------
'   Function Create_Zero_Configs(vSettings)
'---------------------------------------------------------------------------
Function Create_Zero_Configs(ByRef vSettings, Old_Platform_Index)
Dim objFSO, strCfgGlobal, strCfgRE0, Platform,strDirectoryConfig
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCfgGlobal =  Split(vSettings(27),"=")(1)
	strCfgRE0 = Split(vSettings(28),"=")(1)
	Platform = Split(vSettings(14),"=")(1)
	strDirectoryConfig = Split(vSettings(7),"=")(1)
	strDirectoryWork = Split(vSettings(6),"=")(1)
	strHostNameL = 	Split(Split(vSettings(10),"=")(1),",")(2)
	strHostNameR = 	Split(Split(vSettings(11),"=")(1),",")(2)
    strLeft_ip	= Split(vSettings(0),"=")(1)
	strRight_ip = Split(vSettings(1),"=")(1)
	DUT_Platform = Split(vSettings(13),"=")(1)
	strGlobalFileL = strCfgGlobal & "-" & Platform & "-l.conf"
	strGlobalFileR = strCfgGlobal & "-" & Platform & "-r.conf"
	strRe0FileL = strCfgRE0 & "-" & Platform & "-l.conf"
	strRe0FileR = strCfgRE0 & "-" & Platform & "-r.conf"
	On Error Resume Next
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgGlobal & "-" & Old_Platform_Index & "-l.conf", True
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgGlobal & "-" & Old_Platform_Index & "-r.conf", True
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgRE0 & "-" & Old_Platform_Index & "-l.conf", True
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgRE0 & "-" & Old_Platform_Index & "-r.conf", True
	On Error Goto 0
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\" & DUT_Platform & "\" &  strCfgGlobal & ".conf" , _
	                strDirectoryConfig & "\" & strCfgGlobal & "-" & Platform & "-l.conf", True
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\"  & DUT_Platform & "\" &  strCfgGlobal & ".conf" , _
	                strDirectoryConfig & "\" & strCfgGlobal & "-" & Platform & "-r.conf", True
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\"  & DUT_Platform & "\" &  strCfgRE0 & ".conf" , _
	                strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-l.conf", True
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\"  & DUT_Platform & "\" &  strCfgRE0 & ".conf" , _
	                strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-r.conf", True
	Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-l.conf", "host-name","     host-name " & strHostNameL & ";", 0)
    Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-r.conf", "host-name","     host-name " & strHostNameR & ";", 0)
	Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-l.conf", "address","                address " & strLeft_ip & ";", 0)
	Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-r.conf", "address","                address " & strRight_ip & ";", 0)
End Function
'---------------------------------------------------------------------------
'   Function CheckAddrFormat(strAddr,Preffix required[True/False])
'---------------------------------------------------------------------------
Function CheckAddrFormat(strAddr, bPreffix)
	Do
		nCount = UBound(Split(strAddr,"."))
		If nCount <> 3 Then CheckAddrFormat = False : Exit Do : End If
		If bPreffix Then 
			nCount = UBound(Split(strAddr,"/")) 
			If nCount <> 1 Then CheckAddrFormat = False : Exit Do : End If
		End If
		For i = 0 to 3
		    nOctet = Split(Split(strAddr,"/")(0),".")(i)
		    if Not IsNumeric(nOctet) Then CheckAddrFormat = False : Exit Do : End If
			if i = 0 and (Int(nOctet) < 1 or Int(nOctet) > 255) then CheckAddrFormat = False : Exit Do : End If
			if i > 0 and (Int(nOctet) < 0 or Int(nOctet) > 255) then CheckAddrFormat = False : Exit Do : End If
		Next
        If bPreffix Then 
		    nPrefix = Split(strAddr,"/")(1)
			if Int(nPrefix) < 8 or Int(nPrefix) > 30 then CheckAddrFormat = False : Exit Do : End If
		End If
		CheckAddrFormat = True
		Exit Do
	Loop
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUserSYS
'----------------------------------------------------------------------------------
Function GetScreenUserSYS()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserSYS = strScreenUser
End Function
'----------------------------------------------------------------------------------
'    Function Create_CRT_Sessions(ByRef vSettings,ByRef vOld_Settings )
'----------------------------------------------------------------------------------
Function Create_CRT_Sessions(ByRef vSettings,ByRef vSessions, ByRef vOld_Sessions )
Dim strSessionFolder, strSessionNameL, strSessionNameR, strOldSessionFolder, strOldSessionNameL, strOldSessionNameR
Dim  StrWinUser, strLeft_ip, strRight_ip, strCRT_SessionFolder,strDirectoryWork, nIndex
Dim objFSO, nInd
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strDirectoryWork = Split(vSettings(6),"=")(1)
	strCRT_SessionFolder = Split(vSettings(15),"=")(1)
	StrWinUser = GetScreenUserSYS()
	For nInd = 0 to UBound(vSessions) - 1
		Do 
		   ' If vSessions(nInd) = vOld_Sessions(nInd) Then Exit Do 
			strSessionFolder = Split(vSessions(nInd),",")(1)
			strOldSessionFolder =  Split(vOld_Sessions(nInd),",")(1)
			strSessionNameL = Split(vSessions(nInd),",")(2)
			strOldSessionNameL = Split(vOld_Sessions(nInd),",")(2)
			strLeft_ip = Split(vSessions(nInd),",")(0)
			' Check if folder for your session exists
			If strOldSessionFolder <> strSessionFolder and objFSO.FolderExists(strCRT_SessionFolder & "\" & strOldSessionFolder) Then 
				objFSO.DeleteFolder strCRT_SessionFolder & "\" & strOldSessionFolder, True
			End If
            ' - Create New Session Folder
			If Not objFSO.FolderExists(strCRT_SessionFolder & "\" & strSessionFolder) Then 
				objFSO.CreateFolder strCRT_SessionFolder & "\" & strSessionFolder
			End If
			' - Create Root Folder
			If Not objFSO.FolderExists(strCRT_SessionFolder & "\" & strSessionFolder & "\root") then objFSO.CreateFolder strCRT_SessionFolder & "\" & strSessionFolder & "\root"
			' Check if __DataFolder__.ini file exists in ..\sessions
			If Not objFSO.FileExists(strCRT_SessionFolder & "\__FolderData__.ini") Then 
				objFSO.CopyFile strDirectoryWork & "\config\secureCRT\__FolderData__.ini", strCRT_SessionFolder & "\__FolderData__.ini", True
			End If
			' Check if __DataFolder__.ini has information about your session folder
			Call  GetFileLineCountSelect(strCRT_SessionFolder & "\__FolderData__.ini", vFolderData,"", "", "", 0)
			nIndex = GetObjectLineNumber(vFolderData, UBound(vFolderData),"Folder List") - 1
			If strOldSessionFolder <> strSessionFolder and InStr(vFolderData(nIndex),strOldSessionFolder & ":") then 
				vFolderData(nIndex) = Replace(vFolderData(nIndex), strOldSessionFolder & ":","")
			End If
			If Not Instr(vFolderData(nIndex),strSessionFolder  & ":") then 
				vFolderData(nIndex) = vFolderData(nIndex) & strSessionFolder & ":"
				Call WriteArrayToFile(strCRT_SessionFolder & "\__FolderData__.ini",vFolderData, UBound(vFolderData),1,0)
			End If
			' - Delete old session files
			If strOldSessionFolder = strSessionFolder and strOldSessionNameL <> stressionNameL and objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strOldSessionNameL & ".ini") Then
			   objFSO.DeleteFile strCRT_SessionFolder & "\" & strSessionFolder & "\" & strOldSessionNameL & ".ini", True
			End If 
			If strOldSessionFolder = strSessionFolder and strOldSessionNameL <> stressionNameL and objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\root\" & strOldSessionNameL & ".ini") Then
			   objFSO.DeleteFile strCRT_SessionFolder & "\" & strSessionFolder & "\root\" & strOldSessionNameL & ".ini", True
			End If 
			' - Create New Session File for the Node Session
			objFSO.CopyFile strDirectoryWork & "\config\secureCRT\node.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameL & ".ini", True
			objFSO.CopyFile strDirectoryWork & "\config\secureCRT\node_root.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\root\" & strSessionNameL & ".ini", True
			Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameL & ".ini", vSessionFile,"", "", "", 0)
			Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\root\" & strSessionNameL & ".ini", vRootFile,"", "", "", 0)
			nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{hostname}}") - 1
			vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{hostname}}",strLeft_ip)
			vRootFile(nIndex)    = Replace(vRootFile(nIndex),"{{hostname}}",strLeft_ip)
			nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{winusername}}") - 1
			vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{winusername}}",StrWinUser)
			vRootFile(nIndex)    = Replace(vRootFile(nIndex),"{{winusername}}",StrWinUser)
			nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{Workdirectory}}") - 1
			vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{Workdirectory}}",strDirectoryWork)
            vRootFile(nIndex)    = Replace(vRootFile(nIndex),"{{Workdirectory}}",strDirectoryWork)			
			Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameL & ".ini",vSessionFile, UBound(vSessionFile),1,0)
			Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\root\" & strSessionNameL & ".ini",vRootFile, UBound(vSessionFile),1,0)
			' Check if __DataFolder__.ini file exists in ..\sessions\SessionName
			If Not objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini") Then 
				objFSO.CopyFile strDirectoryWork & "\config\secureCRT\__FolderData__.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini", True
			End If
			' Check if __DataFolder__.ini has information about your new sessions
			Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini", vFolderData,"", "", "", 0)
			nIndex = GetObjectLineNumber(vFolderData, UBound(vFolderData),"Folder List") - 1
			If Not Instr(vFolderData(nIndex),"root:") then 
				vFolderData(nIndex) = vFolderData(nIndex) & "root:"
			End If			
			nIndex = GetObjectLineNumber(vFolderData, UBound(vFolderData),"Session List") - 1
			If Not Instr(vFolderData(nIndex),strSessionNameL & ":") then 
				vFolderData(nIndex) = vFolderData(nIndex) & strSessionNameL & ":"
			End If
			Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini",vFolderData, UBound(vFolderData),1,0)	
			' Check if __DataFolder__.ini file exists in ..\sessions\SessionName\root
			If Not objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\root\__FolderData__.ini") Then 
				objFSO.CopyFile strDirectoryWork & "\config\secureCRT\__FolderData__.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\root\__FolderData__.ini", True
			End If
			' Check if __DataFolder__.ini has information about your new sessions			
			Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\root\__FolderData__.ini", vFolderData,"", "", "", 0)
			nIndex = GetObjectLineNumber(vFolderData, UBound(vFolderData),"Session List") - 1
			If Not Instr(vFolderData(nIndex),strSessionNameL & ":") then 
				vFolderData(nIndex) = vFolderData(nIndex) & strSessionNameL & ":"
			End If
			Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\root\__FolderData__.ini",vFolderData, UBound(vFolderData),1,0)	
	        Exit Do
		Loop
	Next
    Set objFSO = Nothing
End Function
'###################################################################################
' Displays a Message Box with Cancel / Continue buttons                 
'###################################################################################
Function Continue(strMsg, strTitle)
    ' Set the buttons as Yes and No, with the default button
    ' to the second button ("No", in this example)
    nButtons = vbYesNo + vbDefaultButton2
    
    ' Set the icon of the dialog to be a question mark
    nIcon = vbQuestion
    
    ' Display the dialog and set the return value of our
    ' function accordingly
    If MsgBox(strMsg, nButtons + nIcon, strTitle) <> vbYes Then
        Continue = False
    Else
        Continue = True
    End If
End Function 
'--------------------------------------------------------------------
' Function Runs MS CMD Command on local or remote PC
'--------------------------------------------------------------------
Function RunCmd(strHost, strPsExeFolder, ByRef vCmdOut, strCMD, nDebug)	
	Dim nResult, f_objFSO, objShell
	Dim nCmd, stdOutFile, objCmdFile, cmdFile, strRnd,strWork,strPsExec
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	strRnd = My_Random(1,999999)
	stdOutFile = "svc-" & strRnd & ".dat"
	cmdFile = "run-" & strRnd & ".bat"
    strWork = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	If strHost = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") or strHost = "127.0.0.1" Then 
		strPsExec = ""
	Else 
		strPsExec = strPsExeFolder & "\psexec \\" & strHost & " -s "
	End If
	'-------------------------------------------------------------------
	'       CREATE A NEW TERMINAL SESSION IF REQUIRED
	'-------------------------------------------------------------------
	Set objCmdFile = objFSO.OpenTextFile(strWork & "\" & cmdFile,ForWriting,True)
	Call TrDebug ("COMMAND: ", strPsExec & strCMD & " >" & strWork & "\" & stdOutFile, objDebug, MAX_WIDTH, 1, nDebug)
	objCmdFile.WriteLine strPsExec & strCMD & " >" & strWork & "\" & stdOutFile
	objCmdFile.WriteLine "Exit"
	objCmdFile.close
	objShell.run strWork & "\" & cmdFile,0,True
	Call TrDebug ("BATCH FILE EXECUTED: ", strWork & "\" & cmdFile, objDebug, MAX_WIDTH, 1, nDebug)
	wscript.sleep 100
	'-----------------------------------------
	' READ OUTPUT FILE AND DELETE WHEN DONE
	'-----------------------------------------
	RunCmd = GetFileLineCountSelect(strWork & "\" & stdOutFile, vCmdOut,"NULL","NULL","NULL",nDebug)
	If f_objFSO.FileExists(strWork & "\" & stdOutFile) Then
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & stdOutFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",stdOutFile, objDebug, MAX_WIDTH, 1, 1)
			On Error goto 0
		End If	
	End If
	If f_objFSO.FileExists(strWork & "\" & cmdFile) Then 
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & cmdFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",cmdFile, objDebug, MAX_WIDTH, 1, 1)
			On Error goto 0
		End If		
	End If
	Set f_objFSO = Nothing
	Set objShell = Nothing
	If RunCmd = 0 Then 
		Call TrDebug ("RunCmd: " & strCMD & " ERROR: ", "CAN'T WRITE TO OUTPUT FILE OR EMPTY FILE" , objDebug, MAX_WIDTH, 1, 1)
		Exit Function 
	End If
End Function
'--------------------------------------------------------------
' Function returns a random intiger between min and max
'--------------------------------------------------------------
Function My_Random(min, max)
	Randomize
	My_Random = (Int((max-min+1)*Rnd+min))
End Function
'----------------------------------------------------------------
'   Function MinimizeParentWindow()
'----------------------------------------------------------------
Function MinimizeParentWindow()
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
    wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "n"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function MinimizeParentWindow()
'----------------------------------------------------------------
Function RestoreParentWindow()
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
    wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function FocusToParentWindow(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function FocusToParentWindow(strPID)
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "%"
	wscript.sleep IE_PAUSE
	objShell.AppActivate strPID			
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function GetAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetAppPID(ByRef strPID, strAppName)
Dim objWMI, colItems
Const IE_PAUSE = 70
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Exit Do
		End If 
'		wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'Launcher Ver.'"  WHERE Name = 'iexplore.exe' OR Name = 'wscript.exe'
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("GetMyPID: RESTORE IE WINDOW:", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
			If pUser = strUser then 
				strPID = process.ProcessId
				Call TrDebug ("GetMyPID: ", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
				Call TrDebug ("GetMyPID: ", "Caption: " & process.Caption & ", CSName " & process.CSName & ", Description: " & process.Description & ", Handle: " &  Process.Handle,objDebug, MAX_LEN, 1, 1) 
			GetMyPID = True
				Exit For
			End If
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'-----------------------------------------------------------------------
'   Function GetAppPID(strPID) Returns focus to the parent Window/Form
'-----------------------------------------------------------------------
Function GetFilterList(vConfigFileLeft, vFilterList, vPolicerList, vCIR, vCBS, nDebug)
Dim nFilter, nPolicer,FW_Start, FoundFilter, FoundPolicer, strLine
Redim vFilterList(0)
Redim vPolicerList(0)
Redim vCIR(0)
Redim vCBS(0)
    GetFilterList = False
	nFilter = 0
	nBraces = 10000
	For Each strLine in vConfigFileLeft
	    If InStr(strLine,"firewall ")<>0 Then FW_Start = True : nBraces = 0
		If InStr(strLine," {")<>0 Then nBraces = nBraces + 1
		If InStr(strLine," }")<>0 Then nBraces = nBraces - 1
		If nBraces = 2 Then FoundFilter = False
		If nBraces = 1 Then FoundPolicer = False
		if nBraces = 0 Then Exit For
		If Instr(strLine,"filter ")<>0 and InStr(strLine, " {")<>0 and Instr(strLine,"filter {")=0 Then 
		    Redim Preserve vFilterList(nFilter + 1)
			Redim Preserve vPolicerList(nFilter + 1)
		    Redim Preserve vCIR(nFilter + 1)
			Redim Preserve vCBS(nFilter + 1)
			If nFilter > 0 Then 
			    If vPolicerList(nFilter-1) = "" Then 
				    vPolicerList(nFilter-1) = "N/A"
				    Call TrDebug ("GetFilterList: NO POLICER FOUND: ", vPolicerList(nFilter-1),objDebug, MAX_LEN, 1, nDebug) 
                End If				
            End If
		    vFilterList(nFilter) = Split(Split(strLine,"filter ")(1)," {")(0)
			Call TrDebug ("GetFilterList: FOUND FILTER:", vFilterList(nFilter),objDebug, MAX_LEN, 1, nDebug)
			GetFilterList = True
			nFilter = nFilter + 1
			FoundFilter = True
		End If
		If FoundFilter Then 
			If Instr(strLine,"-rate")<>0 Then 
				vPolicerList(nFilter-1) = Split(Split(strLine,"-rate ")(1),";")(0)
				Call TrDebug ("GetFilterList: FOUND POLICER (key ""-rate""):", vPolicerList(nFilter-1),objDebug, MAX_LEN, 1, nDebug) 						
			End If
			If Instr(strLine,"policer")<>0 and InStr(strLine,"-policer")=0 and InStr(strLine," {")=0 Then 
				vPolicerList(nFilter-1) = Split(Split(strLine,"policer ")(1),";")(0)
				Call TrDebug ("GetFilterList: FOUND POLICER (key ""-policer""):", vPolicerList(nFilter-1),objDebug, MAX_LEN, 1, nDebug)
			End If
		End If
		If Not FoundFilter Then' - Parser is out of Firewall filter hierarchy
			If InStr(strLine,"policer ")<>0 and InStr(strLine," {")<>0 Then 
			    strFilterCount = ""
				FoundPolicer = False
				For nPolicer=0 to UBound(vPolicerList)-1
				   If InStr(strLine,vPolicerList(nPolicer) & " {")<>0 Then 
				        FoundPolicer = True 
						Call TrDebug ("GetFilterList: FOUND SETTINGS FOR POLICER:", vPolicerList(nPolicer),objDebug, MAX_LEN, 1, nDebug)
						If strFilterCount = "" Then strFilterCount = CStr(nPolicer) Else   strFilterCount = strFilterCount & "," & CStr(nPolicer)
					End If
				Next
			End If
			If FoundPolicer Then 
			    vPolicer = Split(strFilterCount,",")
				For each nLine in vPolicer
				    nPolicer = CInt(nLine)
					If InStr(strLine,"committed-information-rate")<>0 Then vCIR(nPolicer) = Split(Split(strLine,"committed-information-rate ")(1),";")(0)
					If InStr(strLine,"bandwidth-limit")<>0 Then vCIR(nPolicer) = Split(Split(strLine,"bandwidth-limit ")(1),";")(0)
					If InStr(strLine,"peak-information-rate")<>0 Then vCIR(nPolicer) = vCIR(nPolicer) & " / " & Split(Split(strLine,"peak-information-rate ")(1),";")(0)
					If InStr(strLine,"committed-burst-size")<>0 Then vCBS(nPolicer) = Split(Split(strLine,"committed-burst-size ")(1),";")(0)				
					If InStr(strLine,"burst-size-limit")<>0 Then vCBS(nPolicer) = Split(Split(strLine,"burst-size-limit ")(1),";")(0)
					If InStr(strLine,"peak-burst-size")<>0 Then vCBS(nPolicer) = vCBS(nPolicer) & " / " & Split(Split(strLine,"peak-burst-size ")(1),";")(0)		
				Next
			End If
        End If			
	Next
End Function
'-----------------------------------------------------------------------------------------
'      Function Displays a Message with Continue and No Button. Returns True if Continue
'-----------------------------------------------------------------------------------------
 Function IE_CONT (vIE_Scale, strTitle, vLine, ByVal nLine, objIEParent, nDebug)
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim g_objIE, objShell
    Set g_objIE = Nothing
    Set objShell = Nothing
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_CONT = False
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call IE_Hide(objIEParent)
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nRatioX = intX/1920
    nRatioY = intY/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	nTab = Round(20 * nRatioX,0)
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
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
			
	Next		
	
    strHTMLBody = strHTMLBody &_
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; left: " & nTab & "px; bottom: " & BottomH & "px' name='Continue' AccessKey='Y' onclick=document.all('ButtonHandler').value='YES';><u>Y</u>ES</button>" & _
								"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; right: " & nTab & "px; bottom: " & BottomH & "px' name='NO' AccessKey='N' onclick=document.all('ButtonHandler').value='NO';><u>N</u>O</button>" & _
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
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strMyPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then strMyPID = Split(strLine,""",""")(1)
	     ' Call TrDebug("READ TASK PID:" , strLine, objDebug, MAX_LEN, 1, 1)
    Next
    If strMyPID = "" Then Call GetAppPID(strMyPID, "iexplore.exe")
	objShell.AppActivate strMyPID										
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "NO"
				IE_CONT = False
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
			Case "YES"
				IE_CONT = True
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
		Call IE_UnHide(objIEParent)
End Function
'----------------------------------------------------------------------------------------
'   Function IE_Hide(objIE) changes the visibility of the Window referenced by the objIE
'----------------------------------------------------------------------------------------
Function IE_Hide(byRef objIE)
   if objIE = "Null" then exit function
   If objIE.Visible then 
        objIE.Visible = False
    End If 
End Function
'----------------------------------------------------------------------------------------
'   Function IE_UnHide(objIE) changes the visibility of the Window referenced by the objIE
'----------------------------------------------------------------------------------------
Function IE_Unhide(byRef objIE)
   if objIE = "Null" then exit function
   If Not objIE.Visible then 
        objIE.Visible = True
    End If 
End Function
'----------------------------------------------------------------------------------------
'   Function UpdateCfgList(g_objIE, nCfg, strYear, strTag, ByRef vCfgList, htmlCfgSelect)
'----------------------------------------------------------------------------------------
Function UpdateCfgList(byRef g_objIE, nCfg, strYear, strTag, ByRef vCfgList, htmlCfgSelect)
    Dim nOptions_New, nOptions, nInd
	nOptions_New = 0
	g_objIE.document.getElementById("cfg_name").Options(0).Text = "N/A"
	g_objIE.document.getElementById("cfg_name").Options(0).Value = 0				   	
	For nInd = 1 to UBound(vCfgList,1)
		g_objIE.document.getElementById(htmlCfgSelect).Options(nInd).Text = Space(128)
		g_objIE.document.getElementById(htmlCfgSelect).Options(nInd).Value = "N/A"					   
	Next
	nOptions = 0
	For nInd = 0 to UBound(vCfgList,1) - 1
		Do 
			If strYear <> "All" Then 
				If InStr(vCfgList(nInd,1),strYear) = 0 Then 
					Exit Do
				End If
			End If
			If strTag <> "*" Then 
				If InStr(vCfgList(nInd,0),LCase(strTag)) = 0 and InStr(vCfgList(nInd,0),strTag) = 0 Then 
					Exit Do 
				End If
			End If			   
			g_objIE.document.getElementById(htmlCfgSelect).Options(nOptions).Text = vCfgList(nInd,0)
			g_objIE.document.getElementById(htmlCfgSelect).Options(nOptions).Value = nInd
'			MsgBox "nCfg = " & nCfg & chr(13) & "nInd = " & nInd
			If nInd = CInt(nCfg) Then 
				nOptions_New = nOptions
			End If
			nOptions = nOptions + 1
			Exit Do
		Loop 
	Next
'	g_objIE.document.getElementById(htmlCfgSelect).Options(nOptions).Text = SAVE_AS & Space(100)
'	g_objIE.document.getElementById(htmlCfgSelect).Options(nOptions).Value = nInd + 1	
	UpdateCfgList = nOptions_New
End Function
'----------------------------------------------------------------------------------------
'   Function UpdateCfgVer(g_objIE, nCfg, strYear, strTag, ByRef vCfgList, htmlCfgSelect)
'----------------------------------------------------------------------------------------
Function UpdateCfgVer(ByRef g_objIE, ByRef nCfg, ByRef vCfgList, htmlSelect)
	Dim nInd, strLine
	UpdateCfgVer = 0
	If Int(nCfg) >= Ubound(vCfgList,1) Then Exit Function
	strLine = Split(vCfgList(nCfg,1),"=")(1)
	Const MAX_PARAM = 10
	For nInd = 0 to MAX_PARAM
		If nInd < UBound(Split(strLine,",")) Then 
			g_objIE.document.getElementById(htmlSelect).Options(nInd).text = Split(strLine,",")(nInd)
			g_objIE.document.getElementById(htmlSelect).Options(nInd).Value = nInd
		End If
		If nInd = UBound(Split(strLine,",")) Then 
			g_objIE.document.getElementById(htmlSelect).Options(nInd).text = Split(strLine,",")(nInd)
			g_objIE.document.getElementById(htmlSelect).Options(nInd).Value = nInd
			g_objIE.document.getElementById(htmlSelect).selectedIndex = nInd			
		End If
		If nInd > UBound(Split(strLine,",")) Then 
			g_objIE.document.getElementById(htmlSelect).Options(nInd).text = Space(18)
			g_objIE.document.getElementById(htmlSelect).Options(nInd).Value = 0
		End If
	Next
	UpdateCfgVer = g_objIE.document.getElementById(htmlSelect).selectedIndex
End Function 
'----------------------------------------------------------------------------------------
'   Function UpdateSessionStatus(ByRef g_objIE, nCfg, strCfg, ByRef vCfgList,ByRef vSessionCRT, ByRef vSessionEnable)
'----------------------------------------------------------------------------------------
Function UpdateSessionStatus(ByRef g_objIE, nCfg, strCfg, ByRef vCfgList,ByRef vSessionCRT, ByRef vSessionEnable)
	Dim nInd, i, strBoxName
    If InStr(strCfg,SAVE_AS) > 0 Then 
		For nInd = 0 to UBound(vSessionCRT) - 1
		    strBoxName = Split(vSessionCRT(nInd),",")(2)
			If InStr(vSessionEnable(nInd),"Enabled") > 0 Then 
			   'g_objIE.Document.All(strBoxName).Select
			   g_objIE.Document.All(strBoxName).Checked = true
			   'g_objIE.Document.All(strBoxName).Click
			Else 
			    g_objIE.Document.All(strBoxName).Checked = false
			End If 		       
		Next 
    Else 
		For nInd = 0 to UBound(vSessionCRT) - 1
		    strBoxName = Split(vSessionCRT(nInd),",")(2)
			For i = 1 to UBound(vCfgList,2)
			    If i = UBound(vCfgList,2) Then 
'				   g_objIE.Document.All(strBoxName).Select
				   g_objIE.Document.All(strBoxName).Checked = false
                   vSessionEnable(nInd) = "Status " & nInd + 1 & "=Disabled"
				   Exit For
				End If 
			    If vCfgList(nCfg,i) = "" Then 
'				   g_objIE.Document.All(strBoxName).Select
				   g_objIE.Document.All(strBoxName).Checked = false
                   vSessionEnable(nInd) = "Status " & nInd + 1 & "=Disabled"				   
				   Exit For
				End If
				If InStr(vCfgList(nCfg,i),strBoxName) > 0 Then 
'				   g_objIE.Document.All(strBoxName).Select
				   g_objIE.Document.All(strBoxName).Checked = True
'				   g_objIE.Document.All(strBoxName).Click
                   vSessionEnable(nInd) = "Status " & nInd + 1 & "=Enabled"				   
				   Exit For
				End If 		       
			Next
		Next 	    
    End If
End Function
'----------------------------------------------------------------------------
'    Function CreateNewCfg(ByRef strCfg, nCfg, ByRef strVersion, strDirectoryConfig, ByRef vCfgInventory, ByRef vCfgList, nDebug)
'----------------------------------------------------------------------------
Function CreateNewCfg(ByRef strCfg, ByRef nCfg, ByRef strVersion, strDirectoryConfig, ByRef vCfgInventory, ByRef vCfgList, nDebug)
    Dim objFSO, vFileLine, nLine, strLine, nVersion, vLine(1), vCfgTempList
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	If nCfg = 0 or GetExactObjectLineNumber(vCfgInventory, UBound(vCfgInventory),strCfg) = 0 Then 
	    '----------------------------------------------------
		'   CREATE NEW CFG RECORD AND VERSION
		'----------------------------------------------------
	    nCfg = UBound(vCfgInventory)
		' write new cfg name to CfgList file to the END of the list
		Call WriteStrToFile(strDirectoryConfig & "\CfgList.txt", strCfg, vCfgInventory(nCfg - 1), 3, 0)
		' write new cfg name to CfgInventory Array
		Redim Preserve vCfgInventory(nCfg + 1)
		vCfgInventory(nCfg) = strCfg
		' write new cfg name to CfgList Array
		nDim1 = UBound(vCfgList,1)
		nDim2 = UBound(vCfgList,2)
		Redim vCfgTempList(nDim1, nDim2)
		For i = 0 to nDim1 - 1
		    For n = 0 to nDim2 - 1
			   vCfgTempList(i,n) = vCfgList(i,n)
		    Next
		Next
		Redim vCfgList(nDim1 + 1, nDim2)
		For i = 0 to nDim1 - 1
		    For n = 0 to nDim2 - 1
			   vCfgList(i,n) = vCfgTempList(i,n)
		    Next
		Next
		vCfgList(nCfg,0) = strCfg
		' Create new Version Number
		strVersion = Year(Date) & "-" & Month(Date()) & "-" & Day(Date()) & "-" & "v.01"
		Redim vFileLine(3)
		vFileLine(0) = " "
		vFileLine(1) = "[" & strCfg & "]"
		vFileLine(2) = "Version = " & strVersion
		Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vFileLine, UBound(vFileLine),2,0)
	Else 
	    '------------------------------------------------------
		'   CREATE NEW VERSION FOR EXISTED CFG
		'------------------------------------------------------
	    Call GetFileLineCountByGroup(strDirectoryConfig & "\CfgList.txt", vFileLine,strCfg,"","",0)
		nLine = GetObjectLineNumber(vFileLine, UBound(vFileLine),"Version")
		If UBound(Split(vFileLine(nLine - 1),"v.")) > 0 Then 
		    nVersion = CInt(Split(vFileLine(nLine - 1),"v.")(UBound(Split(vFileLine(nLine - 1),"v.")))) + 1
			If nVersion > 9 Then 
			    strVersion = Year(Date) & "-" & Month(Date()) & "-" & Day(Date()) & "-" & "v." & nVersion
			    vFileLine(nLine - 1) = vFileLine(nLine - 1) & "," & strVersion
				vCfgList(nCfg,1) = vCfgList(nCfg,1) & "," & strVersion
			Else 
			    strVersion = Year(Date) & "-" & Month(Date()) & "-" & Day(Date()) & "-" & "v.0" & nVersion
			    vFileLine(nLine - 1) = vFileLine(nLine - 1) & "," & strVersion
				vCfgList(nCfg,1) = vCfgList(nCfg,1) & "," & strVersion
			End If
		Else 
		   strVersion = Year(Date) & "-" & Month(Date()) & "-" & Day(Date()) & "-" & "v.01"
           vFileLine(nLine - 1) = "Version = " & strVersion
		End If 
		Call DeleteFileGroup(strDirectoryConfig & "\CfgList.txt", strCfg, 0)
		vLine(0) = "[" & strCfg & "]"
		Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vLine, UBound(vLine),2,0)
		Call WriteArrayToFile(strDirectoryConfig & "\CfgList.txt", vFileLine, UBound(vFileLine),2,0)
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
' Function DeleteFileLineGroup - Returns number of lines int the text file
'---------------------------------------------------------------------------
 Function DeleteFileLineGroup(strFileName, strGroup1, strParam1,nDebug_)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	DeleteFileLineGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	Redim vFileLines(nIndex)
	Call TrDebug ("DeleteFileLineGroup: String containing """ & strParam1 & """ under Group [" & strGroup1 & "] WILL BE DELETED", "", objDebug, MAX_LEN, 1, 1)					
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
				If nGroupSelector = 1 and InStr(strLine,strParam1) <> 0 Then 
					Call TrDebug ("DeleteFileLineGroup: String containing """ & strParam1 & """", "WAS DELETED", objDebug, MAX_LEN, 1, 1)					
				Else
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					nIndex = nIndex + 1
				End If
		End Select
	Loop
	objDataFileName.Close
	Call WriteArrayToFile(strFileName,vFileLines, UBound(vFileLines),1,nDebug)
    DeleteFileLineGroup = True
End Function	
'---------------------------------------------------------------------------
'   Function AddFileLineInGroup(strFileName, strGroup1, strParam1,nDebug_)
'---------------------------------------------------------------------------
 Function AddFileLineInGroup(strFileName, strGroup1, strParam1,nDebug_)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	AddFileLineInGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	Redim vFileLines(nIndex)
	Call TrDebug ("AddFileLineInGroup: String """ & strParam1 & """ under Group [" & strGroup1 & "] WILL BE ADDED", "", objDebug, MAX_LEN, 1, 1)					
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
								Call TrDebug ("AddFileLineInGroup: String """ & strParam1 & """", "WAS ADDED", objDebug, MAX_LEN, 1, 1)					
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
		Call TrDebug ("AddFileLineInGroup: String """ & strParam1 & """", "WAS ADDED", objDebug, MAX_LEN, 3, 1)					
	End If 
	objDataFileName.Close
	Call WriteArrayToFile(strFileName,vFileLines, UBound(vFileLines),1,nDebug)
    AddFileLineInGroup = True
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
'---------------------------------------------------------------------------
'   Function IsFile(strFile)
'---------------------------------------------------------------------------
Function IsFile(strFile)
Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Do
		If objFSO.FileExists(strFile) Then IsFile = 1 : Exit Do : End If
		If objFSO.FolderExists(strFile) Then IsFile = 2 : Exit Do : End If
		IsFile = 0
		Exit Do
	Loop
	Set objFSO = Nothing
End Function 
'---------------------------------------------------------------------------
'   Function IsFile(strFile)
'---------------------------------------------------------------------------
Function BrowseForFile()
    With CreateObject("WScript.Shell")
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
        Dim tempName : tempName = fso.GetTempName() & ".hta"
        Dim path : path = "HKCU\Volatile Environment\MsgResp"
        With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
            "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
            "close();</script>"
            .Close
        End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
    End With
End Function
'-----------------------------------------------------------------------------------------
'      Function Displays a Message with Continue and No Button. Returns True if Continue
'-----------------------------------------------------------------------------------------
 Function IE_CONT_MULT (vIE_Scale, strTitle, vLine, ByVal nLine, vButton, objIEParent, nDebug)
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim g_objIE, objShell
    Set g_objIE = Nothing
    Set objShell = Nothing
	Dim N, xTab
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_CONT_MULT = False
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call IE_Hide(objIEParent)
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nRatioX = intX/1920
    nRatioY = intY/1080
	CellW = Round(500 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	nTab = Round(20 * nRatioX,0)
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
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
			
	Next
    If UBound(vButton) < 5 Then N = UBound(vButton) + 1 Else N = 4 End If	
    xTab = 	Round((CellW - 2 * nTab - N * nButtonX)/(N - 1),0)
	For nButton = 0 to N - 1
    strHTMLBody = strHTMLBody &_
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nTab + nButton * (xTab + nButtonX)   & "px; bottom: " & BottomH & "px' name='Button_"& nButton & "' onclick=document.all('ButtonHandler').value='Button_"& nButton & "';>" & vButton(nButton) & "</button>" 
	Next
	strHTMLBody = strHTMLBody & "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.Visible = True
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strMyPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then strMyPID = Split(strLine,""",""")(1)
	     ' Call TrDebug("READ TASK PID:" , strLine, objDebug, MAX_LEN, 1, 1)
    Next
    If strMyPID = "" Then Call GetAppPID(strMyPID, "iexplore.exe")
	objShell.AppActivate strMyPID										
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		If InStr(g_objIE.Document.All("ButtonHandler").Value, "Button") > 0 Then  
				IE_CONT_MULT = Split(g_objIE.Document.All("ButtonHandler").Value,"_")(1)
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
	    End If
		Wscript.Sleep 300
		Loop
		Call IE_UnHide(objIEParent)
End Function
'---------------------------------------------------------------------------------------
'   Function FindAndReplaceExactStrInFile(strFile, strFind, strNewLine, nDebug)
'   Search for the First Line which contains "strFind" and Replaces whole Line with "strNewLine"
'---------------------------------------------------------------------------------------
Function FindAndReplaceExactStrInFile(strFile, strFind, strNewLine, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Const FOR_WRITING = 1
	FindAndReplaceExactStrInFile = False
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLine is number of lines counted like 1,2,...,n
	LineNumber = GetExactObjectLineNumber( vFileLine, nFileLine, strFind)
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": FindAndReplaceExactStrInFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	vFileLine(LineNumber - 1) = strNewLine
	If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then FindAndReplaceExactStrInFile = True End If
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
			Call TrDebug ("GetVariable: CAN'T FIND VARIABLE: " & strVar, "Dim1=" & Dim1 & " Dim2=" & Dim2, objDebug, MAX_LEN, 1, nDebug)
			Exit Do 
		End If
		If nResult > 0 and InStr(vFileLines(nResult - 1),"=") = 0 Then 
			GetVariable = "NULL" 
			Call TrDebug ("GetVariable: ERROR: WRONG DEFINITION OF THE VARIABLE " & strVar, "", objDebug, MAX_LEN, 1, 1)
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
    Dim nIndex
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
		Call TrDebug ("GetMyClass: TOTAL OBJECTS IN CLASS: " & vClassIndex(n), nObj, objDebug, MAX_LEN, 1, nDebug)
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
								Call TrDebug ("GetMyClass: LOAD PROPERTIES FOR [Class_" & vClassIndex(n) & "]", "", objDebug, MAX_LEN, 3, nDebug)
								nParam = 1
								nGroupSelector = 1
							Case Else
								nGroupSelector = 0
						End Select
					Case Else	
						If nGroupSelector = 1 Then 
							Call TrDebug ("GetMyClass:" & strLine, "", objDebug, MAX_LEN, 1, nDebug)					
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
									Call TrDebug ("GetMyClass: LOAD DATA FOR: " & strLine, "", objDebug, MAX_LEN, 3, nDebug)
									nParam = 0
									nGroupSelector = 1
								Case Else
									' If nGroupSelector = 1 Then Exit For
									nGroupSelector = 0
							End Select
						Case Else	
							If nGroupSelector = 1 Then 
								Call TrDebug ("GetMyClass:" & strLine, "", objDebug, MAX_LEN, 1, nDebug)					
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
	Call TrDebug ("pIndex: ClassName: " & strClassID & " ClassID = " & ClassID, "", objDebug, MAX_LEN, 1, nDebug)					
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
	Call TrDebug ("SetMyObject: ClassName: " & strClassID & " ClassID = " & ClassID, "", objDebug, MAX_LEN, 1, nDebug)	
	'---------------------------------------
	'   COUNT Params in Given Class
	'---------------------------------------
	nParam = 0
	For i = 0 to UBound(vClass,2)-1
		If InStr(vClass(ClassID,i),"Param") > 0 Then 
			nParam = nParam + 1 
		End If
	Next
	Call TrDebug ("SetMyObject: Found " & n & " Properties for class " & strClassID,"", objDebug, MAX_LEN, 1, 1)
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
	Call TrDebug ("ReplaceFileLineInGroup: String """ & strParam1 & """ under Group [" & strGroup1 & "] WILL BE ADDED", "", objDebug, MAX_LEN, 1, nDebug)					
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
					        If InStr(strLine, strParamOld & " =") > 0 Then strLine = strParam1
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
'##############################################################################
'      Function ASKS USER TO ENTER PASSWORD
'##############################################################################
 Function IE_PromptLoginPassword (objParentWin, vIE_Scale, vLine, nLine, ByRef strUsername, ByRef strPassword, Confirm, nDebug )
    Dim strPID
	Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim g_objIE, g_objShell
	intX = 1920
	intY = 1080
	Dim IE_Menu_Bar
	Dim  IE_Border
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	strPassword = "DO NOT MATCH"
	IE_PromptLoginPassword = False	
	
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nHeader = Round (12 * nRatioY,0)
	LineH = Round (12 * nRatioY,0)
	nTab = 20
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellW = Round(330 * nRatioX,0)
	ColumnW1 = Round(150 * nRatioX,0)
	CellH = 2 * (nLine + 7) * LineH
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
	If Confirm Then 
	    CellH = CellH + 3 * 2 * LineH 
		nOrder = 1
    Else 
	    nOrder = 0
	End If
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------		
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	InputBGColor = HttpBgColor4
	MainTextColor = HttpTextColor1
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = BackGroundColor
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	'----------------------------------------------------------
	'    TITLE
	'----------------------------------------------------------
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: "& HttpBgColor2 & ";" &_
		"width: " & CellW & "px;"">" & _
		"<tbody>"	
	For nInd = 0 to nLine - 1
		 If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
		strHTMLBody = strHTMLBody &_
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & CellW & """>" & _
				"<p style=""text-align: center; font-family: 'arial narrow';font-size: " & nFontSize_12 & ".0pt; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" &_
			"</td>" &_
		"</tr>"
	Next
	strHTMLBody = strHTMLBody & "</tbody></table>"
	
	'----------------------------------------------------------
	'    MAIN FORM FOR ENTERING LOGON AND PASSWORD
	'----------------------------------------------------------
	TableW = CellW
	ColumnW_1 = 3 * Int(TableW/3)
	ColumnW_2 = TableW - ColumnW_1
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: " & (nLine + 1) * LineH * 2 & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: none;;" &_
		"width: " & TableW & "px;"">" & _
		"<tbody>"		
	'----------------------------------------------------	
	'  ROW 1
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">LOGIN NAME</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		"</td>" &_
	"</tr>"		
	'----------------------------------------------------	
	'  ROW 2
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """ >" & _
			"<input name=UserName style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=25 tabindex=1>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='EXIT' name='Cancel' AccessKey='C' tabindex=" & nOrder + 4 & " onclick=document.all('ButtonHandler').value='Cancel';>CANCEL</button>" & _		    
		"</td>" &_		
	"</tr>"
	'----------------------------------------------------	
	'  ROW 3 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  ROW 4
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">PASSWORD</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
		"</td>" &_
	"</tr>"			
	'----------------------------------------------------	
	'  ROW 5
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
			"<input id='PASSWD' name=Password style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=2 " & _
			"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='OK' name='OK' AccessKey='C' tabindex=" & nOrder + 3 & " onclick=document.all('ButtonHandler').value='OK';>SIGN IN</button>" & _		    
		"</td>" &_				
	"</tr>"
	'----------------------------------------------------	
	'  ROW 6 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  CONFIRM PASSWORD ROW
	'----------------------------------------------------
	If Confirm Then 
		'----------------------------------------------------	
		'  ROW 7
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
				"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
				"; font-weight: bold;"">CONFIRM PASSWORD</p>" &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"</td>" &_
		"</tr>"			
		'----------------------------------------------------	
		'  ROW 8
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
				"<input id='PASSWD2' name=Password2 style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
				"; border-radius: 10px " &_
				"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=3 " & _
				"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """>" & _
			"</td>" &_				
		"</tr>"
	End If
	strHTMLBody = strHTMLBody & "</tbody></table>"
    strHTMLBody = strHTMLBody &_
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = "Login and Password"
'	g_objIE.document.getElementById("OK").style.borderRadius = "10px"
'	g_objIE.document.getElementById("EXIT").style.borderRadius = "10px"
	g_objIE.document.getElementById("OK").style.backgroundcolor = ButtonColor
	g_objIE.document.getElementById("EXIT").style.backgroundcolor = ButtonColor
	If Confirm Then
	    g_objIE.Document.getElementById("OK").innerHTML = "OK"
	Else 
	   	g_objIE.Document.getElementById("OK").innerHTML = "SIGN IN"
	End If
	
	g_objIE.Visible = False
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy	
	Set g_objShell = WScript.CreateObject("WScript.Shell")
	Call IE_Unhide(g_objIE)
	Call IE_GetPID(strPID, g_objIE.document.Title & " - " & IE_Window_Title, nDebug)
	g_objShell.AppActivate strPID
	g_objIE.Document.All("UserName").Focus
	g_objIE.Document.All("UserName").Value = strUsername
'    g_objIE.Document.body.addeventlistener "keydown", GetRef("KeyLA"), false
	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case strNothing
			Case "Cancel"
				' The user clicked Cancel. Exit the loop
				IE_PromptLoginPassword = False				
				Exit Do
			Case "OK"
				' strUsername = g_objIE.Document.All("Username").Value
				Select Case Confirm
					Case True
						if g_objIE.Document.All("Password").Value = g_objIE.Document.All("Password2").Value  and _
						   InStr(g_objIE.Document.All("Password").Value," ") = 0 and _
						   g_objIE.Document.All("Password").Value <> "" Then 
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = g_objIE.Document.All("Password").Value
							IE_PromptLoginPassword = True
							Exit Do
						Else
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = "DO NOT MATCH"
							IE_PromptLoginPassword = True
							Exit Do
						End If 
					Case False
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = g_objIE.Document.All("Password").Value
							IE_PromptLoginPassword = True
							Exit Do
				End Select
		End Select
	    Wscript.Sleep 200
    Loop
	g_objIE.quit
	Wscript.Sleep 200
	Set g_objIE = Nothing
	Set g_objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function IE_GetPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function IE_GetPID(ByRef strPID, strWinTitle, nDebug)
Dim strLine, strCmd, vCmdOut
    IE_GetPID = False
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & strWinTitle & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD,nDebug)
    strPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then 
	        strPID = Split(strLine,""",""")(1)
			IE_GetPID = True
			Exit For
	        Call TrDebug("IE_GetPID IE Window (" & strWinTitle & ") PID: " & strPID  , "", objDebug, MAX_LEN, 1, nDebug)
		End If
    Next
End Function