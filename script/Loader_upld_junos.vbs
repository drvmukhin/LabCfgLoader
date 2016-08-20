#$language = "VBScript"
#$interface = "1.0"

'----------------------------------------------------------------------------------
'	JUNIPER LABLOADER SCRIPT: UPDATE JUNOS IMAGE
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 75
Const GLOBAL_CFG = "grp-global-original.conf"
Const RE_CFG = "grp-"
' Define global array which stores parameters of all my objects per class
Dim vObjects
' Define global array which keeps properties of all my Classes' 
Dim vClass
' Define global array for JunosSW objects
Dim objMain, objMinor

Dim nResult
Dim strLine
Dim nOverwrite
Dim strMonthMaxFileName, strFileString, strSkip, strFileButton, strFileInventory, strFileSession
Dim strDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke
Dim strDeviceID, strAccountID
Dim nDebug
Dim nIndex, nInd, nCount
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile
Dim vSession(30)
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand, nRollBack, bConnect
Dim Platform
Dim objTab_L, objTab_Catalog
Dim vCopy, vUpdate
Dim vSessionCRT
Dim nImage
Dim strFileSettings
Dim vDelim, vParamNames, bDownlodOnly
Const SECURECRT_SESSION = "Node Session"
vDelim = Array("=",",",":")	
nDebug = 0
nInfo = 1
strFileSettings = "settings.dat"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = CreateObject("WScript.Shell")
Sub Main()
'------------------------------------------------------------------
'	CHECK NUMBER OF ARGUMENTS AND EXIT IF LESS THEN 3
'------------------------------------------------------------------
	If crt.Arguments.Count < 6 Then
			MsgBox "ERROR: Wrong number of arguments" & chr(13) &_
			"ARG1: Full Path to junos_catalog.dat file" & chr(13) &_
			"ARG2: Full Path to settings.dat file" & chr(13) &_
			"ARG3: Full Path for Loader Work Folder" & chr(13) &_
			"ARG4: SecureCRT session ID (nSession)"	 & chr(13) &_
            "ARG5: Login Name to access image catalog" & chr(13) &_
			"ARG6: Password to access catalog" & chr(13)&_
			"ARG7: Download Only -d"
		crt.quit
		Exit Sub
	End If
	strCatalogFile = crt.Arguments(0)
	strFileSettings = crt.Arguments(1)
	strDirectoryWork = crt.Arguments(2)	
	nSession = crt.Arguments(3)
	CatalogLogin = crt.Arguments(4)
	CatalogPassword = crt.Arguments(5)
	bDownlodOnly = False
	If crt.Arguments.Count = 7 Then 
	    If crt.Arguments(6) = "-d" Then bDownlodOnly = True End If
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
						Set objDebug = objFSO.OpenTextFile(strDirectoryWork & "\Log\" & "debug-terminal-" & nSession & ".log",ForWriting,True)
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
	strLaunch = strDirectoryWork & "\bin\tail.exe -f " & strDirectoryWork & "\log\debug-terminal-" & nSession & ".log"
'	If Not GetAppPID(strPID, strParentPID, "tail.exe") Then 
		objEnvar.run (strLaunch)
'	Else
'		Call FocusToParentWindow(strPID)
'	End If
	Call TrDebug_No_Date ("GetMyPID: PID = " & strPID & " ParentPID = " & strParentPID,"",objDebug, MAX_LEN, 1, nDebug)
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
    Dim strHostL, strLogin, strSessionL, strFolder
	bSuccess = False
	bConnect = False
	'--------------------------------------------------------------------------------
	'          START MAIN CYCLE
	'--------------------------------------------------------------------------------
	'--------------------------------------------------------------------------------
	'          GET NAME OF THE TELNET SESSIONS
	'--------------------------------------------------------------------------------
	nInd = Int(nSession)
	    Do
			'------------------------------------------------------------------
			'          GET NAME OF THE TELNET SESSIONS
			'------------------------------------------------------------------
			strFolder = Split(vSessionCRT(nInd), ",")(1) & "/root/"
			strSessionL = Split(vSessionCRT(nInd), ",")(2)
			strPlatform = Split(vSessionCRT(nInd), ",")(3) 
			strLogin = Split(vSessionCRT(nInd), ",")(4)
			'------------------------------------------------------------------
			'	LOG MAIN VARIABLES
			'------------------------------------------------------------------
			Call TrDebug_No_Date ("TelnetScript: strFolder, strSessionL: " & strFolder & strSessionL & ", " & strHostL,"", objDebug, MAX_LEN, 1, nDebug)						
			'------------------------------------------------------------------
			'  Start SSH session to Node
			'------------------------------------------------------------------
			On Error Resume Next
			Err.Clear
			Set objTab_L = crt.Session.ConnectInTab("/S " & strFolder & strSessionL)
			If Err.Number <> 0 Then 
				Call  TrDebug_No_Date ("ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, 1)
				bConnect = False
				Exit Do
			End If
			On Error Goto 0
			bConnect = True
			Call TrDebug_No_Date ("START UPDATING JUNOS IMAGE ON " & UCase(strSessionL), "", objDebug, MAX_LEN, 3, 1)		
			objTab_L.Caption = strSessionL
			objTab_L.Screen.Synchronous = True
			'--------------------------------------------------------------------------------
			'  Get actual host name of the box
			'--------------------------------------------------------------------------------
			objTab_L.Screen.Send chr(13)
			strLine = objTab_L.Screen.WaitForString ("%")
			objTab_L.Screen.Send "cli" & chr(13)
			strLine = objTab_L.Screen.WaitForString (">")			
			'-------------------------------------------------------
			' CHECK THE PLATFORM TYPE, HOSTNAME and VERSION
			'-------------------------------------------------------
			objTab_L.Screen.Send "show version |no-more" & chr(13)
			objTab_L.Screen.WaitForString "show version |no-more"
			strLine = objTab_L.Screen.ReadString ("JUNOS")
			vLines = Split(strLine,chr(13))
			strVersion = "N/A"
			strHostL = ""
			For Each strLine in vLines
			    If Len(strLine) > 1 Then 
					strLine = Right(strLine,Len(strLine) - 1)
					Select Case Split(strLine,": ")(0)
						Case "Model"
							strPlatform = Split(strLine,": ")(1)
						Case "Hostname"
							strHostL = Split(strLine,": ")(1)
						Case "Junos"
							strVersion = Split(strLine,": ")(1)
					End Select
				End If
			Next
			Call  TrDebug_No_Date ("Hostname: " & strHostL, "", objDebug, MAX_LEN, 1, nInfo)
		    Call  TrDebug_No_Date ("Model: " & strPlatform, "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("Current Junos: " & strVersion, "", objDebug, MAX_LEN, 1, nInfo)
			objTab_L.Screen.WaitForString strHostL & "> "
			'------------------------------------------------------------------
			'          GET PLATFORM PROPERTIES
			'------------------------------------------------------------------			
			nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Supported_Platforms","","",0)
			For nInd = 0 to nInventory - 1
				If Split(vLines(nInd),",")(0) = strPlatform Then 
				    strPlatformType = Split(vLines(nInd),",")(1) 
					nRe = Int(Split(vLines(nInd),",")(2)) - 1
					Select Case nRe
					    Case 0
					        vRe = Array(Split(vLines(nInd),",")(3))  
							vReUpdate = Array("")  
					    Case 1
						    vRe = Array(Split(vLines(nInd),",")(3),Split(vLines(nInd),",")(4))
							vReUpdate = Array(Split(vLines(nInd),",")(3),Split(vLines(nInd),",")(4))
					End Select
					Exit For 
				End If
			Next
			Call  TrDebug_No_Date ("Host " & strSessionL & " was qualified as: (" & strPlatformType & ") " & UCase(strPlatform) , "", objDebug, MAX_LEN, 1, nInfo)
			nImage = -1
            For nInd = 0 to UBound(objMain,1)			
			   If strPlatformType = objMain(nInd,pIndex(0,"Platform")) and objMain(nInd,pIndex(0,"Status")) = "Active" Then nImage = nInd
			Next
            If nImage = -1 Then 
   			    Call  TrDebug_No_Date ("No Junos image was scheduled for platform " & strPlatform, "SKIP", objDebug, MAX_LEN, 1, nInfo)
			    Exit Do
			End If			
			'------------------------------------------------------------------
			'          GET IMAGE FILE NAME
			'------------------------------------------------------------------			
			Folder1 = objMain(nImage,pIndex(0,"Folder1"))
			Folder2 = objMain(nImage,pIndex(0,"Folder2"))
			ImageTemplate = objMain(nImage,pIndex(0,"ImageTemplate1"))
			strMinor = objMain(nImage,pIndex(0,"Minor"))
			strMajor = objMain(nImage,pIndex(0,"Main"))
			ImageFile = ""
			'--------------------------------------------------------------------------------
			'          CONNECT TO CATALOGUE
			'--------------------------------------------------------------------------------
			strSessionCatalog = objMain(nImage,pIndex(0,"SecureCRT_Session"))
			Call TrDebug_No_Date ("TelnetScript: " & objMain(nSession,pIndex(0,"Name")) & " : " & strSessionCatalog,"", objDebug, MAX_LEN, 1, nDebug)						
			'--------------------------------------------------------------------------------
			'  Start SSH session to Catalogue
			'--------------------------------------------------------------------------------
			Call TrDebug_No_Date ("CONNECTING TO CATALOGUE AT " & strSessionCatalog,"", objDebug, MAX_LEN, 1, nInfo)						
			On Error Resume Next
			Err.Clear
			Set objTab_Catalog = crt.Session.ConnectInTab("/S " & strSessionCatalog)
			If Err.Number <> 0 Then 
				Call  TrDebug_No_Date ("CAN'T CONNECT TO " & strSessionCatalog & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, nInfo)
				bConnectCatalog = False
				Exit Do
			End If
			On Error Goto 0
			bConnectCatalog = True
			TryTemplate2 = False
			objTab_Catalog.Caption = Split(strSessionCatalog,"/")(1)
			objTab_Catalog.Screen.Synchronous = True			
			objTab_Catalog.Screen.Send chr(13)
			objTab_Catalog.Screen.WaitForString ">"
			objTab_Catalog.Screen.Send "bash" & chr(13)	
			objTab_Catalog.Screen.WaitForString ("]$")
			objTab_Catalog.Screen.Send "cd " & Folder1 & strMajor & "/" & Folder2 & strMinor & "/ship/" & chr(13)	
			objTab_Catalog.Screen.WaitForString ("]$")
			objTab_Catalog.Screen.Send "ls -l " & ImageTemplate & chr(13)
			objTab_Catalog.Screen.WaitForString (ImageTemplate & chr(13))
			strLine = objTab_Catalog.Screen.ReadString ("[")
			vLines = Split(strLine,chr(13))
			For each strLine in vLines
				If InStr(strLine,"No such file or directory") > 0 Then 
				   Call TrDebug_No_Date ("Can't find image file in catalog " ,"ERROR", objDebug, MAX_LEN, 1, nInfo)
				   Call TrDebug_No_Date ("Will Try Second Template " ,"", objDebug, MAX_LEN, 1, nInfo)
				   TryTemplate2 = True
				   objTab_Catalog.Screen.WaitForString ("]$")
				   Exit For
				End If
			    strLine = RTrim(strLine)
				strLine = RTrim(Split(strLine,"->")(0))
				strLine = Split(strLine," ")(UBound(Split(strLine," ")))
			    If Len(strLine) > 1 Then 
				   ImageFile = strLine	
				   Call TrDebug_No_Date ("TARGET IMAGE FILE: " & ImageFile,"FOUND", objDebug, MAX_LEN, 1, nInfo)
				   objTab_Catalog.Screen.WaitForString ("]$")
				   Exit For
			    End If
			Next

            '-------------------------------------------
            '   TRY SECOND TEMPLATE NAME IF REQUIRED
            '-------------------------------------------	
            If TryTemplate2 and  pIndex(0,"ImageTemplate2") => 0 Then 
			    ImageTemplate = objMain(nImage,pIndex(0,"ImageTemplate2"))
				objTab_Catalog.Screen.Send "ls -l " & ImageTemplate & chr(13)
				objTab_Catalog.Screen.WaitForString (ImageTemplate & chr(13))
				strLine = objTab_Catalog.Screen.ReadString ("[")
				vLines = Split(strLine,chr(13))
				For each strLine in vLines
					If InStr(strLine,"No such file or directory") > 0 Then 
						   Call TrDebug_No_Date ("Can't find image file in catalog " ,"ERROR", objDebug, MAX_LEN, 1, nInfo)
						   objTab_Catalog.Screen.WaitForString ("]$")
						   Exit Do
					End If
					strLine = RTrim(strLine)
					strLine = RTrim(Split(strLine,"->")(0))
					strLine = Split(strLine," ")(UBound(Split(strLine," ")))
					If Len(strLine) > 1 Then 
					   ImageFile = strLine	
					   Call TrDebug_No_Date ("TARGET IMAGE FILE: " & ImageFile,"FOUND", objDebug, MAX_LEN, 3, nInfo)
					   objTab_Catalog.Screen.WaitForString ("]$")
					   Exit For
					End If
				Next
			End If
			'---------------------------------------------
			'   DISCONNECT FROM CATALOGUE
			'---------------------------------------------			
			If bConnectCatalog Then 		
				objTab_Catalog.Session.Disconnect
				objTab_Catalog.Close
			End If
            'Call  TrDebug_No_Date ("Going to install Image: " & ImageFile, "", objDebug, MAX_LEN, 1, nInfo)
			'---------------------------------------------
			'   Download ImageFile to /var/tmp/
			'---------------------------------------------
			LocalFolder = "/var/tmp/"
'			LocalFolder = ""
			vCopy = Array("(yes/no)? ", "assword:", "100%", "file-fetch failed", "ermission denied","filesystem is full")
			bCopy = False
			objTab_L.Screen.Send "copy file " & CatalogLogin & "@svl-jtac-tool01:/" & Folder1 & strMajor & "/" & Folder2 & strMinor & "/ship/" & ImageFile & " " & LocalFolder & ImageFile & chr(13)
			Do
				nResult = objTab_L.Screen.WaitForStrings (vCopy, 600)
				Select Case nResult
					Case 1
						objTab_L.Screen.Send "yes" & chr(13) 
					Case 2
						objTab_L.Screen.Send CatalogPassword & chr(13) 
					Case 3
						Call  TrDebug_No_Date ("New Image Copied to the Node: " & strSessionL , "OK", objDebug, MAX_LEN, 1, nInfo)
						bCopy = True
						Exit Do
					Case 4
						Call  TrDebug_No_Date ("Image: " & ImageFile, "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
						Exit Do
					Case 5 
						Call  TrDebug_No_Date ("WRONG Login/Password ", "ERROR", objDebug, MAX_LEN, 1, nInfo)
						Exit Do
					Case 6
					    Call  TrDebug_No_Date ("Filesystem is full ", "ERROR", objDebug, MAX_LEN, 1, nInfo)
						Exit Do
					Case 0 
						Call  TrDebug_No_Date ("Copy Image to :" & strSessionL, "TIMEOUT", objDebug, MAX_LEN, 1, nInfo)
						Exit Do					    
				 End Select
			 Loop 
			 objTab_L.Screen.WaitForString strHostL & "> "
			 If Not bCopy Then
			    Call  TrDebug_No_Date ("Copy Image to :" & strSessionL, "FAILED", objDebug, MAX_LEN, 1, nInfo)
			    Exit Do
			 End If
			'---------------------------------------------
			'   Check MD5
			'---------------------------------------------
			vCopy = Array("(yes/no)? ", "assword:", "100%", "file-fetch failed", "ermission denied","filesystem is full")
			bCopyMD5 = False
			objTab_L.Screen.Send "copy file " & CatalogLogin & "@svl-jtac-tool01:/" & Folder1 & strMajor & "/" & Folder2 & strMinor & "/ship/" & ImageFile & ".md5 /var/tmp/" & ImageFile & ".md5 " & chr(13)
			Do
				nResult = objTab_L.Screen.WaitForStrings (vCopy, 600)
				Select Case nResult
					Case 1	
					    objTab_L.Screen.Send "yes" & chr(13) 
					Case 2
						objTab_L.Screen.Send CatalogPassword & chr(13) 
					Case 3
						Call  TrDebug_No_Date ("MD5 File Copied to the Node: " & strSessionL , "OK", objDebug, MAX_LEN, 1, nInfo)
						bCopyMD5 = True
						Exit Do
					Case 4
						Call  TrDebug_No_Date ("MD5 File: " & ImageFile & ".md5", "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
						Exit Do
					Case 5 
						Call  TrDebug_No_Date ("WRONG Login/Password ", "ERROR", objDebug, MAX_LEN, 1, nInfo)
						Exit Do
					Case 6
					    Call  TrDebug_No_Date ("Filesystem is full ", "ERROR", objDebug, MAX_LEN, 1, nInfo)
					Case 0 
						Call  TrDebug_No_Date ("Copy Image to :" & strSessionL, "TIMEOUT", objDebug, MAX_LEN, 1, nInfo)
						Exit Do					    
				 End Select
			Loop 
			objTab_L.Screen.WaitForString strHostL & "> "
			bMD5 = True
            If bCopyMD5 Then  
			    objTab_L.Screen.Send "file checksum md5 " & LocalFolder & ImageFile & chr(13)
				objTab_L.Screen.WaitForString "= "
				strLine = objTab_L.Screen.ReadString (">")
				MD5CheckSum = Split(strLine,chr(13))(0)
				Call  TrDebug_No_Date ("Calculate MD5 Check Sum: " & MD5CheckSum, "", objDebug, MAX_LEN, 1, nInfo)
				objTab_L.Screen.Send "file show /var/tmp/" & ImageFile & ".md5" & chr(13)
				strLine = objTab_L.Screen.ReadString (">")
				MD5 = Split(strLine,chr(13))(1)
				MD5 = Right(MD5, Len(MD5) - 1)
				Call  TrDebug_No_Date ("Read MD5 From Catalogue: " & MD5, "", objDebug, MAX_LEN, 1, nInfo)
				If MD5 = MD5CheckSum Then bMD5 = True Else bMD5 = False
			End If 
			If Not bMD5 Then 
				Call  TrDebug_No_Date ("MD5 Check ", "FAILED", objDebug, MAX_LEN, 1, nInfo)
				Exit Do
			End If
			Call  TrDebug_No_Date ("MD5 Check ", "OK", objDebug, MAX_LEN, 1, nInfo)
            If bDownlodOnly	Then 
				Call  TrDebug_No_Date ("Download Only Flag Used ", "UPDATE SKIP", objDebug, MAX_LEN, 1, nInfo)
				Exit Do
            End If 
			'---------------------------------------------
			'   Initiate Junos Image Update
			'---------------------------------------------
			vUpdate = Array(">", "reboot","Saving state for rollback")
			For Each strRe in vReUpdate
			'	objTab_L.Screen.Send "request system software add " & LocalFolder & ImageFile & " " & strRe & chr(13)
				objTab_L.Screen.Send "request system software add " & LocalFolder & ImageFile & " " & strRe & " force" & chr(13)				
				nResult = objTab_L.Screen.WaitForStrings (vUpdate, 900)
				Select Case nResult
					Case 2, 3
					  Call  TrDebug_No_Date ("UPDATE " & strSessionL & " " & strRE, "OK", objDebug, MAX_LEN, 1, nInfo)
					  objTab_L.Screen.WaitForString ">"
					  bSuccess = True
					Case 1
					  Call  TrDebug_No_Date ("UPDATE " & strSessionL & " " & strRE, "OK", objDebug, MAX_LEN, 1, nInfo)					  
					  Call  TrDebug_No_Date ("UPDATE Warning: Check current junos version manually ", "", objDebug, MAX_LEN, 1, nInfo)
					  bSuccess = True
					Case 0 
					  Call  TrDebug_No_Date ("UPDATE " & strSessionL & " " & strRE, "TIMEOUT", objDebug, MAX_LEN, 1, nInfo)
					  Call  TrDebug_No_Date ("UPDATE: Check current junos version manually ", "", objDebug, MAX_LEN, 1, nInfo)
				  Case Else
				End Select
			Next
			'---------------------------------------------
			'   Initiate Reboot of the node
			'---------------------------------------------
            If bSuccess Then
                Call  TrDebug_No_Date ("REBOOT NODE " & strSessionL , "REBOOTING", objDebug, MAX_LEN, 1, nInfo)			
				Select Case nRe
				   Case 0
					  objTab_L.Screen.Send "request system reboot" & chr(13)
				   Case 1
					  objTab_L.Screen.Send "request system reboot both-routing-engines" & chr(13)
				End Select			
				objTab_L.Screen.WaitForString "(no)"
				objTab_L.Screen.Send "yes" & chr(13)
				objTab_L.Screen.WaitForString strHostL & ">"
				Exit Do
			End If
			Exit Do
		Loop   ' END MAIN CYCLE
	'---------------------------------------------
	'   DISCONNECT
	'---------------------------------------------			
	If bConnect Then 		
		objTab_L.Session.Disconnect	
	End If
	Call TrDebug_No_Date ("JOB DONE ", "", objDebug, MAX_LEN, 3, 1)	
	If IsObject(objDebug) Then objDebug.close : End If
	' objEnvar.Run "notepad.exe " & strDirectoryWork & "\Log\debug-terminal.log"	
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