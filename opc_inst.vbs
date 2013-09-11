Option Explicit

'*********************************************************
' Purpose: Checks if strKey is obsolete.
' Inputs: strKey: name of key.
'         strObsoleteKeys: aray of obsolete keys.
' Returns: If key is obsolete, return True.
'          Otherwise return False.
'*********************************************************
Private Function IsObsolete (ByRef strKey, ByRef strObsoleteKeys)
  Dim i

  For i = 0 To UBound (strObsoleteKeys)
    If (StrComp (strKey, strObsoleteKeys(i), 1) = 0) Then
      IsObsolete = True
      Exit Function
    End If
  Next

  IsObsolete = False
End Function

'*********************************************************
' Purpose: To deinstall the hofix deployment bundle
' Inputs: - 
' Returns: True if it is successfully deinstalled
'*********************************************************

Private Function RemHotfixDpl (ByVal strPkg) 

    Dim MyShell, MySysEnv, agentdir, strCmd, strOut, match_hotfix, objRegExp
'    RemHotfixDpl = False
    
'    Set MyShell = CreateObject("WScript.Shell")
'    Set MySysEnv  = MyShell.Environment("SYSTEM")
'    agentdir = MySysEnv("OvInstallDir")
    
'    strCmd = """" & agentdir & "bin\ovdeploy"" -inv -inclbdl" 
      
'    If RunAndGetOutput (strCmd, strOut) Then
'      'create a new instance of the regular expression object
'      Set objRegExp = New RegExp

'      objRegExp.Pattern = strPkg & "_hotfix"   'apply the search pattern
'      objRegExp.Global =  False                ' match all instances if the serach pattern
'      objRegExp.IgnoreCase = True              ' ignore case

'     match_hotfix = objRegExp.Test(strOut)
'     If match_hotfix Then
'		strCmd = """" & agentdir & "bin\ovdeploy"" -remove -bdl " &  strPkg & "_hotfix"           
     
'      		strOut = ""
'      		Set objRegExp = Nothing
'      		If (RunAndGetOutput (strCmd, strOut)) Then
' 			  WScript.Echo strOut

'     	 	End If
'     Else
'      	 PrintMSG  "Hotfix deployment bundle for " &  strPkg & "not found currently on the system.", "Info" 
    
'     End If
'    End If
    	
'    Set objRegExp = Nothing	
'    Set MySysEnv = Nothing
'    Set MyShell = Nothing
    RemHotfixDpl = True	
End Function

'*********************************************************
' Purpose: Runs executable file and catches output.
' Inputs: strFile: name of executable file.
' Outputs: strOutputs: output of program.
' Returns: If program execution is successful, return True.
'          Otherwise return False.
'*********************************************************
Private Function RunAndGetOutput (ByRef strFile, ByRef strOutput)
  Dim MyShell, RetVal, Result
  Dim fso, dat
  Dim filename
  Const ForReading = 1

  ' establish error-handling
  On Error Resume Next
  RetVal = True

  Set MyShell = CreateObject ("WScript.Shell")

  ' executable has to be run with cmd.exe in order to
  ' be able to redirect output
  
  Set fso = CreateObject ("Scripting.FileSystemObject")
  filename = fso.GetTempName 
  Set fso = Nothing
  Result = MyShell.Run ("cmd.exe /c " & strFile _
                      & " >" & filename, 0, True)
                      
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    RunAndGetOutput = False
    Exit Function
  End If

  Set fso = CreateObject ("Scripting.FileSystemObject")
  
  ' if result is non-zero(failure) it is the error stream is written to the oot.out file
  
    If (Result <> 0) Then 
 	  Result = MyShell.Run ("cmd.exe /c " & strFile _
                     & " 2>out.out", 0, True)
	  Retval = False
    End If
  
    Set dat = fso.OpenTextFile (filename, ForReading, False)
    If (Err.Number <> 0) Then
      Err.Clear
      Set dat = Nothing
      Set fso = Nothing
      Set MyShell = Nothing
      RunAndGetOutput = False
      Exit Function
    End If
    strOutput = ""
    ' read everything from file
    strOutput = dat.ReadAll ()
    dat.Close

  fso.DeleteFile filename

  RunAndGetOutput = RetVal
  Set dat = Nothing
  Set fso = Nothing
  Set MyShell = Nothing
End Function 

'*********************************************************
' Purpose: Checks if strString is a comment.
' Inputs: strString: line from input file.
' Returns: If strString is a comment, return True.
'          Otherwise return False.
'*********************************************************
Private Function IsDefaultComment (ByRef strString)
  If (Mid (Trim (strString), 1, 1) = "/") _
     Or (Mid (Trim (strString), 1, 1) = "") Then
    IsDefaultComment = True
  Else
    IsDefaultComment = False
  End If
End Function

'*********************************************************
' Purpose: Extracts key and coresponding value from input string.
' Inputs:  strInpString: input string.
' Outputs: strKey: extracted key.
'          strValue: extracted value.
'          strProc: extracted process restriction (may be empty).
' Returns: If extraction is successful, return True.
'          Otherwise return False.
'*********************************************************

Private Function ParseSvcDiscConfig (ByRef strKey, ByRef strValue, ByRef strProc, _
                             ByRef strInpString)
  Dim strTempList, strTemp

  If (IsDefaultComment (strInpString)) Then
    ParseSvcDiscConfig = False
  Else
    strProc = ParseValPerProc (strInpString, strTemp)
    strTemp = Trim (strTemp)
    strTempList = Split (strTemp, " ", 2, 1)
    strKey = Trim (strTempList(LBound(strTempList)))

    If UBound(strTempList) <> LBound(strTempList) Then
      strValue = Trim (strTempList(UBound(strTempList)))
      ParseSvcDiscConfig = True
    Else
      strValue = ""
      ParseSvcDiscConfig = False
    End If
  End If

End Function
'*******************************************************************************
' Purpose: Convert old service discovery config file to new format.
'          INSTANCE_DELETION_THRESHHOLD and ACTION_TIMEOUT are migrated
' Inputs: strSourceFile: name of backed up JavaAgent.cfg file.
'         
' Returns: If conversion is successful, return True.
'          Otherwise return False.
'********************************************************************************

Public Function convert_svcdisc_config (strSourceFile)

  Const ForReading = 1
  Dim fso, dat, strLine, MyShell, strKey, strValue, outdat
  Dim RetVal
  Dim strProc
  Dim i
  Dim agent_id
  Dim set_core_id
  
  Dim strInstallDir
  Dim ObjFSO
  Dim strCmd
  
  On Error Resume Next
  
  'initialize falgs
  agent_id = 0
  set_core_id = False
  RetVal = False
  
  Set fso = CreateObject ("Scripting.FileSystemObject")
  Set MyShell = CreateObject ("WScript.Shell")
  Set dat = fso.OpenTextFile(strSourceFile, ForReading, False)
  
  If (Err.Number <> 0) Then
     Err.Clear
     Set fso = Nothing
     Set MyShell = Nothing
     Set dat = Nothing           
     PrintMSG  "Service discovery agent configuration file :  " & strSourceFile & " not found.", "Info"     
     RetVal = False	'file not found  
     convert_svcdisc_config = RetVal  
     Exit Function     
  End If 
    
  Do While (dat.AtEndOfStream <> True )   
   strLine = dat.ReadLine ()
   strProc = ""   
   strValue = ""
   If (ParseSvcDiscConfig (strKey, strValue, strProc, strLine)) Then	         
    If (strKey = "INSTANCE_DELETION_THRESHHOLD") Then      
      'Set the config values      
      If (strValue <> "") Then
       strCmd  = "cmd /c ovconfchg -ns agtrep -set INSTANCE_DELETION_THRESHOLD " & strValue        	    
       If (Not RunExternalScript (strCmd)) Then                             
        RetVal = False	'Setting INSTANCE_DELETION_THRESHOLD failed
        PrintMSG "Failed to set INSTANCE_DELETION_THRESHOLD under ns [agtrep]", "Warning"
       End If
      End If
    End If   	    	
    
    If (strKey = "ACTION_TIMEOUT") Then      
      'Set the config values
      If (strValue <> "") Then
       strCmd  = "cmd /c ovconfchg -ns agtrep -set ACTION_TIMEOUT " & strValue        	    
       If (Not RunExternalScript (strCmd)) Then                             
        RetVal = False	'Setting ACTION_TIMEOUT failed
        PrintMSG "Failed to set ACTION_TIMEOUT under ns [agtrep]", "Warning"
       End If
      End If   
    End If	          
   End If      
  Loop      
  
  convert_svcdisc_config = RetVal  
     
End Function

'*********************************************************
' Purpose: Checks if strString is a comment.
' Inputs: strString: line from input file.
' Returns: If strString is a comment, return True.
'          Otherwise return False.
'*********************************************************
Private Function IsComment (ByRef strString)
  If (Mid (Trim (strString), 1, 1) = "#") _
     Or (Mid (Trim (strString), 1, 1) = "") Then
    IsComment = True
  Else
    IsComment = False
  End If
End Function

'*********************************************************
' Purpose: Parses something like this:
'           "KEY(proc) value"
' Inputs:  strInput: input string.
' Outputs: strOutput: output string - would be "KEY value"
' Returns: the process name "proc"
'*********************************************************

Private Function ParseValPerProc (ByVal strInput, ByRef strOutput)
  Dim inPars
  Dim i
  Dim proc
  Dim ch

  inPars = False
  strOutput = ""
  proc = ""

  For i = 1 To Len (strInput)
    ch = Mid (strInput, i, 1)
    If (ch = "(") Then
      inPars = True
    ElseIf (ch = ")") Then
      inPars = False
    Else
      If (Not inPars) Then
        strOutput = strOutput & ch
      Else
        proc = proc & ch
      End If
    End If
  Next

  ParseValPerProc = proc
End Function

'*********************************************************
' Purpose: Extracts key and coresponding value from input string.
' Inputs:  strInpString: input string.
' Outputs: strKey: extracted key.
'          strValue: extracted value.
'          strProc: extracted process restriction (may be empty).
' Returns: If extraction is successful, return True.
'          Otherwise return False.
'*********************************************************

Private Function ParseEntry (ByRef strKey, ByRef strValue, ByRef strProc, _
                             ByRef strInpString)
  Dim strTempList, strTemp

  If (IsComment (strInpString)) Then
    ParseEntry = False
  Else
    strProc = ParseValPerProc (strInpString, strTemp)
    strTemp = Trim (strTemp)
    strTempList = Split (strTemp, " ", 2, 1)
    strKey = Trim (strTempList(LBound(strTempList)))

    If UBound(strTempList) <> LBound(strTempList) Then
      strValue = Trim (strTempList(UBound(strTempList)))
      ParseEntry = True
    Else
      strValue = ""
      ParseEntry = False
    End If
  End If

End Function

'*********************************************************
' Purpose: Test whether the current NS is the one requested, if not set it
'          and write it out
' Inputs:  newNS: requested NS
'          curNS: requested NS (in/out)
'          outdat: output file
'*********************************************************

Private Sub writeNS (ByRef newNS, ByRef curNS, ByRef outdat)
  if (curNS <> newNS) Then
    curNS = newNS
    outdat.WriteLine ("[" & curNS & "]")
  End If
End Sub

'*********************************************************
' Purpose: Write a converted opcinfo entry in both the XPL target format
'          and to the generated batch file
' Inputs:  ns:      namespace
'          key:     key
'          val:     value
'          outdat:  output file
'          batFile: batch file
'*********************************************************

Private Sub writeOpcInfoEntry (ByRef ns, ByRef key, ByRef val, ByRef outdat, _
                               ByRef batFile)
  outdat.WriteLine (key & " = " & val)
  batFile.WriteLine ("ovconfchg -ns " & ns & " -set " & key & " " & val)
End Sub

'*********************************************************
' Purpose: Get install directory for A.07.XX agent
' Inputs: -
' Outputs: strOvAgentDir : install directory for A.07.XX agent
' Returns: If directory is found return True.
'          Otherwise return False.
'*********************************************************
Private Function GetA07InstDir (ByRef strOvAgentDir)
  Dim MyShell, strAgtDir, strAgtDrive, strRegKey, basePath

  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")

  basePath = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\OpenView\ITO"
  strRegKey = basePath & "\Installation Drive"

  strAgtDrive = MyShell.RegRead (strRegKey)
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    GetA07InstDir = False
    Exit Function
  End If

  If (strAgtDrive = "") Then
    Set MyShell = Nothing
    GetA07InstDir = False
    Exit Function
  End If

  strRegKey = basePath & "\Installation Directory"

  strAgtDir = MyShell.RegRead (strRegKey)
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    GetA07InstDir = False
    Exit Function
  End If

  If (strAgtDir = "") Then
    Set MyShell = Nothing
    GetA07InstDir = False
    Exit Function
  End If

  strOvAgentDir = strAgtDrive & strAgtDir 
  Set MyShell = Nothing
  GetA07InstDir = True
End Function

'*******************************************************************************
 ' Purpose:Run a command and get the output assigned to a variant
 '
 ' Inputs: strcommandFile: Command line (do not include cmd /c)
 '	  opcagt -type
 '	  ovcoreid -show
 '        ovconfget sec.core CORE_ID
 ' Outputs:strSTDout
 '         The stdout ("\n" in the output is stripped off) returned by the command.
 '	  HTTPS
 '	  0e84cbe2-1df7-7517-14af-e4c0bc29fc75  
 ' Returns: True if command returned zero and gave some output, otherwise False
 '          If the command returned non-zero and gave output, it is stored in
 '          strSTDout anyway, and False is returned ("change user ..." returns 1 for success)
 '          
 '********************************************************************************
 
 Public Function get_command_stdout (ByRef strCommand, ByRef strSTDout)
 
   On Error Resume Next 
 
   Dim objFSO, objName, objTempFile, objTextFile, strText	
   Dim strCmd
   Dim RetVal, Result
 
   Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
   objName = objFSO.GetTempName
   objTempFile = GetTempDir() & "\" & objName	
 
   'Run the command
 
   strCmd = "cmd /c " + strCommand + " > """ & objTempFile & """"
   If(Not RunExternalScript(strCmd)) Then
     RetVal = False  				'##Could not execute 
     get_command_stdout = RetVal
     Exit Function
   End If
 
   Set objTextFile = objFSO.OpenTextFile(objTempFile, 1)
   If Err <> 0 Then
      Err.Clear
      RetVal = False				'##Could not open the result file
      get_command_stdout = RetVal
      Exit Function
   End If
   
   Do While objTextFile.AtEndOfStream <> True
      strText = strText & objTextFile.ReadLine	  	'## read a line for the output  
   Loop
 
   Retval = True
   objTextFile.Close
   objFSO.DeleteFile(objTempFile)	
 
   strSTDout = strText
 
   If(NOT Result) Then
     'Some windows commands return 1 when successful (!), so do this check as
     ' the last item, maybe we got the output correctly
     RetVal = False  				'##Could not execute or bogus windows error
     get_command_stdout = RetVal
     Exit Function
   End If

   get_command_stdout=RetVal
     
End Function

'*******************************************************************************

' Purpose: Back-Up old DCE based service discovery conf file to TMP dir as javaagent.conv
' Inputs: strSourceFile: Path of the DCE agent javaagent.cfg file.
'         strDestFile  : Path of the saved javaagent.cfg file. <TMP\javaagent.cfg>
' Returns: If back-up is successful, return True.
'          Otherwise return False.
'********************************************************************************

Public Function backup_svcdiscconfig (ByRef strSourceFile, ByRef strDestFile)

  Const ForReading = 1

  Dim fso, dat, strLine, MyShell, strKey, strValue, outdat, RetVal

  Dim strCmd, strSTDOut 

  Set fso = CreateObject ("Scripting.FileSystemObject")
  Set MyShell = CreateObject ("WScript.Shell")
  RetVal = True

  PrintMSG "backing up javaagent configuration file : "  & strSourceFile & _
             " to " & strDestFile , "Info"                      

  Set dat = fso.OpenTextFile(strSourceFile, ForReading, False)  
  
  If (Err.Number <> 0) Then

       Err.Clear
       dat.Close
       Set fso = Nothing
       Set outdat = Nothing
       Set MyShell = Nothing
       Set dat = Nothing
       RetVal = False
       backup_svcdiscconfig = RetVal
       PrintMSG "Service Disciovery agent configuration file not found: " & strSourceFile, "Warning"
       On Error Goto 0
       Exit Function
  End If

  Set outdat = fso.CreateTextFile(strDestFile, True)  

  If (Err.Number <> 0) Then

     Err.Clear
     outdat.Close
     Set fso = Nothing
     Set outdat = Nothing
     Set MyShell = Nothing
     Set dat = Nothing
     RetVal = False
     backup_svcdiscconfig = RetVal
     On Error Goto 0
     Exit Function
  End If
  
  Do While (dat.AtEndOfStream <> True)
    strLine = dat.ReadLine ()
    outdat.WriteLine (strLine)
  Loop  

backup_svcdiscconfig = RetVal

End Function

'*******************************************************************************
' Purpose: Check Terminal server mode, remember it, and set it to
'          INSTALL mode if that is not yet set
'
' Inputs: mode: True: set Install mode
'               False: set Execute mode
'         
' Returns: previous mode: True for Install (or no TS), False for Execute
'********************************************************************************

Public Function TSSetInstallMode(mode)

  Dim strCmd, strSTDOut

  'PrintMsg "Check Terminal Session mode.", "Info"

  strCmd = "change user /query"
  get_command_stdout strCmd, strSTDOut
  'PrintMSG "Got : '" & strSTDOut & "'", "Info"
  If strSTDOut <> "" Then
    If InStr(LCase(strSTDOut), "install") > 0 Then
    '"Application INSTALL mode is enabled. " or "Application EXECUTE mode is enabled.
    ' Install mode does not apply to a Terminal server configured for remote administration."
    'Note, this is not localized and will not work in languages that do not use the word "install"
    '  if you find a better way to check for TS, please replace this check here.

   	  PrintMSG "We have Terminal Server INSTALL mode (or compatible).", "Info"
      TSSetInstallMode = True
    Else
   	  PrintMSG "We have Terminal Server EXECUTE mode.", "Info"
      TSSetInstallMode = False
    End If

    If TSSetInstallMode <> mode Then
 	     'PrintMSG "Need to switch.", "Info"
      If mode Then
        strCmd = "change user /install"
   	    PrintMSG "Switching to INSTALL mode.", "Info"
      Else
        strCmd = "change user /execute"
   	    PrintMSG "Switching to EXECUTE mode.", "Info"
      End If

      get_command_stdout strCmd, strSTDOut
      'PrintMSG "Got : '" & strSTDOut & "'", "Info"
      If strSTDOut = "" Then
        PrintMsg "Switch failed.", "Warning"
      End If
    Else
   	   'PrintMSG "No switch needed.", "Info"
    End If

  Else
    TSSetInstallMode = True
    PrintMsg "No Terminal Services detected.", "Info"
  End If
End Function

'*******************************************************************************
' Purpose: Back-Up old nodeinfo file to TMP dir as nodeinfo.conv
'
' Inputs: strSourceFile: Path of the DCE agent nodeinfo file.
'         strDestFile  : Path of the saved nodeinfo file. <TMP\nodeinfo.conv>
'         
' Returns: If back-up is successful, return True.
'          Otherwise return False.
'********************************************************************************

Public Function backup_nodeinfo (ByRef strSourceFile, ByRef strDestFile)

  Const ForReading = 1
  Dim fso, dat, strLine, MyShell, strKey, strValue, outdat, RetVal
  Dim strCmd, strSTDOut
  
  Set fso = CreateObject ("Scripting.FileSystemObject")
  Set MyShell = CreateObject ("WScript.Shell")
  RetVal = True
  
  '## test if there is already a CoreID assigned
  '## this will be the case when a reporter(https) on DCE managed node 
  '## case when OVPA and DCE agent has been installed
    
  strCmd = "ovcoreid -show"
  If (get_command_stdout(strCmd, strSTDOut)) Then 	'## Return False if Failed
      If (strSTDOut <> " ") Then			'## Some thing I got 1.0 {GUID} 2.0 " "
     	 PrintMSG "The node already has CoreID assinged to : " & strSTDOut,"Info"
     	 PrintMSG "The agentID will not be migrated to CoreID.", "Info"
       	 bRetainAgentID = False      			'## say no to agentID migration
       	 RetVal = False					'## Not an error condition
       	 backup_nodeinfo = RetVal
       	 Exit Function
      Else
      	  RetVal = True					'## sec.core.CORE_ID varaible is set to ""
      End If         
  Else  							
      RetVal = True					'## Let the agentID be migrated to CoreID 
  End If
  
  PrintMSG "backing up nodeinfo file : "  & strSourceFile & _
             " to " & strDestFile , "Info"
  PrintMSG "Original nodeinfo file will be left untouched : " & strSourceFile, "Info"
             
  Set dat = fso.OpenTextFile(strSourceFile, ForReading, False)  
  
  If (Err.Number <> 0) Then
       Err.Clear
       dat.Close
       Set fso = Nothing
       Set outdat = Nothing
       Set MyShell = Nothing
       Set dat = Nothing
       RetVal = False
       backup_nodeinfo = RetVal
       PrintMSG "DCE agent nodeinfo file not found: " & strSourceFile, "Warning"
       On Error Goto 0
       Exit Function
  End If
  
  Set outdat = fso.CreateTextFile(strDestFile, True)
     
  If (Err.Number <> 0) Then
     Err.Clear
     outdat.Close
     Set fso = Nothing
     Set outdat = Nothing
     Set MyShell = Nothing
     Set dat = Nothing
     RetVal = False
     backup_nodeinfo = RetVal
     On Error Goto 0
     Exit Function
  End If
  
  Do While (dat.AtEndOfStream <> True)
    strLine = dat.ReadLine ()
    outdat.WriteLine (strLine)
  Loop  

backup_nodeinfo = RetVal

End Function

'*******************************************************************************
' Purpose: Convert old nodeinfo file to new format.(Only OPC_AGENT_ID converted
'          at this point in time)
' Inputs: strSourceFile: name of backed up nodeinfo.* file.
'         
' Returns: If conversion is successful, return True.
'          Otherwise return False.
'********************************************************************************

Public Function convert_nodeinfo (strSourceFile)

  Const ForReading = 1
  Dim fso, dat, strLine, MyShell, strKey, strValue, outdat
  Dim RetVal
  Dim strProc
  Dim i
  Dim agent_id
  Dim set_core_id
  
  Dim strInstallDir
  Dim ObjFSO
  Dim strCmd
  
  On Error Resume Next
  
  'initialize falgs
  agent_id = 0
  set_core_id = False
  RetVal = False
  
  Set fso = CreateObject ("Scripting.FileSystemObject")
  Set MyShell = CreateObject ("WScript.Shell")
  Set dat = fso.OpenTextFile(strSourceFile, ForReading, False)
  
  If (Err.Number <> 0) Then
     Err.Clear
     Set fso = Nothing
     Set MyShell = Nothing
     Set dat = Nothing           
     PrintMsg "DCE agent file :  " & strSourceFile & " not found.", "Error"
     RetVal = False	'nodeinfo file not found     
     Exit Function     
  End If 
  
  Do While (dat.AtEndOfStream <> True And Not set_core_id)   'Read untill you get OPC_AGENT_ID
      strLine = dat.ReadLine ()
      strProc = ""      
         
      If (ParseEntry (strKey, strValue, strProc, strLine)) Then
		  
         'handle AgentID only at this point in time
         If (strKey = "OPC_AGENT_ID") Then
   	    'ignore strproc
   	    set_core_id = True
   	    agent_id    = strValue   	       	    
   	 Else
	    set_core_id =  False	    
	 End If
	 
      End If
      
   Loop 
   
   strCmd = ""
   If (set_core_id) Then
   
   	strCmd  = "cmd /c ovconfchg -ns sec.core -set CORE_ID " & agent_id        	    
   	If (Not RunExternalScript (strCmd)) Then                             
   	    RetVal = False	'Setting AgentID failed
   	Else
   	    RetVal = True	'Setting AgentID succesfull
   	End If 
   Else
        RetVal = False  'OPC_AGENT_ID not found   
   End If
   
  convert_nodeinfo = RetVal  
     
End Function

'*********************************************************
' Purpose: Convert old opcinfo file to new format.
' Inputs: strSourceFile: name of source file.
'         strDestFile: name of destination file.
' Returns: If conversion is successful, return True.
'          Otherwise return False.
'*********************************************************

Public Function convert_to_ovo8 (ByRef strSourceFile, ByRef strDestFile, _
                                 ByRef batFileName, ByRef manager_name)
  Const ForReading = 1
  Dim fso, dat, strLine, MyShell, strKey, strValue, outdat
  Dim strObsoleteKeys
  Dim RetVal
  Dim add_manager
  Dim strProc
  Dim restrToProcsList
  Dim useRestrToProcsList
  Dim i
  Dim curNS
  Dim batFile

  On Error Resume Next
  RetVal = False
  'initialize indicator if we have MANAGER's name
  add_manager = 0
  strObsoleteKeys = Array ("OPC_COMM_PORT_RANGE", "OPC_RESTRICT_TO_PROCS", _
    "OPC_TRACE", "OPC_TRACE_AREA", "OPC_TRC_PROCS", "OPC_DBG_AREA", _
    "OPC_DBG_PROCS", "OPC_NODE_TYPE", "OPC_IP_ADDRESS", _
    "OPC_NSP_TYPE", "OPC_NSP_VERSION", "OPC_BUFLIMIT_SIZE", _
    "OPC_BUFLIMIT_SEVERITY", "OPC_AGENT_LOG_SIZE", "OPC_AGENT_LOG_DIR", _
    "OPC_HBP_INTERVAL_ON_AGENT", "OPC_COMM_TYPE", "OPC_NODE_CHARSET", _
    "OPC_MGMTSV_CHARSET", "OPC_AGTMSI_ENABLE", "OPC_AGTMSI_ALLOW_OA", _
    "OPC_AGTMSI_ALLOW_AA", "OPC_BUFLIMIT_ENABLE", "OPC_SG", _
    "OPC_SC", "OPC_VC", "OPC_INSTALLED_VERSION", "COMM_INSTALLED_VERSION", _
    "PERF_INSTALLED_VERSION", "OPC_MGMT_SERVER", "OPC_INSTALLATION_TIME")

  PrintMsg "Converting opcinfo file " & strSourceFile & _
           " to temporary file " & strDestFile, ""
         
  Set fso = CreateObject ("Scripting.FileSystemObject")
  Set MyShell = CreateObject ("WScript.Shell")
  Set dat = fso.OpenTextFile(strSourceFile, ForReading, False)
  If (Err.Number <> 0) Then
    Err.Clear
    Set fso = Nothing
    Set MyShell = Nothing
    Set dat = Nothing
    Convert = RetVal
    On Error Goto 0
    Exit Function
  End If
  Set outdat = fso.CreateTextFile(strDestFile, True)
  If (Err.Number <> 0) Then
    Err.Clear
    dat.Close
    Set fso = Nothing
    Set outdat = Nothing
    Set MyShell = Nothing
    Set dat = Nothing
    Convert = RetVal
    On Error Goto 0
    Exit Function
  End If

  Set batFile = fso.CreateTextFile(batFileName, True)
  If (Err.Number <> 0) Then
    Err.Clear
    dat.Close
    Set fso = Nothing
    Set outdat = Nothing
    Set MyShell = Nothing
    Set dat = Nothing
    Convert = RetVal
    On Error Goto 0
    Exit Function
  End If

  useRestrToProcsList = false
  curNS = ""

  Do While (dat.AtEndOfStream <> True)
    strLine = dat.ReadLine ()
    strProc = ""
    If (ParseEntry (strKey, strValue, strProc, strLine)) Then

      'handle MANAGER's name
      If (strKey = "OPC_MGMT_SERVER") Then
        add_manager = 1
        manager_name = strValue
        PrintMsg "Info: Recognized OVO Server name " & manager_name, ""

      ElseIf (strKey = "OPC_RESTRICT_TO_PROCS") Then
        ' -----------------------------------------------------------------
        ' If there is a RESTRICT_TO_PROCS, the value may be something like
        '  opcle, opcmona, opcmsgi
        ' Split these values into a list and make sub-NS from them below.
        ' -----------------------------------------------------------------
        if (instr(strValue, ",") = 0) then
          restrToprocsList = Array(Trim(strValue))
        else
          restrToProcsList = Split (strValue, ",")
          For i = 0 To UBound(restrToProcsList)
            restrToProcsList(i) = Trim (restrToProcsList(i))
          Next
        end if

        PrintMsg "Recognized process restriction " & strValue, ""
        useRestrToProcsList = true

      ElseIf (IsObsolete (strKey, strObsoleteKeys)) Then
        PrintMsg "Ignoring obsolete key " & strKey, "Warning"

      Else
        ' -------------------------------------------------------------------
        ' If the initial line was "KEY(proc) value" the 'proc' setting
        ' overrides an effective RESTRICT_TO_PROCS, so write it out as is
        ' -------------------------------------------------------------------
        If strProc <> "" Then
          PrintMsg "Set NS: eaagt; Direct process: " & strProc & _
                    "; Key: " & strKey & "; Value: " & strValue, ""

          writeNS "eaagt." & strProc, curNS, outdat
          writeOpcInfoEntry curNS, strKey, strValue, outdat, batFile
        Else
          ' -----------------------------------------------------------------
          ' If there is a RESTRICT_TO_PROCS, write a line for each member of
          ' the restrToProcsList which represents an applicable process ID
          ' -----------------------------------------------------------------
          If (useRestrToProcsList) Then
            For i = 0 To UBound(restrToProcsList)
              PrintMsg "Set NS: eaagt; Restricted process: " & _
                       restrToProcsList(i) & "; Key: " & strKey & _
                       "; Value: " & strValue, ""

              writeNS "eaagt." & restrToProcsList(i), curNS, outdat
              writeOpcInfoEntry curNS, strKey, strValue, outdat, batFile
            Next
          Else
            PrintMsg "Set NS: eaagt (global); Key: " & strKey & _
                     "; Value: " & strValue, ""
            writeNS "eaagt", curNS, outdat
            writeOpcInfoEntry curNS, strKey, strValue, outdat, batFile
          End If
        End If
      End If
    End If
  Loop

  'write MANAGER's name if we have one
  If (add_manager = 1) Then
    writeNS "sec.core.auth", curNS, outdat
    writeOpcInfoEntry curNS, "MANAGER", manager_name, outdat, batFile
    PrintMsg "Set MANAGER for NS: sec.core.auth to: " & manager_name, ""
  End If

  PrintMsg "Info: Conversion complete, closing output files ...", ""

  dat.Close
  outdat.Close
  batFile.Close

  If (Err.Number <> 0) Then
    Err.Clear
  Else
    RetVal = True
  End If

  Erase strObsoleteKeys
  Set fso = Nothing
  Set MyShell = Nothing
  Set dat = Nothing
  Set outdat = Nothing
  Set batFile = Nothing
  convert_to_ovo8 = RetVal

  On Error Goto 0
End Function

'*********************************************************
' Purpose: Installs package.
' Inputs: strPackage: name of package to install.
'         blnInteractive: is installation interactive.
'         blnForce: force reinstallation.
'         strInstDir: installation directory
' Returns: If installation is successful started, return True.
'          Otherwise return False.
'*********************************************************
Private Function InstallPackage (ByRef strPackage, ByVal blnInteractive, _
                                 ByVal blnForce, ByRef strInstDir, _
                                 ByRef strDataDir, ByVal doDebug, _
                                 ByRef rebootNeeded)
  Dim MyShell, Result
  Dim MyFSO, MySysEnv
  Dim strIntSwitch
  Dim strForceSwitch
  Dim strInstDirOption
  Dim strDataDirOption
  Dim logOption
  Dim strBreakDep

  On Error Resume Next  
  Set MyShell = CreateObject ("WScript.Shell")
  Set MyFSO = CreateObject ("Scripting.FileSystemObject")
  Set MySysEnv = MyShell.Environment ("USER")

  'process interactivity switch
  If (blnInteractive) Then
    strIntSwitch = "b"
  Else
    strIntSwitch = "n"
  End If
  'process force switch
  If (blnForce) Then
    'strForceSwitch = "/fa"
    strForceSwitch = "/i"
    strBreakDep = "TRUE"
  Else
    strForceSwitch = "/i"
    strBreakDep = "FALSE"
  End If

  If (doDebug) Then
    logOption = "/l*v """ & GetTempDir() & "\" & strPackage & ".log"""
  Else
    logOption = "/le+ """ & instLogPath & """"
  End If

  ' if installation directory is supplied
  ' include appropriate command line option
  If (strInstDir <> "") Then
    strInstDirOption = "INSTALLDIR=""" & strInstDir & """"
  Else
    strInstDirOption = ""
  End If

  ' if data directory is supplied
  ' include appropriate command line option
  If (strDataDir <> "") Then
    strDataDirOption = "DATADIR=""" & strDataDir & """"
  Else
    strDataDirOption = ""
  End If

  rebootNeeded = False

  Result = MyShell.Run ("%SYSTEMROOT%\system32\msiexec.exe " _
                      & strForceSwitch & " " & strPackage & ".msi " _
                      & "REBOOT=ReallySuppress " _
                      & "FORCEINSTALL=" & strBreakDep & " " _
                      & "/q" & strIntSwitch & " " _
                      & logOption & " " _
                      & strInstDirOption & " " _
                      & strDataDirOption _
                      , 1, True)  
  If (Err.Number <> 0) Then
    Err.Clear
    InstallPackage = False
    Set MySysEnv = Nothing
    Set MyShell = Nothing
    Set MyFSO = Nothing
    On Error Goto 0
    Exit Function
  End If

  If (Result = 3010) Then
    PrintMsg "Package " & strPackage & " requests system reboot.", ""
    'rebootNeeded = True
  End If

  InstallPackage = True
  Set MySysEnv = Nothing
  Set MyShell = Nothing
  Set MyFSO = Nothing
  On Error Goto 0
End Function

'*********************************************************
' Purpose: Checks if A.07.XX agent is installed.
' Inputs: -
' Returns: If agent is found, return True.
'          Otherwise return False.
'*********************************************************
Private Function ChkOldAgent ()
  Dim MyShell, MySysEnv, RetVal, Result, MyFSO, agtDir
  Dim strRegKey

  On Error Resume Next
  RetVal = False  
  Set MyShell = CreateObject ("WScript.Shell")
  Set MyFSO = CreateObject ("Scripting.FileSystemObject")
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\OpenView\ITO"

  ' we obtain agent directory from registry since this
  ' is more reliable method
  If ( Not GetA07InstDir (agtDir) ) Then
    agtDir = ""
  End If
  
  If (agtDir <> "") Then
    Dim opcinfoFile

    PrintMsg "Registry " & strRegKey & " yields directory " & agtDir, ""

    opcinfoFile = agtDir & "\bin\OpC\install\opcinfo"

    If (Not MyFSO.FileExists (opcinfoFile)) Then
      PrintMsg "Registry keys in " & strRegKey & " set but file " & _
               opcinfoFile & " does not exist.", "Warning"
    Else
      Result = MyShell.Run ("find ""OPC_INSTALLED_VERSION A.07"" """ & _
                            opcinfoFile & """", 0, True)
      If (Err.Number <> 0) Then
        PrintMsg "Error checking OVO version in opcinfo file.", "Warning"
        Err.Clear
      Else
        If (Result = 0) Then
          PrintMsg "Found version string of A.07.x in opcinfo file.", ""
          RetVal = True
        Else
          PrintMsg "opcinfo file " & opcinfoFile & " exists but does not " & _
                   "contain OPC_INSTALLED_VERSION or version < 7.x" & _
                   Chr(10) & "Ignoring. Clean up manually.", "Warning"
        End If
      End If
    End If

  Else
    PrintMsg "Registry keys in " & strRegKey & " not set.", ""
  End If

  ChkOldAgent = RetVal
  Set MyShell = Nothing
  Set MySysEnv = Nothing
  Set MyFSO = Nothing
  On Error Goto 0
End Function

'*********************************************************
' Purpose: Deinstalls old agent.
' Inputs: -
' Returns: If deinstallation succeeded, return True.
'          Otherwise return False.
'*********************************************************
Private Function DeinstallOldAgent ()
  Dim MyShell, RetVal, Result, setupU, errNo, objFSO
  Dim strAgtDir
  Dim res 
  Dim sDeInstallString

  On Error Resume Next
  Result = False
  Set MyShell  = CreateObject ("WScript.Shell")
  Set objFSO   = CreateObject ("Scripting.FileSystemObject")
  currDir      = MyShell.CurrentDirectory

  If ( Not GetA07InstDir (strAgtDir) ) Then
    PrintMsg "Could not determine installation directory for OVO 7 agent.", _
             "Error"
    Set objFSO = Nothing
    Set MyShell = Nothing
    DeinstallOldAgent = False
    Exit Function
  End If

  PrintMsg "Deinstalling OVO 7 agent from " & strAgtDir & _
           " ...", ""
  'Fix QXCR1000430963
  'check-if de-installer exists  
  If (objFSO.FileExists ( dce_deinstall.exe )) Then   
     Result = MyShell.Run ("dce_deinstall.exe /DoRemoval /l dce_deinstall.log""", 1, True)   

     errNo  = Err.Number  
  Else   
   PrintMsg "DCE agent de-installer not found, assuming OVO/U based DCE agent ...", ""         
   'This is the OVO/UX version ...
    setupU = strAgtDir & "\bin\OpC\opcsetup.exe"
    'Fix QXCM1000340143
    opcPath   = strAgtDir & "\bin\OpC"
    opcsetupU = "opcsetup.exe -u"
    If (objFSO.FileExists (setupU)) Then
       PrintMsg "OVO 7.x agent belongs to OVO/Unix server ...", ""      
       Result = MyShell.Run ("%ComSpec% /C ""cd /d " & strAgtDir & "\bin\OpC && opcsetup.exe -u""", 0, True)
       errNo  = Err.Number
    End If
  End If

  'If (errNo <> 0 ) Then

   ' PrintMsg "De-installation failed.", "Error"

    'Err.Clear

    'DeinstallOldAgent = False

    'Set MyShell = Nothing

    'On Error Goto 0

    'Exit Function

  'End If

  If (Result = 0) Then
    RetVal = True
  Else
    PrintMsg "De-installation failed." & _
             " Result: " & Result , "Error"
  End If

  DeinstallOldAgent = RetVal
  Set MyShell = Nothing
  Set objFSO = Nothing
End Function

'*********************************************************
' Purpose: Deinstalls package.
' Inputs: strPackage: name of package.
'         dctGUIDs: dictionary of package update GUIDs.
'         blnInteractive: is deinstallation interactive.
'         doDebug:
'         rebootNeeded:
'         breakDep: break dependencies.
' Returns: If package is deinstalled, return True.
'          Otherwise return False.
'*********************************************************
Private Function DeinstallPackage (ByRef strPackage, ByRef dctGUIDs, _
                                   ByVal blnInteractive, ByVal doDebug, _
                                   ByRef rebootNeeded, _
                                   ByVal breakDep)
  Dim strRegKey, MyShell, RetVal, strKeyValue
  Dim Result
  Dim strIntSwitch
  Dim logOption
  Dim strForceOpt
 
  RetVal = False
  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")

  'Get ProductCode
  If (dctGUIDs.Exists(strPackage)) Then
    strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView"
    strRegKey = strRegKey & "\" & dctGUIDs.Item(strPackage) & "\ProductCode"
    strKeyValue = MyShell.RegRead(strRegKey)
    If (Err.Number <> 0) Then
      Err.Clear
      Set MyShell = Nothing
      DeinstallPackage = False
      On Error Goto 0
      Exit Function
    End If

    strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView"
    strRegKey = strRegKey & "\" & dctGUIDs.Item(strPackage) & "\Depends"
    strKeyValue = MyShell.RegRead(strRegKey)
    If (Err.Number = 0) Then
      If (Len(Trim(strRegValue)) = 0) Then
        Err.Clear
        Set MyShell = Nothing
        DeinstallPackage = False
        On Error Goto 0
        Exit Function
      End If
    End If

    PrintMsg "Completed the dependency Check" , "Info"

    'Start deinstallation of package
    If (blnInteractive) Then
      strIntSwitch = "b"
    Else
      strIntSwitch = "n"
    End If

    If (doDebug) Then
      logOption = "/l*v """ & GetTempDir() & "\" & strPackage & ".log"""
    Else
      logOption = "/le+ """ & instLogPath & """"
    End If

    rebootNeeded = False

    If (breakDep) Then
      strForceOpt = "TRUE"
    Else
      strForceOpt = "FALSE"
    End If

    Result = MyShell.Run ("%SYSTEMROOT%\system32\msiexec.exe /x " _ 
             & strKeyValue & " " _ 
             & "REBOOT=ReallySuppress " _
             & "ZBIGNORECONDITIONS=" & strForceOpt & " " _
             & " /q" & strIntSwitch & " " _
             & logOption, 1, True)
    If (Err.Number <> 0) Then
      Err.Clear
      Set MyShell = Nothing
      DeinstallPackage = False
      On Error Goto 0
      Exit Function
    End If

    If (Result = 0) Then
      RetVal = True
    ElseIf (Result = 3010) Then
      PrintMsg "Package " & strPackage & " requests system reboot.", ""
      RetVal = True
      'rebootNeeded = True
    End If
  End If

  DeinstallPackage = RetVal
  Set MyShell = Nothing
  On Error Goto 0
End Function

'*********************************************************
' Purpose: Checks if package strPackage is installed.
' Inputs: strPackage: name of package.
'         dctGUIDs: dictionary of package update GUIDs.          
' Returns: If package is installed, return True.
'          Otherwise return False.
'*********************************************************
Private Function IsPackageInstalled (ByRef strPackage, ByRef dctGUIDs)
  Dim strRegKey, MyShell, RetVal, strKeyValue
  
  RetVal = False
  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView"
  If (dctGUIDs.Exists(strPackage)) Then
    strRegKey = strRegKey & "\" & dctGUIDs.Item(strPackage) & "\PackageName"
    strKeyValue = MyShell.RegRead(strRegKey)
    If (Err.Number <> 0) Then
      Err.Clear
      Set MyShell = Nothing
      IsPackageInstalled = False
      On Error Goto 0
      Exit Function
    End If
    If ( InStr(1,strKeyValue,strPackage,1) > 0 ) Then
      RetVal = True
    End If
  End If
  IsPackageInstalled = RetVal
  Set MyShell = Nothing
  On Error Goto 0
End Function

'**************************************************************
' Purpose: Creates directory path with all intermediate 
'          directories.
' Inputs: strBaseDir: base directory from which directory path
'                     is created.
'         aryDirPath: aray of directories to be created.
' Returns: If directory path is successfully created, 
'          return True.
'          Otherwise return False.
'**************************************************************
Private Function CreateDirectoryPath (ByVal strBaseDir, ByRef aryDirPath)
  Dim objFSO, objFolder, Folders
  Dim strPath, i
  Dim RetVal

  On Error Resume Next
  RetVal = False
  Set objFSO = CreateObject ("Scripting.FileSystemObject")
  strPath = strBaseDir
  For i = 0 To UBound (aryDirPath)
    Set objFolder = objFSO.GetFolder (strPath)
    If (Err.Number <> 0) Then
      Err.Clear
      CreateDirectoryPath = False
      Set objFSO = Nothing
      Set objFolder = Nothing
      On Error Goto 0
      Exit Function
    End If
    strPath = objFSO.BuildPath (strPath, aryDirPath(i))
    If (Not objFSO.FolderExists (strPath)) Then
      Set Folders = objFolder.SubFolders
      Folders.Add (aryDirPath(i))
    End If
    RetVal = True
  Next
  CreateDirectoryPath = RetVal
  Set objFSO = Nothing
  Set objFolder = Nothing
  Set Folders = Nothing
  On Error Goto 0
End Function

'**************************************************************
' Purpose: Copy converted opcinfo file to appropriate location. 
' Inputs: -
' Returns: If file is successfully copied, return True.
'          Otherwise return False.
'**************************************************************
Private Function CopyConverted ()
  Dim MyShell, strRegKey, strDataDir, RetVal, Result
  Dim strTempDir, objFSO, aryDir(3)

  On Error Resume Next
  RetVal = False
  Set MyShell = CreateObject ("WScript.Shell")
  Set objFSO = CreateObject ("Scripting.FileSystemObject")
  'Get <DataDir> value from registry.
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\DataDir"
  strDataDir = MyShell.RegRead(strRegKey)
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    Set objFSO = Nothing
    CopyConverted = False
    On Error Goto 0
    Exit Function
  End If
  'Check and create target directory structure if necessary.
  aryDir(0) = "conf"
  aryDir(1) = "confpar"
  aryDir(2) = "default"
  If (Not CreateDirectoryPath (strDataDir, aryDir) ) Then
    Err.Clear
    Set MyShell = Nothing
    Set objFSO = Nothing
    CopyConverted = False
    On Error Goto 0
    Exit Function
  End If
  'Copy opcinfo.converted to target directory.
  objFSO.CopyFile GetTempDir() & "\" & opcinfoConvFile, _
         strDataDir & "conf\confpar\default\ovo_old.ini", True
  If (Err.Number <> 0 ) Then
    Err.Clear
    Set MyShell = Nothing
    Set objFSO = Nothing
    CopyConverted = False
    On Error Goto 0
    Exit Function
  End If
  If (Result = 0) Then
    RetVal = True
  End If

  CopyConverted = RetVal
  Set MyShell = Nothing
  Set objFSO = Nothing
  On Error Goto 0
End Function

'**************************************************************
' Purpose: Run external script 
' Inputs: strScriptName - name of script to be executed.
' Outputs: -
' Returns: If execution is successfull, return True.
'          Otherwise return False.
'**************************************************************
Private Function RunExternalScript (strScriptName)
  Dim MyShell, RetVal, Result

  PrintMsg "Executing " & strScriptName & " ...", ""

  'Establish error-handling
  On Error Resume Next
  RetVal = False
  Set MyShell = CreateObject ("WScript.Shell")

  Result = MyShell.Run ( strScriptName, 0, True)
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    RunExternalScript = False
    On Error Goto 0
    Exit Function
  else
    RetVal = True
  End If

  RunExternalScript = RetVal
  Set MyShell = Nothing
  On Error Goto 0
End Function

'******************************************************************************
' Purpose: Get InstallDir from Windows Registry.
' Inputs: -
' Outputs: strInstallDir: InstallDir directory path.
' Returns: If everything is ok, return True.
'          Otherwise return False.
'******************************************************************************
Private Function GetInstallDir (ByRef strInstallDir)
  Dim MyShell
  Dim strRegKey

  'Establish error handling
  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  'Registry key for InstallDir
  strRegKey = _
    "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\InstallDir"
  'read Registry
  strInstallDir = MyShell.RegRead (strRegKey)
  if (Err.Number <> 0) Then
    Err.Clear
    GetInstallDir = False
    Exit Function
  End If
  GetInstallDir = True
  Set MyShell = Nothing
End Function

'******************************************************************************
' Purpose: Get list of installed packages
' Inputs:  dctGUIDs: dictionary of package update GUIDs.
' Outputs: strList: list of installed packages.
' Returns: If at least one package is installed, return True.
'          Otherwise return False.
'******************************************************************************
Private Function GetInstalledPackages (ByRef dctGUIDs, ByRef strList, _
                                       ByRef missing)
  Dim MyShell
  Dim pkgs
  Dim i
  Dim blnAnythingInstalled

  'Establish error handling
  On Error Resume Next
  strList = ""
  missing = ""
  blnAnythingInstalled = False

  pkgs = dctGUIDs.Keys
  For i = 0 To dctGUIDs.Count - 1
    If (IsPackageInstalled ( pkgs(i), dctGUIDs )) Then
      strList = strList & "  " & pkgs(i) & Chr(10)
      blnAnyThingInstalled = True
    Else
      missing = missing & "  " & pkgs(i) & Chr(10)
    End If
  Next
  GetInstalledPackages = blnAnythingInstalled
End Function

'*********************************************************
' Purpose: To clean the registry entry in case the entry 
'           it points to dont exist.
' Inputs: strRegKey: The registry key .
'         
' Returns: If conversion is successful, return True.
'          Otherwise return False.
'*********************************************************

Public Function CleanRegistry(strRegKey)

       On Error Resume Next
 
       Dim WshShell,fileSystemObj
       Set WshShell = CreateObject ("WScript.Shell")
       Set fileSystemObj=CreateObject("Scripting.fileSystemObject")
 
        'get Datadir from Registry
        strDataDir = WshShell.RegRead(strRegKey)

          If (Err.Number <> 0) Then
    		Err.Clear
	        Set MyShell = Nothing
       		Exit Function
          End If

                Wscript.echo "The datadir is " & strDataDir

          If (fileSystemObj.FolderExists(strDataDir)=True) Then
                 Wscript.echo "DataDir  Exists"
             
          Else 
                 Wscript.echo "DataDir do not   Exist,hence deleting the key"
                 WshShell.regdelete(strRegKey)
          End If
        
End Function

'*********************************************************
' Purpose: To clean the registry entry 
' Inputs: strRegKey: The registry key .
'         
' Returns: True.
'*********************************************************

Public Function deleteregistry(strRegKey)

       On Error Resume Next
 
       Dim WshShell
       Set WshShell = CreateObject ("WScript.Shell")
 
        'get Datadir from Registry
        WshShell.regdelete(strRegKey)

          If (Err.Number <> 0) Then
    		Err.Clear
	        Set MyShell = Nothing
       		Exit Function
          End If

          Wscript.echo "Deleted registry entry " & strRegKey
        
End Function

'******************************************************************************
' Purpose: Print opc_inst.vbs usage information. 
' Inputs: blnInteractive : interactivity flag 
' Outputs: -
' Returns: -
'******************************************************************************
Private Sub PrintUsageInfo (byVal blnInteractive)
  Dim strLine1, strLine2, strLine3, strLine4, strLine5
  Dim strLine6, strLine7, strLine8, strLine9, strLine10
  Dim strLine11, strLine12, strLine13, strLine14,strLine15

  strLine1 = "opc_inst.vbs [ -help|-h ]"
  strLine2 = "             [ -non_int|-ni ]"
  strLine3 = "             ( (-remove|-r) |"
  strLine4 = "               (-verify|-v) ) | "
  strLine5 = "             [ -force|-f ] [ -no_start|-ns]"
  strLine6 = "             [ -configure|-c <config file> ]"
  strLine7 = "             [ -srv|-s <management server> "
  strLine8 = "               [ -cert_srv <certificate server> ] ]"
  strLine9 = "             [ -wscript | -w ]"
  strLine10 = "             [ -inst_dir | -id <install dir> ]"
  strLine11 = "             [ -no_boot | -nb ]"
  strLine12 = "             [ -break_dep ]"
  strLine13 = "             [PackageName]"
  strLine14 = "             [-no_instnotify]"
  strLine15 = "             [ -data_dir | -dd <data dir> ]"

  If (blnInteractive) Then
    WScript.Echo strLine1 + Chr(10) + strLine2 + Chr(10)_
               + strLine3 + Chr(10) + strLine4 + Chr(10)_
               + strLine5 + Chr(10) + strLine6 + Chr(10)_
               + strLine7 + Chr(10) + strLine8 + Chr(10) _
               + strLine9 + Chr(10) + _
               + strLine10 + Chr(10) + _
               + strLine11 + Chr(10) + _
               + strLine12 + Chr(10) + _
               + strLine13 + Chr(10) + _
			   + strLine14 + Chr(10) + _
			   + strLine15 + Chr(10) + _			   
      "-help      Display this message." + Chr(10) + _
      "-remove    Remove installed agent packages." + Chr(10) + _
      "-verify    Check for installed agent packages." + Chr(10) + _
      "-force     Force installation if same version is already installed." + Chr(10) + _
      "-break_dep Deinstall package even if other products depend on it." + Chr(10) + _
      "-non_int   Non-interactive mode. No GUI displayed, no log output." + _
                  Chr(10) + _
      "           (default: installer windows and log output appear)." + _
                  Chr(10) + _
      "-no_start  Avoid start of configuration process." + Chr(10) + _
      "-configure Configure product using agent profile." + Chr(10) + _
      "-srv       Set management server" + Chr(10) + _
      "-cert_srv  Set certificate server" + Chr(10) + _
      "-wscript   If started within wscript, run w/o confirmation." + Chr(10) + _
      "           Recommended is to run within cscript." + Chr(10) + _
      "-inst_dir  Specify install directory" + Chr(10) + _
      "           (default: %ProgramFiles%\HP OpenView)" + Chr(10) + _
      "-data_dir  Specify data directory" + Chr(10) + _
      "           Default Data Dir path is OS specific"  + Chr(10) + _
      "           Dafault Windows 2003 Data Dir Path: %ALLUSERSPROFILE%\Application Data\HP\HP BTO Software" + Chr(10) + _
      "           Dafault Windows 2008 Data Dir Path: %ALLUSERSPROFILE%\Hp\HP BTO Software" + Chr(10) + _  
      "-no_boot   Create service in manual start mode" + Chr(10) + _
      "           (default: auto)." + Chr(10) + _
      "-no_instnotify	Avoid creating a installation notification file." + Chr(10) + _
      "Without any options the OVO agent installation will be started."
  End If
End Sub

'**************************************************************
' Purpose: Get ovc command or error if not present
' Inputs:  -
' Outputs: ovcCmd path
' Returns: True if the command exists, False if not
'**************************************************************

Private Function GetOvcCommand (ByRef ovcCmd)
  Dim strInstallDir, objFSO
  Dim RetVal

  'Establish error-handling
  On Error Resume Next

  RetVal = False

  If (GetInstallDir (strInstallDir)) Then
    Set objFSO = CreateObject ("Scripting.FileSystemObject")
    ovcCmd = strInstallDir & "bin\ovc.exe"
    RetVal = objFSO.FileExists(ovcCmd)

    Set objFSO = Nothing
  End If

  if (Not RetVal) Then
    PrintMsg "OvCtrl command ovc.exe not found.", ""
  End If

  GetOvcCommand = RetVal
End Function

'**************************************************************
' Purpose: Stop L-Core
' Inputs: -
' Returns: If stopping is successfull, return True.
'          Otherwise return False.
'**************************************************************

Private Function StopLCore ()
  Dim MyShell, RetVal, Result
  Dim ovcCmd

  'Establish error-handling
  On Error Resume Next

  If (Not GetOvcCmd (ovcCmd)) Then
    StopLCore = False
    Exit Function
  End If
  
  PrintMsg "Stopping OVO agent ...", ""
  RetVal = False

  Result = MyShell.Run ("""" & ovcCmd & """ -stop", 0, True)
  If (Err.Number <> 0) Then
    Err.Clear
    PrintMsg "Failed to stop OVO agent.", "Error"
    Set MyShell = Nothing
    StopLCore = False
    Exit Function
  End If
  If (Result = 0) Then
    RetVal = True
    PrintMsg "Stopped OVO agent.", ""
  End If

  StopLCore = RetVal
  Set MyShell = Nothing
End Function

'**************************************************************
' Purpose: Start L-Core
' Inputs: -
' Returns: If starting is successfull, return True.
'          Otherwise return False.
'**************************************************************
Private Function StartLCore ()
  Dim MyShell, RetVal, Result
  Dim strRegKey, strInstallDir

  'Establish error-handling
  On Error Resume Next
  RetVal = False
  Set MyShell = CreateObject ("WScript.Shell")
  'Get <InstallDir> from registry
  strRegKey = _
    "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\InstallDir"
  strInstallDir = MyShell.RegRead(strRegKey)
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    StartLCore = False
    Exit Function
  End If

  Result = MyShell.Run ("""" & strInstallDir & "bin\ovc.exe"" -start", _
			0, True)

  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    StartLCore = False
    Exit Function
  End If
  If (Result = 0) Then
   RetVal = True
  End If

  StartLCore = RetVal
  Set MyShell = Nothing
End Function
'**************************************************************
' Purpose: Kill L-Core
' Inputs: -
' Returns: If termination is successfull, return True.
'          Otherwise return False.
'**************************************************************

Private Function KillLCore ()
  Dim MyShell, RetVal, Result
  Dim ovcCmd

  'Establish error-handling
  On Error Resume Next

  If (Not GetOvcCmd (ovcCmd)) Then
    KillLCore = False
    Exit Function
  End If

  PrintMsg "Killing OV services ...", ""
  RetVal = False
 
  Result = MyShell.Run ("""" & ovcCmd & """ -kill", 0, True)
  If (Err.Number <> 0) Then
    Err.Clear
    PrintMsg "Failed to kill OV services.", "Error"
    Set MyShell = Nothing
    KillLCore = False
    Exit Function
  End If
  If (Result = 0) Then
    PrintMsg "Killed OV services.", ""
    RetVal = True
  End If

  KillLCore = RetVal
  Set MyShell = Nothing
End Function

'*********************************************************
' Purpose: Gets TMP directory
' Returns: Temp directory
'*********************************************************
Private Function GetTempDir()
  Dim MyShell
  Dim MySysEnv, MyProcEnv, MyUserEnv
  Dim sysRoot

  Set MyShell   = CreateObject ("WScript.Shell")
  Set MySysEnv  = MyShell.Environment("SYSTEM")
  Set MyProcEnv = MyShell.Environment("PROCESS")
  Set MyUserEnv = MyShell.Environment("USER")

  sysRoot = MySysEnv("SystemRoot")
  If (Err.Number <> 0) Then
    Err.Clear
    sysRoot = ""
  End If

  ' For some reason sometimes the SYSTEM env does not yield SystemRoot ... 
  If (sysRoot = "") Then
    sysRoot = MyProcEnv("SystemRoot")
    If (Err.Number <> 0) Then
      Err.Clear
      sysRoot = ""
    End If
  End If

  If (sysRoot = "") Then
    sysRoot = MyUserEnv("SystemRoot")
    If (Err.Number <> 0) Then
      Err.Clear
      sysRoot = ""
    End If
  End If

  If (sysRoot = "") Then
    GetTempDir = "\"
  Else
    GetTempDir = sysRoot & "\Temp"
  End If
  
  Set MyUserEnv = Nothing
  Set MyProcEnv = Nothing
  Set MySysEnv  = Nothing
  Set MyShell   = Nothing
End Function

'******************************************************************************
' Purpose: Get log file name
' Returns: The name of the log file to be used
'******************************************************************************
Private Function GetLogFileName(ByRef fName)
  GetLogFileName = GetTempDir() & "\" & fName
End Function

'******************************************************************************
' Purpose: Write text to logfile.
' Inputs: strText - text to be written to logfile
' Outputs: -
' Returns: If everything is ok, return True.
'          Otherwise return False.
'******************************************************************************
Private Function WriteToLog (ByRef strText)
  Dim fso, dat
  Const ForAppending = 8

  ' we establish error handling
  On Error Resume Next

  Set fso = CreateObject ("Scripting.FileSystemObject")
  ' logfile is opended for appending and is created if one does
  ' not exist
  Set dat = fso.OpenTextFile (instLogPath, ForAppending, True, True)
  If (Err.Number <> 0) Then
    Err.Clear
    WriteToLog = False
    Set dat = Nothing
    Set fso = Nothing
    Exit Function
  End If

  dat.WriteLine strText
  If (Err.Number <> 0) Then
    Err.Clear
    WriteToLog = False
    dat.Close
    Set dat = Nothing
    Set fso = Nothing
    Exit Function
  End If

  dat.Close
  WriteToLog = True
  Set dat = Nothing
  Set fso = Nothing
End Function

'**************************************************************
' Purpose: Copy opc_inst.log to <DataDir>\log to appropriate location.
' Inputs: -
' Returns: If file is successfully copied, return True.
'          Otherwise return False.
'**************************************************************
Private Function CopyLogFile ()
  Dim MyShell, strRegKey, strDataDir
  Dim objFSO
  Dim targetFile

  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  Set objFSO = CreateObject ("Scripting.FileSystemObject")

  CopyLogFile = True

  'Get <DataDir> value from registry.
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\DataDir"
  strDataDir = MyShell.RegRead(strRegKey)
  If (Err.Number <> 0) Then
    Err.Clear
    CopyLogFile = False
  Else
    targetFile = strDataDir & "log\" & instLogName
    PrintMsg "Moving install log file to " & targetFile, ""

    ' Rename existing log file
    If (objFSO.FileExists(targetFile)) Then
      Dim oldFile
      Dim oldFileDate
      Set oldFile = objFSO.GetFile(targetFile)
      oldFileDate = Replace(oldFile.DateLastModified, "/", "-")
      oldFileDate = Replace(oldFileDate, "\", "-")
      oldFileDate = Replace(oldFileDate, " ", "-")
      oldFileDate = Replace(oldFileDate, ":", "-")
      objFSO.CopyFile targetFile, targetFile & "-" & oldFileDate, True
      Set oldFile = Nothing
    End If

    'Copy opc_inst.log to target <DataDir>\log
    objFSO.CopyFile instLogPath, targetFile, True
    If (Err.Number <> 0 ) Then
      PrintMsg "Could not move installation log file to " & targetFile & _
               ". See " & instLogPath, "Error"
      Err.Clear
      CopyLogFile = False
    Else
      objFSO.DeleteFile instLogPath, True
    End If
  End If

  Set MyShell = Nothing
  Set objFSO = Nothing
End Function 

'******************************************************************************
' Purpose: Get DataDir from Windows Registry.
' Inputs: -
' Outputs: strDataDir: DataDir directory path.
' Returns: If everything is ok, return True.
'          Otherwise return False.
'******************************************************************************
Private Function GetDataDir (ByRef strDataDir)
  Dim MyShell
  Dim strRegKey

  'Establish error handling
  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  'Registry key for DataDir
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\DataDir"
  'read Registry
  strDataDir = MyShell.RegRead (strRegKey)
  if (Err.Number <> 0) Then
    Err.Clear
    GetDataDir = False
    Exit Function
  End If
  GetDataDir = True
  Set MyShell = Nothing
End Function

'**************************************************************
' Purpose: Remove folders which have to be removed after
'          deinstalling HPOvXpl
' Inputs: -
' Returns: If file is successfully copied, return True.
'          Otherwise return False.
'**************************************************************
Private Sub XplCleanup (ByRef strDataDir, ByRef sysEnv)
Dim objFSO

  On Error Resume Next

  ' Do this only if XPL is not installed anymore
  If (IsPackageInstalled (XplName, dctPackageGUID) ) Then
    Exit Sub
  End If

  Set objFSO = CreateObject ("Scripting.FileSystemObject")
  objFSO.DeleteFolder strDataDir & "conf\xpl\config", True
  If (Err.Number <> 0) Then
    Err.Clear
  End If
  objFSO.DeleteFolder strDataDir & "datafiles\xpl\config", True
  If (Err.Number <> 0) Then
    Err.Clear
  End If
  objFSO.DeleteFolder strDataDir & "datafiles\sec\ks", True
  If (Err.Number <> 0) Then
    Err.Clear
  End If

  ' Clear these variables. Actually this is the job of XPL - so remove this
  ' code here when XPL does it.
  sysEnv.Remove "OvInstallDir"
  sysEnv.Remove "OvDataDir"

  Set objFSO = Nothing
End Sub

'**************************************************************
' Purpose: Kills A.07 agent's processes
' Inputs: -
' Returns: If no error occurs return True.
'          Otherwise return False.
'**************************************************************
Private Function KillA07Agent ()
  Dim strAgtDir, MyShell, Result, strAgtKill

  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")

  If ( GetA07InstDir (strAgtDir) ) Then
    strAgtKill = strAgtDir & "\bin\OpC\opcagt" & chr(34) & " -kill" & chr(34)

    Result = MyShell.Run ( chr(34) & strAgtKill, _
                           0, True)
    If (Err.Number <> 0) Then
      Err.Clear
      KillA07Agent = False
      Set MyShell = Nothing
      Exit Function
    End If

    If ( Result = 0 ) Then
      KillA07Agent = True
      Set MyShell = Nothing
      Exit Function
    End If
  End If

  KillA07Agent = False
  Set MyShell = Nothing
End Function

'**************************************************************
' Purpose: Saves ECS, coda, stored facts data and nodeinfo file
' Inputs: -
' Returns: If data is found and saved return True.
'          Otherwise return False.
'**************************************************************
Private Function BackupData ()
  Dim MyShell, objFSO, strRegKey
  Dim strTMP, strOVDIR
  Dim strSource, strDest
  
  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  Set objFSO = CreateObject ("Scripting.FileSystemObject")

  ' determine <TMP>
  strTMP = GetTempDir()

  ' determine <OV_DIR>
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\OpenView\ITO\Installation Directory"
  strOVDIR = MyShell.RegRead(strRegKey)
  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell = Nothing
    Set objFSO = Nothing
    BackupData = False
    Exit Function
  End If

  If ( Not GetA07InstDir (strOVDIR) ) Then
    PrintMsg "Could not determine installation directory for OVO 7 agent.", _
             "Error"
    BackupData = False
    Set objFSO = Nothing
    Set MyShell = Nothing
    Exit Function
  End If

  ' create temporary directory
  ' <TMP>\ecs_fact
  If (Not objFSO.FolderExists (strTMP)) Then
    objFSO.CreateFolder strTMP
    If (Err.Number <> 0) Then
      Err.Clear
      Set objFSO = Nothing
      Set MyShell = Nothing
      BackupData = False
      Exit Function
    End If
  End If
  If (Not objFSO.FolderExists (strTMP + "\ecs_fact")) Then
    objFSO.CreateFolder strTMP + "\ecs_fact"
    If (Err.Number <> 0) Then
      Err.Clear
      Set objFSO = Nothing
      Set MyShell = Nothing
      BackupData = False
      Exit Function
    End If
  End If

  ' create temporary directory
  ' <TMP>\coda
  If (Not objFSO.FolderExists (strTMP + "\coda")) Then
    objFSO.CreateFolder strTMP + "\coda"
    If (Err.Number <> 0) Then
      Err.Clear
      Set objFSO = Nothing
      Set MyShell = Nothing
      BackupData = False
      Exit Function
    End If
  End If
  
  ' remove any previously saved data
  ' <TMP>\ecs_fact\*.ds, <TMP>\ecs_fact\*.fs
  objFSO.DeleteFile strTMP + "\ecs_fact\*.ds", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot remove previously saved *.fs files " & Err.Description
    Err.Clear
  End If
  objFSO.DeleteFile strTMP + "\ecs_fact\*.fs", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot remove previously saved *.ds files " & Err.Description
    Err.Clear
  End If
  objFSO.DeleteFile strTMP + "\coda\*.*", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot remove previously saved coda files " & Err.Description
    Err.Clear
  End If
  objFSO.DeleteFile strTMP + "\nodeinfo.conv", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot remove previously saved nodeinfo.conv file " & Err.Description
    Err.Clear
  End If

  ' because some files will be moved
  ' we have to kill A.07 agent
  If ( Not KillA07Agent () ) Then
    PrintMsg "Could not stop OVO 7 agent processes.", "Error"
  End If
  
  ' copy *.ds data
  ' <OV_DIR>\conf\OpC\*.ds ->  <TMP>\ecs_fact
  objFSO.CopyFile strOVDIR + "\conf\OpC\*.ds", strTMP + "\ecs_fact\", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot backup *.ds files " & Err.Description
    Err.Clear
  End If
  ' copy *.fs data
  ' <OV_DIR>\conf\OpC\*.fs ->  <TMP>\ecs_fact
  objFSO.CopyFile strOVDIR + "\conf\OpC\*.fs", strTMP + "\ecs_fact\", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot backup *.fs files " & Err.Description
    Err.Clear
  End If
  ' move coda data
  ' <OV_DIR>\conf\Bbc\default.txt -> <TMP>\coda
  objFSO.CopyFile strOVDIR + "\conf\BBC\default.txt", strTMP + "\coda\", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot backup default.txt file in BBC folder " & Err.Description
    Err.Clear
  End If
  ' <OV_DIR>\databases\coda* -> <TMP>\coda
  objFSO.MoveFile strOVDIR + "\databases\coda*.*", strTMP + "\coda"
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot backup files in Coda folder " & Err.Description
    Err.Clear
  End If
  
   '<OV_DIR>\conf\OpC\nodeinfo   -> <TMP>\nodeinfo.conv
   strSource = strOVDIR + "\conf\OpC\nodeinfo"
   strDest   = strTMP   + "\nodeinfo.conv"
   If (Not backup_nodeinfo(strSource, strDest )) Then		
       bRetainAgentID = False					'##some error has occured 
       								'##a. nodeinfo file not present 
       								'##b. nodeinfo.conv file could not be created 
       								'##c. CoreID already exsits  
       
       PrintMsg "Skipping migration of AgentID to OvCoreID", "Info"
   End If

  '<OV_DIR>\conf\svcdisc\javaagent.cfg   -> <TMP>\javaagent.conv
   strSource = strOVDIR + "\conf\svcdisc\OvJavaAgent.cfg"
   strDest   = strTMP   + "\OvJavaAgent.conv"
  
   If (Not backup_svcdiscconfig(strSource, strDest )) Then		
       bMigrateDiscoveryConfig = False					'##some error has occured   
         								'##a. javaagent.cfg file not present   
         								'##b. javaagent.conv file could not be created        								
                 
    PrintMsg "Skipping migration of service discovery agent configuration", "Info"  
   End If

  BackupData = True
  
End Function

'**************************************************************
' Purpose: Restores ECS, coda and stored facts data
' Inputs: strDataDir - data dir for new agent
' Returns: If data is found and saved return True.
'          Otherwise return False.
'**************************************************************
Private Function RestoreData (ByVal strDataDir)
  Dim MyShell, objFSO
  Dim strTMP, aryDir(2), aryDir2(1)

  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  Set objFSO = CreateObject ("Scripting.FileSystemObject")

  ' create target directories
  ' determine TMP directory
  strTMP = GetTempDir()

  ' create target directories if neccessary
  If (Not objFSO.FolderExists (strDataDir + "\conf\Bbc")) Then
    aryDir(0) = "conf"
    aryDir(1) = "Bbc"
    If (Not CreateDirectoryPath (strDataDir, aryDir)) Then
      Set objFSO = Nothing
      Set MyShell = Nothing
      RestoreData = False
    End If
  End If

  If (Not objFSO.FolderExists (strDataDir + "\conf\eaagt")) Then
    aryDir(0) = "conf"
    aryDir(1) = "eaagt"
    If (Not CreateDirectoryPath (strDataDir, aryDir)) Then
      Set objFSO = Nothing
      Set MyShell = Nothing
      RestoreData = False
    End If
  End If

  If (Not objFSO.FolderExists (strDataDir + "\datafiles")) Then
    aryDir2(0) = "datafiles"
    If (Not CreateDirectoryPath (strDataDir, aryDir2)) Then
      Set objFSO = Nothing
      Set MyShell = Nothing
      RestoreData = False
    End If
  End If

  ' copy *.ds data
  objFSO.CopyFile strTMP + "\ecs_fact\*.ds", strDataDir + "\conf\eaagt\", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot restore *.ds files " & Err.Description
    Err.Clear
  End If
  ' copy *.fs data
  objFSO.CopyFile strTMP + "\ecs_fact\*.fs", strDataDir + "\conf\eaagt\", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot restore *.fs files " & Err.Description
    Err.Clear
  End If
  ' move coda data
  objFSO.CopyFile strTMP + "\coda\default.txt", strDataDir + "\conf\Bbc\", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot restore Coda files " & Err.Description
    Err.Clear
  End If
  objFSO.MoveFile strTMP + "\coda\coda*", strDataDir + "\datafiles"
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot move coda folder  " & Err.Description
    Err.Clear
  End If
  
  If (bRetainAgentID) Then	'##Convert only if the AgentID value is read from nodeinfo file
  				'## Or if the CoreID for the node doesn't exists.
     If (Not convert_nodeinfo (strTMP + "\nodeinfo.conv")) Then
      PrintMsg "Either DCE node has no agentID assigned or nodeinfo file couldn't be backed up", "Info"
      PrintMsg "Skipping AgentID migration to CoreID.", "Info"
     Else
      PrintMSG "DCE AgentID is now migrated to HTTPS Agent CoreID.", "Info"
     End If
     
  End If
  
  ' remove saved data
  objFSO.DeleteFile strTMP + "\ecs_fact\*.ds", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot delete temporarily saved *.ds files " & Err.Description
    Err.Clear
  End If
  objFSO.DeleteFile strTMP + "\ecs_fact\*.fs", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot delete temporarily saved *.fs files " & Err.Description
    Err.Clear
  End If
  objFSO.DeleteFile strTMP + "\coda\*.*", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot delete temporarily saved coda files " & Err.Description
    Err.Clear
  End If  
  objFSO.DeleteFile strTMP + "\nodeinfo.conv", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot delete temporarily saved nodeinfo.conv file " & Err.Description
    Err.Clear
  End If
  
  ' remove temporary directories
  objFSO.DeleteFolder strTMP + "\ecs_fact", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot delete temporarily saved ecs_fact folder " & Err.Description
    Err.Clear
  End If
  objFSO.DeleteFolder strTMP + "\coda", True
  If (Err.Number <> 0) Then
    WriteToLog "opc_inst Error:Cannot delete temporarily saved coda folder " & Err.Description
    Err.Clear
  End If
  
  RestoreData = True
  
End Function

'**************************************************************
' Purpose: Print a message either as console output or message
'          box. Evaluate blnInteractive accordingly.
' Inputs:  -
' Returns: -
'**************************************************************
Private Sub PrintMsg (ByVal message, ByVal sev)
  Dim doPopup

  If (sev = "") Then
    doPopup = false
    sev = "Info"
  Else
    doPopup = true
  End If
    
  WriteToLog "opc_inst " & sev & ": " & message

  If (blnInteractive) Then
    If (isCscript) Then
      Wscript.echo "opc_inst " & sev & ": " & message
    Else
      If doPopup Then
        If (sev = "Error") Then
          MyShell.Popup message,, sev, 16
        Else
          MyShell.Popup message,, sev, 0
        End If
      End If
    End If
  ElseIf (sev = "Error") Then
    Wscript.echo "opc_inst " & sev & ": " & message
  End If
End Sub

'**************************************************************
' Purpose: Test whether the script runs within wscript or cscript
' Inputs:  -
' Returns: true if cscript, false otherwise
'**************************************************************

Private Function testCscript ()
  Dim scriptName
  Dim regEx
  scriptName = Wscript.FullName
  set regEx  = New RegExp
  regEx.Pattern = "cscript.exe$"
  regEx.IgnoreCase = True

  testCscript = regEx.test(scriptName)
  set regEx = Nothing
End Function

'**************************************************************
' Purpose: Determines where to install packages.
' Inputs: -
' Outputs: strInstDir - where to install packages.
'          strDataDir - data directory, not supported yet, always ""
' Returns: If inst_dir.tmp is found and processed return True.
'          Otherwise return False.
'**************************************************************
Private Function WhereToInstall (ByRef strInstDir, ByRef strDataDir)
  Dim objFSO
  Dim strFileName, strTmp
  Dim dat
  Dim MyArray(3)

  Const ForReading = 1

  ' establish error handling
  On Error Resume Next

  Set objFSO = CreateObject ("Scripting.FileSystemObject")

  If (objFSO.FileExists("inst_dir.tmp")) Then
    strFileName = "inst_dir.tmp"
  ElseIf (objFSO.FileExists("..\..\files\inst_dir.tmp")) Then
    strFileName = "..\..\files\inst_dir.tmp"
  Else
    Set objFSO = Nothing
    WhereToInstall = False
    Exit Function
  End If

  Set dat = objFSO.OpenTextFile(strFileName, ForReading, False)
  If (Err.Number <> 0) Then
    Err.Clear
    Set dat = Nothing
    Set objFSO = Nothing
    WhereToInstall = False
    Exit Function
  End If

  strTmp = dat.ReadLine ()
  If (Err.Number <> 0) Then
    Err.Clear
    dat.Close
    Set dat = Nothing
    Set objFSO = Nothing
    WhereToInstall = False
    Exit Function
  End If

  strTmp = Trim (strTmp)

  ' replace "/"s with "\"s
  strTmp = Replace (strTmp, "/", "\")
  ' replace "*"s with spaces
  strTmp = Replace (strTmp, "*", " ")

  If (strTmp <> "") Then
    ' workaround - colon is ommited from path if 
    ' install directory is entered in GUI
    If (Mid(strTmp, 2, 1) <> ":") Then
      ' insert ":" in path at position 1
      MyArray(0) = Left (strTmp, 1)
      MyArray(1) = ":"
      MyArray(2) = Right (strTmp, Len(strTmp)-1)
      strTmp = Join (MyArray, "")
    End If
    strInstDir = strTmp
    WhereToInstall = True
  Else
    WhereToInstall = False
  End If

  dat.Close

  Set dat = Nothing
  Set objFSO = Nothing

End Function

'**************************************************************
' Purpose: Checks if NO BOOT was requested by mgmt server
' Inputs: -
' Outputs: -
' Returns: If NO BOOT was requested return True.
'          Otherwise return False.
'**************************************************************
Private Function CheckForNoBoot ()
  Dim fso
  Dim RetVal

  ' we establish error handling
  On Error Resume Next

  ' False by default
  RetVal = False

  Set fso = CreateObject ("Scripting.FileSystemObject")

  If ( fso.FileExists("no_boot") _
       Or fso.FileExists("..\..\files\no_boot") ) Then
    RetVal = True
  End If
  
  Set fso = Nothing
  CheckForNoBoot = RetVal
End Function

'**************************************************************
' Purpose: Dump contents of text file
' Inputs:  fName  - Path of file to be printed (if present)
'          prefix - Some text to print before
'**************************************************************

Private Sub DumpFile (ByRef fName, ByRef prefix)
  Dim objFSO
  Dim dat, isOK
  Const ForReading = 1

  isOK = False

  On Error Resume Next
  Set objFSO = CreateObject ("Scripting.FileSystemObject")

  If (objFSO.FileExists(fName)) Then
    Set dat = objFSO.OpenTextFile(fName, ForReading, False)
    If (Err.Number = 0) Then
      PrintMsg prefix & Chr(10) & dat.ReadAll, ""
      isOK = True
      dat.Close
      Set dat = Nothing
    End If
  End If

  If (Not isOK) Then
    PrintMsg "Could not print file " & fName & ".", "Warning"
  End If

  Set objFSO = Nothing
End Sub

'**************************************************************
' Purpose: Test MSI installer version
' Inputs:  minVers - minimum version
' Output:  vers    - actual version
' Return:  true is OK, false otherwise
'**************************************************************

Private Function TestMsiVersion (ByVal minVers, ByRef vers)
  Dim strRegKey, MyShell, RetVal, strKeyValue
  Dim Result

  PrintMsg "Testing MSI installer version. Need at least " & minVers & ".", ""
 
  RetVal = False
  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")

  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\DataAccess\FullInstallVer"
  vers = MyShell.RegRead(strRegKey)
  If (Err.Number <> 0) Then
    PrintMsg "Could not determine MSI installer version from registry " & _
             strRegKey & ".", "Error"
    Err.Clear
    RetVal = False
  Else
    Dim regEx

    PrintMsg "MSI installer version is " & vers & ".", ""

    set regEx  = New RegExp
    regEx.Pattern = "^" & minVers
    regEx.IgnoreCase = True
    RetVal = regEx.test(vers)
    set regEx = Nothing
  End If

  Set MyShell = Nothing
  TestMsiVersion = RetVal
End Function

'**************************************************************
' Purpose: Print MANAGER name
'**************************************************************

Private Sub PrintManager()
  Dim fso
  Dim MyShell, RetVal, Result
  Dim strTMP
  Dim outFile, errFile

  ' establish error-handling
  On Error Resume Next

  Set MyShell = CreateObject ("WScript.Shell")

  ' determine <TMP>
  strTMP = GetTempDir()

  outFile = strTMP & "\out.out"
  errFile = strTMP & "\out.err"

  ' run ovconfget sec.core.auth MANAGER to get name of Mgmt Srv
  ' we have to process output of this program
  ' cmd.exe is used to enable redirection of output 
  ' output is redirected to temporary file out.out
  Result = MyShell.Run ("cmd.exe /c ""ovconfget.exe sec.core.auth MANAGER" & _
                        " >""" & outFile & """" & _
                        " 2>""" & errFile & """""", _
                        0, True)

  ' error-handling
  If ((Err.Number <> 0) Or (Result <> 0)) Then
    PrintMsg "Failed to determine Management Server name.", "Error"
    DumpFile errFile, "Error from ovconfget: "
    Err.Clear
  Else
    DumpFile outFile, "Management Server: "
  End If

  Set fso = CreateObject ("Scripting.FileSystemObject")
  fso.DeleteFile outFile
  fso.DeleteFile errFile

  Set fso = Nothing
  Set MyShell = Nothing
End Sub

'**************************************************************
' Purpose: Create installation lockfile
' Inputs: -
' Output: -
' Return:  true is OK, false otherwise
'**************************************************************
Private Function CreateInstLockFile ()
  Dim fso, MyFile
  Dim strLockFile

  strLockFile = GetTempDir() & "\" & lockFileName
  CreateInstLockFile = True

  On Error Resume Next
  Set fso = CreateObject("Scripting.FileSystemObject")

  Set MyFile = fso.CreateTextFile(strLockFile, True)
  If (Err.Number <> 0) Then
    Err.Clear
    CreateInstLockFile = False
  Else
    MyFile.Close
  End If

  Set MyFile = Nothing
  Set fso = Nothing
End Function

'**************************************************************
' Purpose: Create installation notification file
' Inputs: -
' Output: -
' Return:  true is OK, false otherwise
'**************************************************************
Private Function CreateInstNotifyFile ()
  Dim fso, MyFile
  Dim strLockFile

  strLockFile = strDataDir + "\tmp\OpC\install_notif"
  CreateInstNotifyFile = True

  On Error Resume Next
  Set fso = CreateObject("Scripting.FileSystemObject")

  If (Not fso.FolderExists (strDataDir + "\tmp\OpC")) Then
	fso.CreateFolder strDataDir + "\tmp"
	fso.CreateFolder strDataDir + "\tmp\OpC"
  End If

  Set MyFile = fso.CreateTextFile(strLockFile, True)
  If (Err.Number <> 0) Then
    Err.Clear
    CreateInstNotifyFile = False
  Else
    MyFile.Close
  End If

  Set MyFile = Nothing
  Set fso = Nothing
End Function

'**************************************************************
' Purpose: Remove file
' Inputs: -
' Output: -
' Return:  true is OK, false otherwise
'**************************************************************
Private Function RemoveFile (ByVal strFile)
  Dim fso

  RemoveFile = True

  On Error Resume Next
  Set fso = CreateObject("Scripting.FileSystemObject")

  fso.DeleteFile strFile, True
  If (Err.Number <> 0) Then
    Err.Clear
    RemoveFile = False
  End If

  Set fso = Nothing
End Function

'**************************************************************
' Purpose: Remove installation lockfile
' Inputs: -
' Output: -
' Return:  true is OK, false otherwise
'**************************************************************
Private Function RemoveInstLockFile ()
  RemoveInstLockFile = RemoveFile(GetTempDir() & "\" & lockFileName)
End Function

'**************************************************************
' Purpose: Check installation lockfile
' Inputs: -
' Output: strWhereIsIt - location of lockfile if it exists
' Return:  If lockfile is found return true, false otherwise
'**************************************************************
Private Function CheckInstLockFile (ByRef strWhereIsIt)
  Dim fso
  Dim strLockFile

  strLockFile = GetTempDir() & "\" & lockFileName

  On Error Resume Next

  Set fso = CreateObject("Scripting.FileSystemObject")

  If ( Not fso.FileExists(strLockFile) ) Then
    CheckInstLockFile = False
  Else
    strWhereIsIt = strLockFile
    CheckInstLockFile = True
  End If

  Set fso = Nothing
End Function

'**************************************************************
' Purpose: Clean up and exit
' Inputs: exit code
' Output: -
'**************************************************************
Private Sub DoExit (ByVal exitCode)

  ' switch back to Execute mode, if we switched on install mode on 
  ' a terminal server
  TSInstallMode = TSSetInstallMode(TSInstallMode)
  
  ' copy logfile to <DataDir>\log
  CopyLogFile

  ' Removal of lock file
  RemoveInstLockFile
  WScript.Quit (exitCode)
End Sub

'**************************************************************
' Purpose: Checks if file exists.
' Inputs: strFileName - file.
' Outputs: -
' Returns: true if strFileName exists, false otherwise.
'**************************************************************
Private Function DoesFileExist (ByVal strFileName)
  Dim fso

  On Error Resume Next

  Set fso = CreateObject ("Scripting.FileSystemObject")
  If ( fso.FileExists(strFileName) ) Then
    DoesFileExist = True
  Else
    DoesFileExist = False
  End If

  Set fso = Nothing
End Function

'**************************************************************
' Purpose: Backups OVO Settings
' Inputs: -
' Outputs: -
' Returns: true if everything is ok, false otherwise.
'**************************************************************
Private Function BackupOVOSettings ()
  Dim fso
  Dim MyShell, RetVal, Result
  Dim strTMP
  Dim errFile
  Dim strDataDir, strFileName

  ' establish error-handling
  On Error Resume Next

  Set MyShell = CreateObject ("WScript.Shell")

  ' determine <TMP>
  strTMP = GetTempDir()

  errFile = strTMP & "\out.err"

  If (Not GetDataDir (strDataDir)) Then
    Set MyShell = Nothing
    BackupOVOSettings = False
    Exit Function
  End If

  strFileName = strDataDir & "log\OVO_settings_backup.log"

  ' run ovconfget > <DataDir>\log\OVO_settings_backup.log
  ' cmd.exe is used to enable redirection of output
  ' error output is redirected to temporary file err.out
  Result = MyShell.Run ("cmd.exe /c ""ovconfget.exe" & _
                        " >""" & strFileName & """" & _
                        " 2>""" & errFile & """""", _
                        0, True)

  ' error-handling
  If ((Err.Number <> 0) Or (Result <> 0)) Then
    PrintMsg "Failed to save OVO settings.", "Error"
    DumpFile errFile, "Error from ovconfget: "
    Err.Clear
    Set MyShell = Nothing
    BackupOVOSettings = False
    Exit Function
  End If

  Set fso = CreateObject ("Scripting.FileSystemObject")
  fso.DeleteFile errFile

  PrintMsg "A backup copy of current settings is saved in " _
           & strFileName & ".", ""

  Set fso = Nothing
  Set MyShell = Nothing
  BackupOVOSettings = True
End Function

'*********************************************************
' Purpose: Compares two version strings.
' Inputs:  ver1, ver2: version strings
' Outputs: -
' Returns: 0: if versions are equal
'          1: if first version is greater
'          2: if second version is greater
'         -1: if error occured
'*********************************************************
Private Function CompareVersions (ByVal ver1, ByVal ver2)
  Dim aryTmp1, aryTmp2
  Dim aryVer1(4), aryVer2(4)
  Dim RetVal, i

  ' establish error handling
  On Error Resume Next

  ' split numbers from version string
  aryTmp1 = Split (ver1, ".", -1, 1)
  aryTmp2 = Split (ver2, ".", -1, 1)

  ' extract numbers to array
  For i = 0 To 3
    If ( i > UBound (aryTmp1) ) Then
      'if number is missing use zero
      aryVer1(i) = 0
    Else
      aryVer1(i) = Eval ( aryTmp1(i) )
      'if error occured during conversion from string to number
      If (Err.Number <> 0) Then
        Err.Clear
        ' use zero
        aryVer1(i) = 0
      End If
    End If

    If ( i > UBound(aryTmp2) ) Then
      'if number is missing use zero
      aryVer2(i) = 0
    Else
      aryVer2(i) = Eval ( aryTmp2(i) )
      'if error occured during conversion from string to number
      If (Err.Number <> 0) Then
        Err.Clear
        ' use zero
        arryVer2(i) = 0
      End If
    End If
  Next

  ' perform comparison
  If ( aryVer1(0) < aryVer2(0) ) Then
    RetVal = 2
  ElseIf ( aryVer1(0) > aryVer2(0) ) Then
    RetVal = 1
  ElseIf ( aryVer1(1) < aryVer2(1) ) Then
    RetVal = 2
  ElseIf ( aryVer1(1) > aryVer2(1) ) Then
    RetVal = 1
  ElseIf ( aryVer1(2) < aryVer2(2) ) Then
    RetVal = 2
  ElseIf ( aryVer1(2) > aryVer2(2) ) Then
    RetVal = 1
  ElseIf ( aryVer1(3) < aryVer2(3) ) Then
    RetVal = 2
  ElseIf ( aryVer1(3) > aryVer2(3) ) Then
    RetVal = 1
  Else
    RetVal = 0
  End If

  CompareVersions = RetVal

End Function

Sub OvCslCtrlRegComponent(Source,Destination)

	'On Error Resume Next

	Dim WshShell, hr, fileSystemObj

	Set WshShell = CreateObject("WScript.Shell") 	
	Set fileSystemObj=CreateObject("Scripting.fileSystemObject")

	If (fileSystemObj.FileExists(Source)=True) Then

		Dim  cmd
		'Destination = strDataDir & "\installation\inventory\Operations-agent.xml"
		cmd = "CMD /C copy /Y """ & Source & """ """ & Destination & """"		
		hr = WshShell.Run (cmd,0,TRUE)

	End If	
	
	Set fileSystemObj = Nothing
	Set WshShell = Nothing
	
End Sub

Sub OvCslCtrlUnregComponent(Source)

	'On Error Resume Next
       Dim fileSystemObj, Destination,WshShell,strRegKey

       Set WshShell = CreateObject ("WScript.Shell")
	Set fileSystemObj=CreateObject("Scripting.fileSystemObject")
      ' get DataDir from registry
	strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\DataDir"
	strDataDir = WshShell.RegRead(strRegKey)
      If ( NOT IsEmpty(strDataDir)) Then
         Destination = strDataDir & "\installation\inventory\" & Source
       	If (fileSystemObj.FileExists(Destination)=True) Then
	      	fileSystemObj.DeleteFile(Destination)
      	End If	
      End If
	Set fileSystemObj = Nothing
     	Set WshShell = Nothing

End Sub

'*********************************************************
' Purpose: Gets version of package from package descriptor.
' Inputs:  strXmlFile: name of package descriptor file.
' Outputs: strVersion: version of package.
' Returns: True if everything is ok,
'          False otherwise
'*********************************************************
Private Function GetXmlVersion (ByVal strXmlFile, ByRef strVersion)
  Const ForReading = 1
  Dim fso, file
  Dim strLine
  Dim start, ending

  'establish error handling
  On Error Resume Next

  Set fso = CreateObject ("Scripting.FileSystemObject")

  'open package descriptor
  Set file = fso.OpenTextFile (strXmlFile, ForReading, False)
  If (Err.Number <> 0) Then
    Err.Clear
    Set file = Nothing
    Set fso = Nothing
    GetXmlVersion = False
    Exit Function
  End If

  'check every line
  Do While (file.AtEndOfStream <> True)
    strLine = file.ReadLine ()
    strLine = Trim (strLine)

    'find line with string <version>....</version>
    If ( InStr (1, strLine, "<version>", 1) > 0 ) Then
      If ( InStr (1, strLine, "</version>", 1) > 0) Then
        'extract version string
        start = InStr (1, strLine, "<version>", 1) + 9
        ending = InStr (1, strLine, "</version>", 1) - 1
        strVersion = Mid (strLine, start, ending - start + 1)
        'determine if "A." is prepended
        If ( Left (strVersion,2) = "A." ) Then
          'strip it
          strVersion = Right (strVersion, Len(strVersion) - 2)
        End If
        file.Close
        Set file = Nothing
        Set fso = Nothing
        GetXmlVersion = True
        Exit Function        
      End If
    End If
  Loop
  
  file.Close
  Set file = Nothing
  Set fso = Nothing

  GetXmlVersion = False  

End Function

'*********************************************************
' Purpose: Gets version of package from registry.
' Inputs:  strPackage: name of package
'          dctGUIDs: dictionary of package update GUIDs
' Outputs: strVersion: version of package
' Returns: True if everything is ok,
'          False otherwise.          
'*********************************************************
Private Function GetInstalledVersion (ByVal strPackage, ByRef dctGUIDs, ByRef strVersion)
  Dim strRegKey, strKeyValue
  Dim MyShell
  Dim index1, index2
  Dim ver1, ver2, ver3

  'establish error handling
  On Error Resume Next

  Set MyShell = CreateObject ("WScript.Shell")

  'get version string from registry
  strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView"
  If ( dctGUIDs.Exists(strPackage) ) Then
    strRegKey = strRegKey & "\" & dctGUIDs.Item ( strPackage ) & "\ProductVersion"
    strKeyValue = MyShell.RegRead (strRegKey)
    If (Err.Number <> 0) Then
      Err.Clear
      Set MyShell = Nothing
      GetInstalledVersion = False
      strVersion = ""
      Exit Function
    End If
  End If

  strKeyValue = Trim (strKeyValue)
  
  'determine if "A." is prepended
  If ( Left (strKeyValue,2) = "A." ) Then
    'strip it
    strKeyValue = Right (strKeyValue, Len(strKeyValue) - 2)
  End If
  
  'tokenize
  ver1 = Left(strKeyValue, InStr(strKeyValue, ".") - 1)
  strKeyValue=Right(strKeyValue, Len(strKeyValue)  - InStr(strKeyValue, "."))
  ver2 = Left(strKeyValue, InStr(strKeyValue, ".") - 1)
  ver3 = Right(strKeyValue, Len(strKeyValue)  - InStr(strKeyValue, "."))

  ver1  = Left("00", 2 - Len(ver1)) & ver1
  ver2  = Left("00", 2 - Len(ver2)) & ver2
  ver3  = Left("000", 3 - Len(ver3)) & ver3

  strVersion = ver1 & "." & ver2 & "." & ver3

  Set MyShell = Nothing

  GetInstalledVersion = True

End Function

' **********************************************************
' cleanup the PendingFileRenameOperations registry key to
' avoid that the opcagt.cat is deleted after the next reboot
' **********************************************************
Private Function cleanPendingFileRenameOperations()

  const HKEY_LOCAL_MACHINE = &H80000002
  const strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
  const strKeyName = "PendingFileRenameOperations"
  const strToBeRemoved = "\??\C:\Program Files\HP OpenView\msg\C\opcagt.cat"

  Dim oReg
  Dim strKeyValueCurrent
  Dim blnFound
  Dim strKeyValueNew()
  Dim intSize

  blnFound = False
  intSize = 0

  'establish error handling
  On Error Resume Next

  Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
    ".\root\default:StdRegProv")
  If (Err.Number <> 0) Then
    Err.Clear
    Set oReg = Nothing
    PrintMsg "Could not remove opcagt.cat from PendingFileRenameOperations - WMI problem", "Warning"
    Exit Function
  End If
  oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,_
    strKeyName,strKeyValueCurrent
  If (Err.Number <> 0) Then
    Err.Clear
    Set oReg = Nothing
    PrintMsg "Could not remove opcagt.cat from PendingFileRenameOperations - reg read problem", "Warning"
    Exit Function
  End If

  ' Ok, now we have to walk through the strKeyValueCurrent array
  if (blnDebug) Then
    PrintMsg "Checking for opcagt.cat in PendingFileRenameOperations reg key",""
  End If

  For i = LBound(strKeyValueCurrent) To UBound(strKeyValueCurrent)
    if strKeyValueCurrent(i) = strToBeRemoved Then
      blnFound = True
      i = i + 1
    Else
      ReDim Preserve strKeyValueNew(intSize)
      strKeyValueNew(intSize) = strKeyValueCurrent(i)
      intSize = intSize + 1
    End If
  Next

  if blnFound Then
    PrintMsg "Removing opcagt.cat from PendingFileRenameOperations reg key",""
    oReg.SetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,_
      strKeyName,strKeyValueNew
    If (Err.Number <> 0) Then
      Err.Clear
      Set oReg = Nothing
      PrintMsg "Could not remove opcagt.cat from PendingFileRenameOperations - reg write problem", "Warning"
      Exit Function
    End If
  End If

End Function

'*********************************************************
' Purpose: Reboot the local system
'*********************************************************
Private Sub DoReboot()
  Dim objOSSet, objOS

  ' copy logfile to <DataDir>\log
  CopyLogFile
  ' Removal of lock file
  RemoveInstLockFile

  Set objOSSet = GetObject("winmgmts:{impersonationLevel=impersonate," & _
                 "(Shutdown)}!/root/cimv2").ExecQuery("select * from " & _
                 "Win32_OperatingSystem where Primary=true")

  For each objOS in objOSSet
    objOS.Reboot()
  Next
End Sub

'*********************************************************
' Purpose: Check whether a reboot is needed and prompt the user.
'*********************************************************
Private Sub EvalReboot(needReboot)
  Dim intButton

  If (needReboot) Then
    If (blnForce) Then
      'PrintMsg "At least one OV package requires a system reboot." & _
      '         Chr(10) & "Rebooting due to option force ...", ""
    Else
      If (blnInteractive) Then
        intButton = MyShell.Popup _
                    ("At least one OV package requires a system reboot." & _
                     Chr(10) & "Possibly files are still accessed." & _
                     Chr(10) & "Do you want to reboot now? " & _
                     Chr(10) & "If not, subsequent re-installations " & _
                     "of OVO may fail or get corrupted.",,, 4 + 32)
      Else
        PrintMsg "At least one OV package requires a system reboot." & _
                 Chr(10) & " - not rebooting without force." & _
                 Chr(10) & "It is strongly recommended to reboot the " & _
                 "system before re-installing OVO." & Chr(10) & _
                 "If not, subsequent re-installations " & _
                 "of OVO may fail or get corrupted.", "Warning"
        DoExit (1)
      End If
    End If

    If ((intButton = 6)) Then
      ' Reboot now ...
      DoReboot
    Else
      PrintMsg "System reboot not confirmed nor forced. Skipping it.", "Warning"
      PrintMsg "OVO Maintenance script ends", "Warning"
      DoExit (0)
  End If
  End If
End Sub

'Defect CR QXCR1000430963: 7.33 windows agent does not uninstall through 'opcsetup -u'.
'*********************************************************
' Purpose: Get Registry key and Value
' Inputs : 
' Outputs: 
' Returns: True   - If everything is favourable
'          False  - If something is not well 
'*********************************************************
Function GetRegKeyValue (ByRef strRegPath, ByRef strRegId, ByRef strRegValue)
  Dim MyShell, strDisplayName, strRegKey, basePath

  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")

  basePath = "HKEY_LOCAL_MACHINE\" & strRegPath
  strRegKey = basePath & "\" & strRegId

  strDisplayName = MyShell.RegRead (strRegKey)

  If (Err.Number <> 0) Then
    Err.Clear
    Set MyShell      = Nothing
    GetRegKeyValue   = False
    Exit Function
  End If

  If (strDisplayName = "") Then
    Set MyShell      = Nothing
    GetRegKeyValue   = False
    Exit Function
  End If

  strRegValue = strDisplayName
  GetRegKeyValue    = True

End Function

'******************************************************************
' Purpose:  Enumerate the registry keys for un-installation string
' Inputs:   RegRoot   -> Defaults to HKEY_LOCAL_MACHINE.
'           SPath     -> Path to the parent registry key.
'           SubKeyId  -> Sub key ID to match
'           SubKeyVal -> Sub key value to match with
' Outputs:  
' Returns: Valid Value:  If good
'          Empty ""   : If bad          
'*******************************************************************
Function GetUninstallKey(RegRoot, SPath, SubKeyId, SubKeyVal) 
  Dim sKeys() 
  Dim SubKeyCount 
  Dim objRegistry 
  Dim sKeyVal
  Dim lRC
  Dim Key
  Dim sUninstallKey
  Set objRegistry = GetObject("winmgmts:root\default:StdRegProv") 
  GetUninstallKey = ""
  lRC = objRegistry.EnumKey(RegRoot, sPath, sKeys) 

  If (lRC = 0) And (Err.Number = 0) Then 
     for each Key in sKeys 
          'WScript.Echo "Found KEY: " & Key & " under :  " &  sPath            
          If 0 <> GetRegKeyValue(SPath & "\" & Key, SubKeyId, sKeyVal) Then
           'Compare sKeyVal is equal to "HP Operations Manager Manual Agent" -or- SubKeyVal
           if 0 <> StrComp(sKeyVal, SubKeyVal) Then
             'Comparison wasn't succesfull
             'WScript.Echo " No Luck in the key: " & Key
           Else
            'Comparison was succesfull
            'WScript.Echo " Found a match in the key: " & Key
            'Get the unistall key value
            GetRegKeyValue SPath & "\" & Key, "UninstallString", sUninstallKey
            GetUninstallKey = sUninstallKey
           End If
          End If          
      next      
  Else 
      'WScript.Echo "COULDN'T ENUMERATE KEY: " & sPath 
      GetUninstallKey = ""
  End If 
End Function

'*********************************************************
' Purpose: Get the agent de-install string
' Inputs : 
' Outputs: 
' Returns: True  - if everything is good
'          false - if something is foul
'*********************************************************
Function GetA07AgtInstString (ByRef strOvAgentInstallString)

 Dim RegHive 
 Dim RegKey 
 Dim KeyId  
 Dim KeyVal 
 Dim sUninstallString
 Const HKEY_LOCAL_MACHINE = 2147483650 

 'Set the filters
 KeyId  = "DisplayName"
 KeyVal = "HP Operations Manager Manual Agent"

 'Select registry hive constant from above list. 
 RegHive = HKEY_LOCAL_MACHINE 

 'Path to key to delete (no leading/trailing slashes). 
 RegKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" 

 ' Recursive sub to enumerate all subkeys and parent key. 
 ' This function specifically looks for the key with "DisplayName" value as "HP Operations Manager Manual Agent"
 sUninstallString = GetUninstallKey(RegHive, RegKey, KeyId, KeyVal)
 If "" <> sUninstallString  Then
   'WScript.Echo "Already found the key and uninstall key is - " & sUninstallString
   strOvAgentInstallString = sUninstallString
   GetA07AgtInstString     = True
 Else
   strOvAgentInstallString = ""
   GetA07AgtInstString     = False
 End If
End Function 

'*********************************************************
' Purpose: de-install OVO/W based DCE agent
' Inputs : 
' Outputs: 
' Returns: True  - if OVO/W DCE agent got de-installed
'          False - if no DCE agent was found
'*********************************************************
Function deinstall_ovowdceagent ( ) 

 Dim sDeInstallString
 Dim sRemoteInstallKey
 Dim Result
 Dim bFailed

 bFailed = 0
 deinstall_ovowdceagent=True

 If 0 <> GetA07AgtInstString(sDeInstallString) Then 
 'Agent was originally installed manually
  sDeInstallString = Replace(sDeInstallString, "/I", "/x")
  sDeInstallString = sDeInstallString & " /qr" 
  WScript.Echo "Agent was originally installed manually - " & sDeInstallString
 Else
 'Agent was installed using remote deployment
  sRemoteInstallKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OpenView Operations Agent"
  If 0 <> GetRegKeyValue(sRemoteInstallKey, "UninstallString", sDeInstallString) Then
   WScript.Echo "Info: Agent was originally deployed remotely, de-installing using command - " & sDeInstallString
  Else
   bFailed = 1
   WScript.Echo "Info: No OVO/W-DCE Agent found on the machine, check for OVO/U-DCE Agent next" 
   deinstall_ovowdceagent=False
  End If
 End If

 If 1 <> bFailed Then
  Dim shShell  
  Set shShell = WScript.CreateObject("WScript.Shell")  
  Result = shShell.Run (sDeInstallString, 1, True)  
  If (Err.Number <> 0) Then
   Err.Clear
   Set shShell = Nothing    
   WScript.Echo "Error:De-Installing the DCE agent using - " & sDeInstallString & " failed"
   deinstall_ovowdceagent=False
  End If
 End If 
End Function

'*********************************************************
' Purpose: Rename package.
' Inputs: strPackage   :name of current package to change.
'         strNewPackage:the new name
' Returns: If renaming is successful, return True.
'          Otherwise return False.
'*********************************************************
Private Function RenamePackage (ByRef strPackage, ByVal strNewPackage )

  Dim MyShell, Result
  Dim MyFSO

  On Error Resume Next  
  Set MyShell = CreateObject ("WScript.Shell")
  Set MyFSO = CreateObject ("Scripting.FileSystemObject")

  'return if source and destination is same.  
  If (strPackage <> strNewPackage ) Then
    'retunrs True if exists or False
    If (MyFSO.FileExists(strPackage)) Then
    'Do renaming of the file here
      MyFSO.MoveFile strPackage, strNewPackage
      If (Err.Number <> 0) Then
         RenamePackage = False
         Set MyShell = Nothing
         Set MyFSO = Nothing
         On Error Goto 0
         Exit Function  
      End If   
   Else
       RenamePackage = False
       Set MyShell = Nothing
       Set MyFSO = Nothing
       On Error Goto 0
       Exit Function
    End If  
  End If

   'Clean-up
   Set MyShell  = Nothing
   Set MyFSO    = Nothing
   RenamePackage = True
   On Error Goto 0
   
 End Function

'*********************************************************
 ' Purpose: Create scrambled product GUID.
 ' Inputs:  Product GUID eg: FCF34170-C53F-4C57-A7F4-16BACCF9A761
 ' Returns: Returns scrambled GUID
 '          
 '*********************************************************
 Private Function CreateScrambledGUID (ByRef productGUID, ByRef scrambledCode)

  Dim MyString, MyArray, Msg, counter1, counter2, scrambled, index, strTemp, partGUID
  Dim strLeft, i

  MyString = productGUID
  index = 0
  scrambled = ""

  'establish error handling
   On Error Resume Next

  MyArray = Split(MyString, "-", -1, 1) 

  'Msg = MyArray(0) & " " & MyArray(1)
  'Msg = Msg   & " " & MyArray(2) & " " & MyArray(3) & " " & MyArray(4)
  'MsgBox Msg

  counter1 = UBound(MyArray)

  For Each partGUID in MyArray
    If (index < 3) Then
     'Do whole string reverse
     scrambled = scrambled &  StrReverse (MyArray(index))    
    Else
     'Do 2 bytes reverse   
     strTemp  = MyArray(index)
     counter2 = Len(strTemp)    
     counter2 = counter2/2     
     i = 1
     For counter1 = 1 To Counter2
      strLeft = Mid(strTemp, i, 2 )
      i = i + 2     
      scrambled = scrambled &  StrReverse(strLeft)
     Next         
    End If   
    index = index + 1
  Next

  'Msg =  scrambled 
  'MsgBox Msg
  
  scrambledCode = scrambled 
  CreateScrambledGUID = True

 End Function
 
 '*********************************************************
 ' Purpose: Get product code from Package GUID.
 ' Inputs:  Package GUID eg: {93E26950-3687-4027-8804-5D06002C8A5D}
 ' Returns: Returns product GUID for the pacakage
 '          
 '*********************************************************
 Private Function GetProductGUID (ByVal strPackage, ByVal dctGUIDs, ByRef strProductCode)
  'Read the key "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\<packageGUID>"
  'Update the dictionary object 
  Dim strRegKey, strKeyValue
  
  'establish error handling
   On Error Resume Next
  
   Set MyShell = CreateObject ("WScript.Shell")  

   'get product code from registry
   strRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView"
   If ( dctGUIDs.Exists(strPackage) ) Then
      strRegKey = strRegKey & "\" & dctGUIDs.Item ( strPackage ) & "\ProductCode"      
      strKeyValue = MyShell.RegRead (strRegKey)
      If (Err.Number <> 0) Then
        Err.Clear        
        GetProductGUID = False
        Exit Function
      End If
   End If
   
   strKeyValue = Trim (strKeyValue)

   'determine if "{" is prepended
   If ( Left (strKeyValue,1) = "{" ) Then
    'strip it
    strKeyValue = Right (strKeyValue, Len(strKeyValue) - 1)      
   End If

   If ( Right (strKeyValue,1) = "}" ) Then
    'strip it
    strKeyValue = Left (strKeyValue, Len(strKeyValue) - 1)      
   End If   

   strProductCode = strKeyValue  

   GetProductGUID = True

 End Function
 
 '*********************************************************
 ' Purpose: Updates the MSI cache with the package name.
 ' Inputs:  Product scrambled code eg: 07143FCFF35C75C47A4F61ABCC9F7A16
 ' Returns:  
 '*********************************************************
 Private Function UpdateMSICache (ByVal scrambledCode, ByVal strPackageFileName)

  Set MyShell = CreateObject ("WScript.Shell")
 'get MSI Cache data from registry

 Dim strRegKey, strKeyValue
 
 'establish error handling
  On Error Resume Next

 strRegKey = "HKEY_CLASSES_ROOT\Installer\Products\" & scrambledCode & "\SourceList"
  strKeyValue = MyShell.RegRead (strRegKey & "\PackageName")  
  If (Err.Number <> 0) Then
    Err.Clear
    UpdateMSICache = False
    Exit Function
  End If

  'Update the value with strPackageFileName
  If (StrComp(strKeyValue, strPackageFileName, 1) = 0) Then   
   UpdateMSICache = True			'return False shall mean an error
   Exit Function  
  Else
   ' Update the MSI cache with the newer name
   MyShell.RegWrite strRegKey & "\PackageName", strPackageFileName, "REG_SZ"
   If (Err.Number <> 0) Then
      Err.Clear      
      UpdateMSICache = False
      Exit Function
   End If 
  End If

  UpdateMSICache = True

 End Function

Private Function GetOvConfData (ByRef param)
  Const ForReading = 1
  Dim MyShell,strTMP
  Dim strInstallDir, Result,fso,dat
  Dim outFile,errFile

  On Error Resume Next
  Set MyShell = CreateObject ("WScript.Shell")
  Set fso = CreateObject ("Scripting.FileSystemObject")
  If (Not GetInstallDir(strInstallDir)) Then
    GetOvConfData = ""
    Set MyShell = Nothing
    Set fso = Nothing
    Exit Function
  End If
  strTMP = GetTempDir()
  outFile = strTMP & "\out_conf.out"
  errFile = strTMP & "\out_conf.err"

  Result = MyShell.Run ("cmd.exe /c """"" & strInstallDir _
                      & "bin/ovconfget.exe""  "_
                      & param _
                      & " >""" & outFile & """" _
                      & " 2>""" & errFile & """", _
                        0, True)
   
  ' error-handling
  If (Err.Number <> 0) Then
    PrintMsg "Failed to get value for " & param , "Error"
    DumpFile errFile, "Error from ovconfget : "
    Err.Clear
    GetOvConfData = ""
    fso.DeleteFile(outFile)
    fso.DeleteFile(errFile)
    Set fso = Nothing
    Set MyShell = Nothing
    Exit Function
  End If
  ' process output only if ovconfget succeeded
  If (Result = 0) Then
    Set dat = fso.OpenTextFile (outFile, ForReading, False)
    Do While (dat.AtEndOfStream <> True)
      Dim strLine 
      strLine = dat.ReadLine ()
      'PrintMsg "Read from " & outFile & " ==> " & strLine , "Info"
	  If(Len(strLine) > 0) Then
           GetOvConfData = strLine
	  End if
	  'exit do here -
	  Exit Do
    Loop
  End If

  dat.Close
  fso.DeleteFile(outFile)
  fso.DeleteFile(errFile)

'''''''''''''''''''''''''''''''''''''
  Set fso = Nothing
  Set MyShell = Nothing
End Function

Private Function IsCoreIDSet(strLine)
  Dim position
  position = InStr(1, strLine, ":CORE_ID=", 1)
  If ( position > 0 ) Then
	IsCoreIDSet = Mid (strLine, position+9)
  Else
	IsCoreIDSet = ""
  End If
End Function

Private Function GetCoreIdFromProfile(ByVal strFileName)

  Dim objFSO,org_file,strLine,retStr
  Const ForReading = 1
  retStr = ""  
  Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
  If (objFSO.FileExists (strFileName)) Then
    Set org_file = objFSO.OpenTextFile(strFileName,ForReading, False)
  Else
    PrintMsg "Profile file " & strFileName & " does not exist","Warn" 
    Set objFSO =  nothing
    GetCoreIdFromProfile = retStr
    Exit Function
  End if  
  
  Do While (org_file.AtEndOfStream <> True)
    strLine = org_file.ReadLine ()
	retStr = IsCoreIDSet(strLine)
	If(Len(retStr) > 0) Then
	  PrintMsg "Value from Server " & retStr , "Info"
	  GetCoreIdFromProfile = retStr
	  org_file.Close
	  Set objFSO = Nothing
	  Exit Function
	  'exit from here
	End if
  Loop
  org_file.Close
  Set objFSO = Nothing
  GetCoreIdFromProfile = retStr
End Function

'*********************************************************
' MAIN
'*********************************************************
Dim dctPackageGUID, dctPackageFiles, pkgs, i, guid, scrambledCode
Dim MyShell, intButton, Result
Dim MySysEnv
Dim blnDeinstalledOldAgent
Dim objArgs
Dim blnInteractive, blnDeinstall, objFSO
Dim blnNoStart, blnNoBoot
Dim strOpcActivateCall
Dim strInstallDir
Dim blnInventory, strPackageList, missingPkgs
Dim blnForce, blnDebug, bInstallNotify
Dim blnSrv, strSrv  ' for -srv|-s <server> switch
Dim strConfigure  'for -configure|-c <conf. file> switch
Dim procEnv, curPath
Dim blnCoreUpdate ' true if we are updating one of CORE packages
Dim blnAnyUpdate ' true if we are updating at least one package
Dim blnAgentUpdate ' true if we are updating Agent package
Dim strDataDir 'DataDir path
Dim strWhereToInstall, strDataDirConf
Dim isCscript  ' true if we are running in cscript, false for wscript
Dim forceWscript
Dim needReboot, nrb ' needReboot
Dim msiVers
Dim strA07AgtDir ' install directory for A.07.XX Agent
Dim instLogPath
Dim actLogPath
Dim strLockFile ' Where is installation lockfile
Dim blnCertSrv, strCertSrv  ' for Certificate Server
Dim blnBreakDep ' Break dependencies (if needed)
Dim blnPackageSpecified, strPackageSpecified ' If package is specified
Dim strPkgVersion ' version of package (from HPOv*.xml)
Dim strInstVersion ' version of installed package (from registry)
Dim blnPermitInstall ' do we allow installation of package?
Dim bRetainAgentID   ' True if agentID has to be restored else False
Dim bMigrateDiscoveryConfig    ' True if javaagent.cfg exists else False
Dim ValueOnNode , TempCmd ' To check if core id already set in the node  QXCR1000353406
Dim TSInstallMode    ' True if we can install right away, False, if we found a TS in EXECUTE mode

Set MyShell = CreateObject ("WScript.Shell")
Set MySysEnv = MyShell.Environment ("SYSTEM")

Const minMsiVers = "2|6.*"  ' Minimum MSI installer version. Is a regex
Const XplName = "HPOvXpl"
Const AgtExName = "HPOvAgtEx"
Const instLogName = "opc_inst.log"
Const actLogName  = "opcactivate.log"
Const lockFileName = "opc_inst.lock"
Const opcinfoConvFile = "opcinfo.conv"
Const opcinfoConvBat = "opcinfo.bat"
Const pf = "WinNT4.0-release"

Set dctPackageGUID = CreateObject ("Scripting.Dictionary")
dctPackageGUID.Add XplName,     "{93E26950-3687-4027-8804-5D06002C8A5D}"
dctPackageGUID.Add "HPOvSecCo", "{FF7FB1EA-7D65-4984-8C30-07721F804F9D}"
dctPackageGUID.Add "HPOvBbc",   "{CB50535E-9EA4-4C26-9701-D5A2155BE929}"
dctPackageGUID.Add "HPOvSecCC", "{F1BAB3F3-25F3-4EE2-B07D-29ED86DC2326}"
dctPackageGUID.Add "HPOvCtrl",  "{FF578D75-B1DC-4E87-8C55-D8FE41FE62FD}"
dctPackageGUID.Add "HPOvDepl",  "{478FA165-04FD-439F-8B5B-D5A87E606159}"
dctPackageGUID.Add "HPOvConf",  "{8298AD72-7ADD-475C-9556-92F877C9333E}"
dctPackageGUID.Add "HPOvPerlA", "{752f3981-8a03-11d7-86fa-00108301d3a3}"
dctPackageGUID.Add "HPOvPacc",  "{1e3873f7-6238-4154-9b4d-521f541377b3}"
dctPackageGUID.Add "HPOvPCO",   "{26384c92-6f09-480d-a721-8aec1e1898ad}"
dctPackageGUID.Add "HPOvEaAgt", "{95e8bc5c-c38d-42e0-9983-fe70fce81fa2}"
dctPackageGUID.Add "HPOvLcja",  "{B2422FF1-081B-4AD4-B77B-9007F532DB96}"
dctPackageGUID.Add "HPOvLces",  "{C3084B86-9527-42C5-8432-5A295F686671}"
dctPackageGUID.Add "HPOvLcko",  "{12368FFE-07E7-4F91-8B11-92643C66F288}"
dctPackageGUID.Add "HPOvLczC",  "{9A6A55C8-1496-4DE6-A4EF-4A92E667CFED}"
dctPackageGUID.Add "HPOvEaAja", "{33031540-B025-4509-ACAC-D1682E0111B1}"
dctPackageGUID.Add "HPOvEaAes", "{9E669239-6C16-4E58-BB4B-669A1B9ACFE0}"
dctPackageGUID.Add "HPOvEaAko", "{60FA76B2-E431-4056-9613-9365709DFF0E}"
dctPackageGUID.Add "HPOvEaAzC", "{E04C77CD-FAED-4A03-BD9F-B52DFF9C47D3}"
dctPackageGUID.Add "HPOvXercesA", "{8CDA8784-37E7-481E-A440-10F2630183A7}"
dctPackageGUID.Add "HPOvXalanA", "{70C382BC-7ACA-43F2-B4C6-72FB47D9371A}"
dctPackageGUID.Add "HPOvAgtEx", "{A79A754D-F07E-4805-9757-32F41A33C615}"

pkgs = dctPackageGUID.Keys

isCscript = testCscript()
forceWscript = False

'Installation is interactive by default
blnInteractive = True

'Installation is default operation
blnDeinstall = False

'Agent is started by default
blnNoStart = False

'Agent service will be installed as auto-start by default
blnNoBoot = False

'No inventory by default
blnInventory = False

'No force by default
blnForce = False

'No server by default
blnSrv = False
strSrv = ""

' Standard config file configure by default
strConfigure = "chg00000.job"

'No CORE update by default
blnCoreUpdate = False

' Nothing to update by default
blnAnyUpdate = False
blnAgentUpdate = False

strWhereToInstall = ""
strDataDirConf    = ""

'no certificate server by default
blnCertSrv = False
strCertSrv = ""

'package is not specified by default
blnPackageSpecified = False
strPackageSpecified = ""

' no break_dep by default
blnBreakDep = False

blnDebug = False ' changes msiexec log mode to /l*v - otherwise /le+
needReboot = False

'Default convert nodeinfo variables
bRetainAgentID = True	

bMigrateDiscoveryConfig = True    

'By default notification file would be created to send Unmanaged node information
bInstallNotify = True

' check for NO BOOT
blnNoBoot = CheckForNoBoot()

instLogPath = GetLogFileName(instLogName)

' ----------------------------------------------------------
PrintMsg "OVO (De-)Installation script starting at " & Now() & "." & _
         Chr(10) & "Log file: " & instLogPath, ""
' ----------------------------------------------------------

'Check command line arguments
Set objArgs = WScript.Arguments
For i = 0 To objArgs.Count - 1
  'If switch -non_int|ni is supplied, do non-interactive installation
  If ((objArgs(i) = "-non_int") Or (objArgs(i) = "-ni")) Then
    blnInteractive = False
  ElseIf ((objArgs(i) = "-wscript") Or (objArgs(i) = "-w")) Then
    forceWscript = True
  'If switch -remove|r is supplied, deinstall packages
  ElseIf ((objArgs(i) = "-remove") Or (objArgs(i) = "-r")) Then
    blnDeinstall = True
  ' -no_start|-ns to skip starting of L-Core at the end of installation
  ElseIf ((objArgs(i) = "-no_start") Or (objArgs(i) = "-ns")) Then
    blnNoStart = True
  ' -no_boot|-nb to register service in manual start-mode
  ElseIf ((objArgs(i) = "-no_boot") Or (objArgs(i) = "-nb")) Then
    blnNoBoot = True
  ' -verify|-v to list installed L-Core packages
  ElseIf ((objArgs(i) = "-verify") Or (objArgs(i) = "-v")) Then
    blnInventory = True
  ' -force|-f to do reinstall of already installed packages
  ElseIf ((objArgs(i) = "-force") Or (objArgs(i) = "-f")) Then
    blnForce = True
  ' -debug to do enable debug logging for msiexec
  ElseIf (objArgs(i) = "-debug") Then
    blnDebug = True
  ' -configure|-c <conf. file> to specify configuration parameters
  ElseIf ((objArgs(i) = "-configure") Or (objArgs(i) = "-c")) Then
    i = i + 1 ' next argument must be configuration file
    If i <= objArgs.Count - 1 Then
      strConfigure = objArgs(i)
    End If
  ' -srv|-s <mgmt server> to specify management server
  ElseIf ((objArgs(i) = "-srv") Or (objArgs(i) = "-s")) Then
    blnSrv = True
    i = i + 1 ' next argument must be management server
    If i <= objArgs.Count - 1 Then
      strSrv = objArgs(i)
    End If
    ' here we check if certificate server is specified
    If (i + 2) <= objArgs.Count - 1 Then
      i = i + 1
      ' -cert_serv <certificate server>
      If (objArgs(i) = "-cert_srv") Then
        blnCertSrv = True
        i = i + 1
        strCertSrv = objArgs(i)
      Else
        i = i - 1
      End If
    End If
  ElseIf ((objArgs(i) = "-inst_dir") Or (objArgs(i) = "-id")) Then
    i = i + 1 ' next argument must be the install dir
    If i <= objArgs.Count - 1 Then
      strWhereToInstall = objArgs(i)
    End If
  ElseIf ((objArgs(i) = "-data_dir") Or (objArgs(i) = "-dd")) Then
i = i + 1 ' next argument must be the data dir
    If i <= objArgs.Count - 1 Then
      strDataDirConf = objArgs(i)
  End If

  ' -no_instnotify to avoid creating the installation notification file.
  ElseIf (objArgs(i) = "-no_instnotify") Then
    bInstallNotify = False
  ' -help|-h to print usage
  ElseIf ((objArgs(i) = "-help") Or (objArgs(i) = "-h")) Then
    PrintUsageInfo True
    WScript.Quit (0)
  ElseIf (objArgs(i) = "-break_dep") Then
    blnBreakDep = True
  Else
    ' check if package is specified
    strPackageSpecified = objArgs(i)
    If (dctPackageGUID.Exists(strPackageSpecified)) Then
      blnPackageSpecified = True
    Else
      PrintUsageInfo True
      WScript.Quit (1)
    End If
  End If
Next

If (not isCscript) Then
  If(forceWscript) Then
    PrintMsg "Wscript execution forced by -w option.", ""
  Else
    Dim choice
    Dim txt
    txt = "It is recommended to execute this script using cscript " & _
          "and not using wscript." & Chr(10) & _
        "With wscript the actual (de)-installation will run in the " & _
      "background and you won't get any progress messages -" & Chr(10) & _
      "in this case review the installation log file to follow " & _
      "the progress." & Chr(10) & _
      "If you want to proceed with wscript, please confirm with OK " & _
      "otherwise hit Cancel and restart the script as:" & Chr(10) & _
      "  cscript opc_inst.vbs <parameters>"
    choice = MyShell.popup(txt,, "Warning", 1 + 48)

    If (choice = 2) Then
      PrintMsg "User cancelled (de)-installation.", ""
      WScript.Quit (0)
    End If

    forceWscript = True
  End If
End If

' ----------------------------------------------------------
' Here we check if lock file exists - perhaps some other
' installation is in progress...
' ----------------------------------------------------------
If ( CheckInstLockFile (strLockFile) ) Then
  PrintMsg "Installation lock file found.", "Error"
  PrintMsg "If there is no other installation in progress installation lock file " & strLockFile & " could be removed.", "Info"
  WScript.Quit (1)
End If

' ----------------------------------------------------------
' Here we create lock file
' WARNING: Do not forget to remove it at exit!
' ----------------------------------------------------------
CreateInstLockFile

' ----------------------------------------------------------
' Check if we are running a Terminal Server in EXECUTE mode
' and switch it to INSTALL mode if so
' ----------------------------------------------------------
TSInstallMode = TSSetInstallMode(True)

' ----------------------------------------------------------
'Installed packages inventory
' ----------------------------------------------------------
If blnInventory Then
  PrintMsg "Retrieving inventory ...", ""

  missingPkgs = ""
  If GetInstalledPackages (dctPackageGUID, strPackageList, missingPkgs) Then
    PrintMsg "Installed packages: " & Chr(10) & strPackageList, ""
    If missingPkgs = "" Then
      DoExit (0)
    End If

    PrintMsg "Missing packages: " & Chr(10) & missingPkgs, "Error"
  Else
    PrintMsg "There are no OV packages installed at all.", "Error"
  End If

  DoExit (1)
End If

'-----------------------------------------------------------
'Unregistering the xerces/xalan and bundle descriptor
'-----------------------------------------------------------
If (blnDeinstall) Then
   OvCslCtrlUnregComponent("Operations-agent.xml")
End If

' ----------------------------------------------------------
'Deinstallation of packages
' ----------------------------------------------------------
If (blnDeinstall) Then
  'backup OVO settings
  BackupOVOSettings

  ' if only one package is required to be removed
  If (blnPackageSpecified) Then
    
    ' Remove Hotfix deployment bundle before deinstalling specified pkg EaAgt
    If (Not RemHotfixDpl(strPackageSpecified)) Then
        	PrintMsg "Error while deinstalling" & strPackageSpecified & " Hotfix deployment bundle. Exiting.", "Error"	
        	WScript.Quit (1)
    End If
   
    PrintMsg "De-installing package " & strPackageSpecified & ".msi ...", ""
    If ( IsPackageInstalled (strPackageSpecified, dctPackageGUID) ) Then
      If (Not DeinstallPackage (strPackageSpecified, _
                                dctPackageGUID, blnInteractive, _
                                blnDebug, nrb, blnBreakDep)) Then
        PrintMsg "Error while deinstalling package " & strPackageSpecified _
                 & ".", "Error"
        Set MyShell = Nothing

        DoExit (1)
      End If

      If (nrb) Then
        needReboot = True
      End If
    Else
      PrintMsg "Package " & strPackageSpecified _ 
               & " not found. Ignoring ...", "Warning"
    End If

    If (strPackageSpecified = XplName) Then
      ' Workaround: some directories have to be removed
      XplCleanup strDataDir, MySysEnv
    End If
    Set MyShell = Nothing

    DoExit (0)
  Else
    ' all packages are to be removed
    PrintMsg "De-installing OVO agent ...", ""

    'first we have to stop LCore
    KillLCore  'we dont care about result of this operation

    ' get DataDir from registry
    If ( Not GetDataDir (strDataDir) ) Then
      strDataDir = "\Program Files\HP OpenView\data\"
    End If

     For i = dctPackageGUID.Count - 1 To 0 Step -1
      ' Remove Hotfix deployment bundle
      If (Not RemHotfixDpl(pkgs(i)) ) Then
      	PrintMsg "Error while deinstalling Hotfix deployment bundle. Exiting.", "Error"	
       	WScript.Quit (1)
      End If
      
      PrintMsg "De-installing package " & pkgs(i) & ".msi ...", ""

      If ( IsPackageInstalled (pkgs(i), dctPackageGUID) ) Then
        If (Not DeinstallPackage (pkgs(i), dctPackageGUID, blnInteractive, _
                                  blnDebug, nrb, blnBreakDep)) Then
          PrintMsg "Error while deinstalling package " & pkgs(i) & ".", "Error"

    ' Try for other packages also they might not have any dependency like perl packages.

           'Set MyShell = Nothing

           'DoExit (1)
        End If

        If (nrb) Then
          needReboot = True
        End If
      Else
        PrintMsg "Package " & pkgs(i) & " not found. Ignoring ...", "Warning"
      End If
    Next

' start ovc (if only E/A was deinstalled there )
' and if restart of OVO is required

  If ( Not StartLCore () ) Then
    PrintMsg "All components deinstalled nothing to start", ""
  Else
    PrintMsg "Successfully started Components.", ""
  End If

    ' Workaround: some directories have to be removed
    XplCleanup strDataDir, MySysEnv

    PrintMsg "Packages successfully deinstalled.", "Info"
    EvalReboot needReboot

    Set MyShell = Nothing
    DoExit (0)
  End If
End If

' ----------------------------------------------------------
' Test for installer version. We need at least 2.0
' ----------------------------------------------------------

If (Not TestMsiVersion(minMsiVers, msiVers)) Then
  PrintMsg "MSI Installer version too low. Need at least " & minMsiVers & _
           " this is " & msiVers & ".", "Error"
  DoExit (1)
End If

' ----------------------------------------------------------
'Install vcredist (irrespective of manual/remote installation)
' ----------------------------------------------------------
If (Not RunExternalScript ("vcredist_x86.exe /Q:a /c:" & Chr(34) & "msiexec /i vcredist.msi /qn" & Chr(34))) Then
          PrintMsg "Failed to install vcredist, please run this utility manually" & _
          " and then restart installation","Error"
	  DoExit (1)
End If

' ----------------------------------------------------------
'Deinstallation of old 7.x agent
' ----------------------------------------------------------

PrintMsg "Testing for OVO 7.x agent ...", ""

blnDeinstalledOldAgent = False

'Check if A.07.XX Agent is installed
If (ChkOldAgent ()) Then
  Dim ovo7mgrName,strDataRegKey
  ovo7mgrName = ""

  PrintMsg "Detected OVO 7.x agent.", ""

  ' ----------------------------------------------------------------
  ' Convert old opcinfo file and store it to temporary file
  ' This also yields the original MgmtSrv name
  ' ----------------------------------------------------------------
  If ( GetA07InstDir (strA07AgtDir) ) Then
    PrintMsg "Converting OVO 7.x opcinfo file " & _
             "(original file will be left untouched) ...", ""
    If ( convert_to_ovo8(strA07AgtDir & "\bin\OpC\install\opcinfo", _
                         GetTempDir() & "\" & opcinfoConvFile, _
                         GetTempDir() & "\" & opcinfoConvBat, _
                         ovo7mgrName)) Then
      PrintMsg "Generated BAT script with OVO 7.x opcinfo entries " & _
               "for later activation.", ""
    End If
  End If

  If (ovo7mgrName <> "") Then
    PrintMsg "Local OVO 7.x agent belongs to Management Server " & _
             ovo7mgrName, ""
  Else
    PrintMsg "No Management Server found for local OVO 7.x agent.", "Warning"
    ovo7mgrName = "<unknown>"
  End If

  If (blnForce) Then
    PrintMsg "Found OVO 7.x agent - option force upgrades ...", ""
  Else
    If (blnInteractive) Then
      intButton = MyShell.Popup ("A.07.XX OVO Agent found belonging to " & _
         "Management Server " & ovo7mgrName & Chr(10) & _
         "Do you want to upgrade it to OVO 8? " & Chr(10) & _
         "If not, you cannot install the OVO 8 agent and this script will " & _
         "exit.",,, 4 + 32)
    End If
  End If

  If ((intButton = 6) Or blnForce Or Not blnInteractive) Then
    ' Backup ECS, coda and stored fact data
    If ( Not BackupData () ) Then
      PrintMsg "Backup of ECS/coda/stored fact data failed.", "Error"       
    End If

    If (Not DeinstallOldAgent ()) Then
      PrintMsg "Error while uninstalling A.07.XX Agent.", "Error"
      Set MyShell = Nothing
      DoExit (1)
    Else
      PrintMsg "A.07.XX Agent successfully uninstalled", ""
      blnDeinstalledOldAgent = True

      strDataRegKey="HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP Openview\DataDir"
      CleanRegistry(strDataRegKey)
      Wscript.echo "Completing the cleanup of registry for DataDir"

      strDataRegKey="HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP Openview\AgentDataDir"
      deleteregistry(strDataRegKey)

      strDataRegKey="HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP Openview\AgentInstallDir"
      deleteregistry(strDataRegKey)
 
    End If
  Else
    PrintMsg "OVO 7.x Agent de-installation not confirmed. Aborting.", "Error"
    DoExit (1)
  End If
Else
  PrintMsg "No OVO 7.x agent found.", ""
End If

' ----------------------------------------------------------
'Installation of HTTPS based agent
' ----------------------------------------------------------

PrintMsg "Installing OVO agent ...", ""

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Here we check if any of CORE packages is to be updated
For i = 0 To 5
  If ( objFSO.FileExists ( pkgs(i) & ".msi" ) ) Then
    blnCoreUpdate = True
    blnAnyUpdate = True
    Exit For
  End If
Next

' Here we check if there is any package at all to be updated
If (Not blnCoreUpdate) Then
  For i = 6 To 11
    If ( objFSO.FileExists ( pkgs(i) & ".msi" ) ) Then
      blnAnyUpdate = True
      Exit For
    End If
  Next
End If  

For i = 0 To 11    
If ( Not(blnForce And GetInstalledVersion (pkgs(i), _
                                               dctPackageGUID, _
                                               strInstVersion)) ) Then                                                
	' if version of package is lower that the version of already
	' installed package, installation must not be allowed
        If (CompareVersions (strPkgVersion, strInstVersion) = 0  Or _
           CompareVersions (strPkgVersion, strInstVersion) = 2 ) Then	      
	        blnPermitInstall = 0		
        Else
	        blnPermitInstall = 1		          
        End If      
Else
       	If (blnForce) Then
       		 blnPermitInstall = 1
        End If
End If    

' Here we remove the hotfix deployment bundle
If ( blnPermitInstall ) Then
	If ( objFSO.FileExists ( pkgs(i) & ".msi" ) ) Then
	     ' Remove Hotfix deployment bundle
	      If (Not RemHotfixDpl(pkgs(i)) ) Then
	              	PrintMsg "Error while deinstalling Hotfix deployment bundle. Exiting.", "Error"	
	              	WScript.Quit (1)
	      End If
	End If
End If
Next

' Here we stop LCore appropriately
If (blnCoreUpdate) Then
  KillLCore
ElseIf (blnAnyUpdate) Then
  StopLCore
End If

' here we check if inst_dir.tmp file with installation
' directory is supplied
If strWhereToInstall = "" Then
  WhereToInstall strWhereToInstall, strDataDirConf
End If

If strWhereToInstall <> "" Then
  PrintMsg "Non-default installation directory: " & strWhereToInstall & ".", ""
End If

'check and notify user if backup of OVO settings is found
' get DataDir from registry
If ( Not GetDataDir (strDataDir) ) Then
     strDataDir = "\Program Files\HP OpenView\data\"
End If
If ( DoesFileExist(strDataDir & "log\OVO_settings_backup.log") ) Then
  PrintMsg "A backup of OVO settings found in " _
           & strDataDir & "log\OVO_settings_backup.log.", ""
End If

If (blnPackageSpecified) Then
  ' ----------------------------------------------------------
  ' Install only specified package.
  ' ----------------------------------------------------------
  
      ' Remove Hotfix deployment bundle before installing specified pkg EaAgt
        If (Not RemHotfixDpl(strPackageSpecified)) Then
              	PrintMsg "Error while deinstalling" & strPackageSpecified & " Hotfix deployment bundle. Exiting.", "Error"	
              	WScript.Quit (1)
    End If
  
  If ( objFSO.FileExists ( strPackageSpecified & ".msi" ) ) Then
    PrintMsg "Installing package " & strPackageSpecified _
             & ".msi ...", ""
    nrb = False

    ' check if package descriptor exists
    If ( Not objFSO.FileExists ( strPackageSpecified & ".xml" ) ) Then
      PrintMsg "Package descriptor " _
               & strPackageSpecified & ".xml not found.", "Error"
      DoExit(1)
    End If

    If ( Not GetXmlVersion (strPackageSpecified & ".xml", _
                            strPkgVersion) ) Then
      PrintMsg "Could not get version string from " _
             & strPackageSpecified & ".xml", "Error"
      DoExit(1)
    End If

    If ( IsPackageInstalled ( pkgs(i), dctPackageGUID ) ) Then
      If ( Not GetInstalledVersion (strPackageSpecified, _
                                    dctPackageGUID, _
                                    strInstVersion) ) Then
        PrintMsg "Package " & pkgs(i) & "is installed but " _
               & "could not get its version from registry.", "Error"
        DoExit(1)
      End If
    End If

    If ( Not(blnForce And GetInstalledVersion (pkgs(i), _
                                               dctPackageGUID, _
                                               strInstVersion)) ) Then                                                
       ' if version of package is lower that the version of already
       ' installed package, installation must not be allowed
       If (CompareVersions (strPkgVersion, strInstVersion) = 0  Or _
           CompareVersions (strPkgVersion, strInstVersion) = 2 ) Then	      
        blnPermitInstall = 0		
       Else
        blnPermitInstall = 1		          
       End If      
    Else
       If (blnForce) Then
        blnPermitInstall = 1
       End If
    End If    

    If ( blnPermitInstall ) Then
      If (Not InstallPackage (strPackageSpecified, blnInteractive, _
                              blnForce, strWhereToInstall, strDataDirConf, _
                              blnDebug, nrb)) Then
        PrintMsg "Error while installing package " _
                 & strPackageSpecified & ".msi", "Error"
        DoExit (1)
      End If
    Else
      PrintMsg "Could not reinstall package with same or lower version", "Warning"
    End If

    If (nrb) Then
      needReboot = True
    End If

    ' Special to extend PATH by $InstallDir/bin needed by subsequent packages
    ' to find XPL libs.
    ' Do this only after a successful installation of the XPL package - then
    ' the registry should contain the InstallDir setting.
    
    If ( strPackageSpecified = XplName ) Then
      If (GetInstallDir (strInstallDir)) Then
        Set procEnv = MyShell.Environment ("Process")
        curPath = procEnv("PATH")
        procEnv("PATH") = curPath & ";" & strInstallDir & "\bin"        
      End If

      ' After XPL has been installed, we can set config variables. To influence
      ' the ovcd startup-mode we have to set START_ON_BOOT before the OvCtrl
      ' MSI package gets installed
      if(Not blnNoBoot) Then
        PrintMsg "Setting Auto-start flag for ovcd installation ...", ""
	If (Not RunExternalScript ("cmd /c " & """" & strInstallDir & "bin\ovconfchg" & """" & " -ns ctrl -set START_ON_BOOT true")) Then
          PrintMsg "Failed to set Auto-start flag for ovcd service.", _
                   "Error"
        End If
      End If

      ' Here we restore ECS, coda and stored fact data
      ' this point is appropriate since we now know which
      ' is installation directory
      If (blnDeinstalledOldAgent) Then
        If (GetInstallDir (strInstallDir)) Then
          If (Not RestoreData (strInstallDir & "\data") ) Then
            PrintMsg "Restore of ECS/coda/stored fact data failed.", "Error"
          End If
        Else
          PrintMsg "Could not determine installation directory.", "Error"
          PrintMsg "Restore of ECS/coda/stored fact data failed.", "Error"
        End If
      End If
    End If 
  Else
    PrintMsg "Package file " & strPackageSpecified _
             & ".msi does not exist.", "Warning"
  End If

Else
  ' ----------------------------------------------------------
  ' Install all packages.
  ' ----------------------------------------------------------
  Dim strOrigPkgName, strSource,strDestination, strTMP
  
  For i = 0 To dctPackageGUID.Count - 1
      ' Remove Hotfix deployment bundle
      If (Not RemHotfixDpl(pkgs(i)) ) Then
              	PrintMsg "Error while deinstalling Hotfix deployment bundle. Exiting.", "Error"	
              	WScript.Quit (1)
      End If
      If ( objFSO.FileExists ( pkgs(i) & ".msi" ) ) Then
      PrintMsg "Installing package " & pkgs(i) & ".msi ...", ""

      nrb = False

      ' check if package descriptor exists
      If ( Not objFSO.FileExists (  pkgs(i) & ".xml" ) ) Then
        PrintMsg "Package descriptor " _
                 & pkgs(i) & ".xml not found.", "Error"
        DoExit(1)
      End If

      If ( Not GetXmlVersion ( pkgs(i) & ".xml", _
                              strPkgVersion) ) Then
        PrintMsg "Could not get version string from " _
               & pkgs(i) & ".xml", "Error"
        DoExit(1)
      End If

      If ( IsPackageInstalled ( pkgs(i), dctPackageGUID ) ) Then
        If ( Not GetInstalledVersion (pkgs(i), _
                                      dctPackageGUID, _
                                      strInstVersion) ) Then
          PrintMsg "Package " & pkgs(i) & "is installed but " _
                 & "could not get its version from registry.", "Error"
          DoExit(1)
        End If
      End If

      If ( Not(blnForce And GetInstalledVersion (pkgs(i), _
                                                 dctPackageGUID, _
                                                 strInstVersion)) ) Then                                                
       ' if version of package is lower that the version of already
       ' installed package, installation must not be allowed
       If (CompareVersions (strPkgVersion, strInstVersion) = 0  Or _
           CompareVersions (strPkgVersion, strInstVersion) = 2 ) Then	      
         blnPermitInstall = 0		
       Else
         blnPermitInstall = 1		          
       End If        
      Else
        If (blnForce) Then
         blnPermitInstall = 1
       End If       
      End If    

      'initialise the value referring to dictionary object
      strOrigPkgName = pkgs(i)
       
       If ( blnPermitInstall ) Then          
       'Renaming should happen only for L-Core packages
       If (strOrigPkgName = "HPOvXpl" OR  strOrigPkgName = "HPOvSecCo" OR  strOrigPkgName = "HPOvBbc" OR strOrigPkgName = "HPOvSecCC" OR  strOrigPkgName = "HPOvCtrl" OR strOrigPkgName = "HPOvDepl" OR strOrigPkgName = "HPOvConf" OR strOrigPkgName = "HPOvPCO" OR   strOrigPkgName = "HPOvPacc") Then
         strOrigPkgName = pkgs(i) & "-" & strPkgVersion & "-" & pf        
         'PrintMsg "Renaming the package name " & pkgs(i) &  " to " & strOrigPkgName, "Info"        
       End If       
        strSource = pkgs(i) & ".msi"
        strDestination = strOrigPkgName & ".msi"
        If (Not RenamePackage (strSource, strDestination)) Then
          PrintMsg "Error while manipulating the inventory for " & pkgs(i) & ".msi", "Warning"
          'DoExit (1)
          strOrigPkgName = pkgs(i)
        End If

        strSource = pkgs(i) & ".xml"
        strDestination = strOrigPkgName & ".xml"
        If (Not RenamePackage (strSource, strDestination)) Then
            PrintMsg "Error while manipulating the inventory for " & pkgs(i) & ".xml", "Warning"
            'DoExit (1)
        End If       

        'Update the product MSI cache before installing      
	If (GetProductGUID (pkgs(i), dctPackageGUID, guid) = True) Then
	 If (CreateScrambledGUID (guid, scrambledCode) =  True) Then     	    
	    If (UpdateMSICache (scrambledCode, strOrigPkgName & ".msi") = False) Then	     
	     PrintMsg "Error while updating the MSI cache " & pkgs(i) & ".msi", "Warning"
	   End If
	 End If		
	End If 	

        If (Not InstallPackage (strOrigPkgName,blnInteractive, _
                                blnForce, strWhereToInstall, strDataDirConf, _
                                blnDebug, nrb)) Then
          PrintMsg "Error while installing package " & pkgs(i) & ".msi", "Error"
          DoExit (1)
        End If

        If (Not IsPackageInstalled ( pkgs(i), dctPackageGUID )) Then
          PrintMsg "Package " & pkgs(i) & " successfully installed " & _
                   "but not found in inventory.", "Error"
          DoExit (1)
        End If

        'Change back the package name as the installation is successfull
	strSource = strOrigPkgName & ".msi"
	strDestination = pkgs(i) & ".msi"
	If (Not RenamePackage (strSource, strDestination)) Then
	    PrintMsg "Error while manipulating the inventory for " & pkgs(i) & ".xml", "Warning"
	End If

	'Change the descriptor file to the original name
	strSource = strOrigPkgName & ".xml"
	strDestination = pkgs(i) & ".xml"
	'Change back the package name to the original
	If (Not RenamePackage(strSource, strDestination)) Then
	  PrintMsg "Error while manipulating the inventory for " & pkgs(i) & ".xml", "Warning"
        End If
      Else
        PrintMsg "Could not reinstall package with same or lower version", "Warning"
      End If

      If (nrb) Then
        needReboot = True
      End If

      ' Special to extend PATH by $InstallDir/bin needed by subsequent packages
      ' to find XPL libs.
      ' Do this only after a successful installation of the XPL package - then
      ' the registry should contain the InstallDir setting.
  
      If ( pkgs(i) = XplName ) Then
        If (GetInstallDir (strInstallDir)) Then
          Set procEnv = MyShell.Environment ("Process")
          curPath = procEnv("PATH")
          procEnv("PATH") = curPath & ";" & strInstallDir & "\bin"
        End If
	' QXCR1000353406 - after installing xpl check if a core id already exist in the node
	ValueOnNode = GetOvConfData(" sec.core  CORE_ID")
	If(Len(ValueOnNode) > 0) Then
	   PrintMsg "CORE_ID already set to " & ValueOnNode,"Warning"
	Else
	   ValueOnNode = GetCoreIdFromProfile(strConfigure)
	   If(Len(ValueOnNode) > 0) Then
	     PrintMsg "Setting core id from  profile data ..." & ValueOnNode, ""  
	     	TempCmd = "cmd /c ovconfchg -ns sec.core -set CORE_ID " & ValueOnNode
	     If (Not RunExternalScript (TempCmd)) Then
               PrintMsg "Failed to set core id from profile.", "Error"
             End If
	   Else
	     PrintMsg "new Core Id will be generated by sec core", "Info"
	   End If
	End if

        ' After XPL has been installed, we can set config variables. To influence
        ' the ovcd startup-mode we have to set START_ON_BOOT before the OvCtrl
        ' MSI package gets installed
        if(Not blnNoBoot) Then
          PrintMsg "Setting Auto-start flag for ovcd installation ...", ""
	  If (Not RunExternalScript ("""" & strInstallDir & "bin\ovconfchg" & """" & " -ns ctrl -set START_ON_BOOT true")) Then
            PrintMsg "Failed to set Auto-start flag for ovcd service.", _
                     "Error"
          End If
        End If

        ' Here we restore ECS, coda and stored fact data
        ' this point is appropriate since we now know which
        ' is installation directory
        If (blnDeinstalledOldAgent) Then
          If (GetInstallDir (strInstallDir)) Then
            If (Not RestoreData (strInstallDir & "\data") ) Then
              PrintMsg "Restore of ECS/coda/stored fact data failed.", "Error"
            End If
          Else
            PrintMsg "Could not determine installation directory.", "Error"
            PrintMsg "Restore of ECS/coda/stored fact data failed.", "Error"
          End If
        End If
      End If
      
      If (pkgs(i) = AgtExName ) Then
        ' determine TMP directory
  	 strTMP = GetTempDir()
         If (bMigrateDiscoveryConfig) Then      '## Migrate service discovery agent configurations      
             convert_svcdisc_config (strTMP + "\OvJavaAgent.conv")             
        End If
      End If

    Else
      PrintMsg "Package file " & pkgs(i) & ".msi does not exist.", "Warning"
    End If
  Next
End If

' ----------------------------------------------------------
' Determine InstallDir and DataDir
' ----------------------------------------------------------

If (Not GetInstallDir (strInstallDir)) Then
  PrintMsg "Error while getting InstallDir from Windows Registry.", "Error"
  DoExit (1)
End If

If (Not GetDataDir (strDataDir)) Then
  PrintMsg "Error while getting DataDir from Windows Registry.", "Error"
  DoExit (1)
End If

PrintMsg "Using InstallDir: " & strInstallDir & ", DataDir: " & strDataDir, ""

' ----------------------------------------------------------
' Set OvInstallDir and OvDataDir environment variables
' Actually this is the job of XPL - so remove this code here when XPL does it.
' ----------------------------------------------------------

If (MySysEnv ("OvInstallDir") = "") Then
  MySysEnv("OvInstallDir") = strInstallDir
End If
If (MySysEnv ("OvDataDir") = "") Then
  MySysEnv("OvDataDir") = strDataDir
End If

' ----------------------------------------------------------
' Copy the converted opcinfo file to the final place - this is for
' reference only, so just warn if this fails.
' Run the batch file generated from the OVO 7 opcinfo settings. This should
' not fail - if so exit and do not activate the OVO agent.
' ----------------------------------------------------------

If (blnDeinstalledOldAgent) Then
  If (Not CopyConverted) Then
    PrintMsg "Error while copying converted opcinfo file.", "Warning"
  End If

  If (Not RunExternalScript (GetTempDir() & "\" & opcinfoConvBat)) Then
    PrintMsg "Failed to activate the converted opcinfo entries." & _
             Chr(10) & "Verify generated file " & GetTempDir() & "\" & _
             opcinfoConvBat & " and run manually.", "Error"
    DoExit (1)
  Else
    PrintMsg "Activated converted OVO 7.x opcinfo file.", ""
  End If
End If

' ----------------------------------------------------------
' cleanup the PendingFileRenameOperations registry key to
' avoid that the opcagt.cat is deleted after the next reboot
' ----------------------------------------------------------
cleanPendingFileRenameOperations

' ----------------------------------------------------------

' Create an installation file under <DataDir>/tmp/OpC by default.
' If opc_inst is called with -no_instnotify option, then
' don't create the file.
' ----------------------------------------------------------

If (bInstallNotify) Then
	CreateInstNotifyFile
End If

'---------------------------------------------
'Calling the registration of OVO-Agent.xml
'---------------------------------------------

Dim fileSystemObj, Destination, Source, WshShell,strTmpRegKey,strTmpDataDir,strOpcinst,strTmpInstDir,DestinationBundle

Set WshShell = CreateObject ("WScript.Shell")
Set fileSystemObj=CreateObject("Scripting.fileSystemObject")

' get DataDir from registry
strTmpRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\DataDir"
strTmpDataDir = WshShell.RegRead(strTmpRegKey)
strTmpRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Hewlett-Packard\HP OpenView\InstallDir"
strTmpInstDir= WshShell.RegRead(strTmpRegKey) 

Destination = strTmpDataDir & "\installation\inventory\Operations-agent.xml"
OvCslCtrlRegComponent "OVO-Agent.xml", Destination 

strOpcinst= strTmpInstDir & "\bin\OpC\install\opc_inst.vbs"

DestinationBundle = strTmpDataDir & "\installation\bundles"

Destination = DestinationBundle & "\Operations-agent"

If (Not fileSystemObj.FolderExists(Destination) ) Then
       if (Not fileSystemObj.FolderExists(DestinationBundle) ) Then
            fileSystemObj.CreateFolder DestinationBundle
       End If     
       fileSystemObj.CreateFolder Destination
End If

Destination = strTmpDataDir & "\installation\bundles\Operations-agent\opc_inst.vbs"

OvCslCtrlRegComponent strOpcinst, Destination

Set fileSystemObj = Nothing
Set WshShell = Nothing

'---------------------------------------------
'Calling NOMULTIPLEPOLICIES
'--------------------------------------------

Dim strcmd
strCmd  = "cmd /c ovconfchg -ns conf.server -set NOMULTIPLEPOLICIES  mgrconf,msgforwarding,servermsi,ras"        	    
If (Not RunExternalScript (strCmd)) Then                             
    PrintMsg "Setting NOMULTIPLEPOLICIES failed.", "Info"
Else
    PrintMsg "Setting NOMULTIPLEPOLICIES  succesfull.", "Info"
End If

' ----------------------------------------------------------
' execute opcactivate.vbs script with parameters:
'  -non_int|ni  -  non-interactive operation
'  -no_start|ns -  do not start lcore after configuration
'  -configure|-c <conf. file> - provide config. file
'  -srv|-s <server> - provide management server
' ----------------------------------------------------------

If (Not blnNoStart) Then
    actLogPath = GetLogFileName(actLogName)
    PrintMsg "Executing opcactivate (see " & actLogPath & " for results) ...", ""
    
    strOpcActivateCall = "cscript """ & strInstallDir & _
		     "\bin\OpC\install\opcactivate.vbs"" -c " & _
		     """" & strConfigure & """"
    
    If (Not blnInteractive) Then
      strOpcActivateCall = strOpcActivateCall & " -ni"
    End If
If (blnNoStart) Then
  strOpcActivateCall = strOpcActivateCall & " -ns"
End If
    If (blnSrv) Then
      strOpcActivateCall = strOpcActivateCall & " -s " & strSrv
      If (blnCertSrv) Then
        strOpcActivateCall = strOpcActivateCall & " -cert_srv " & strCertSrv
      End If
    End If
    If (forceWscript) Then
      strOpcActivateCall = strOpcActivateCall & " -w "
    End If
    
    If (Not RunExternalScript ( strOpcActivateCall )) Then
      PrintMsg "Error while executing opcactivate.vbs.", "Error"
      DumpFile actLogPath, "Output of opcactivate: "
'  DoExit (1)
End If

    DumpFile actLogPath, "Output of opcactivate: "
    RemoveFile(actLogPath)
    
    PrintMsg "OVO Agent successfully installed.", "Info"

Else
    PrintMsg "OVO Agent packages successfully installed.", "Info"
    PrintMsg "To perform configuration, execute opcactivate.vbs manually.", "Info"
End If

EvalReboot needReboot

PrintMsg "OVO Maintenance script ends", "Info"
DoExit (0)
