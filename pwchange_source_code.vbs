#$language = "VBScript"
#$interface = "1.0"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' pwChange.vbs
' Release Production
' Version 1
'
' Release  Revision Change
' -------- -------- ------
' 08-11-09 1        Initial release
'
'
' Supported Network Devices:
'   Enterasys Matrix 1H582-51
'   Enterasys Matrix 1H582-25 (not tested)
'   Enterasys Matrix N7 Platinum
'   Enterasys Matrix N5 Platinum
'   Enterasys Matrix X16
'   Enterasys X-Pedition ER16
'   Enterasys X-Pedition Security Router XSR-3020
'   Enterasys X-Pedition Security Router XSR-1805 (not tested)
'   Enterasys G3G124-24 (tested with reservations)
'   Enterasys C3G124-24 (not tested)
'   Cisco 3020 Blade switch (not tested)
'
' Console and Log Messages:
'   Message                                                      Routine
'   -------                                                      --------
'   Password changed.                                            MainLoop
'   The device is not supported.                                 MainLoop
'   No IP connectivity or inactive SSH2 server.                  AuthenticateSSH
'   SSH2 authentication failed due to bad username or password.  AuthenticateSSH
'   SSH2 authentication timed out.                               AuthenticateSSH
'   AuthenticateSSH2 function debug code error.                  AuthenticateSSH
'   IO_1H582_51 WaitForString #1 Timeout                         IO_1H582_51
'   IO_1H582_51 WaitForString #2 Timeout                         IO_1H582_51
'   IO_1H582_51 WaitForString #3 Timeout                         IO_1H582_51
'   IO_1H582_51 WaitForString #4 Timeout                         IO_1H582_51
'   IO_Matrix WaitForString #1 Timeout                           IO_Matrix
'   IO_Matrix WaitForString #2 Timeout                           IO_Matrix
'   IO_Matrix WaitForString #3 Timeout                           IO_Matrix
'   IO_Matrix WaitForString #4 Timeout                           IO_Matrix
'   IO_ER16 WaitForString #1 Timeout                             IO_ER16
'   IO_ER16 WaitForString #2 Timeout                             IO_ER16
'   IO_ER16 WaitForString #3 Timeout                             IO_ER16
'   IO_ER16 WaitForString #4 Timeout                             IO_ER16
'   IO_ER16 WaitForString #5 Timeout                             IO_ER16
'   IO_ER16 WaitForString #6 Timeout                             IO_ER16
'   IO_ER16 WaitForString #7 Timeout                             IO_ER16
'   IO_ER16 WaitForString #8 Timeout                             IO_ER16
'   IO_ER16 WaitForString #9 Timeout                             IO_ER16
'   IO_XSR3020 WaitForString #1 Timeout                          IO_XSR3020
'   IO_XSR3020 WaitForString #2 Timeout                          IO_XSR3020
'   IO_XSR3020 WaitForString #3 Timeout                          IO_XSR3020
'   IO_XSR3020 WaitForString #4 Timeout                          IO_XSR3020
'   IO_XSR3020 WaitForString #5 Timeout                          IO_XSR3020
'   IO_XSR3020 WaitForString #6 Timeout                          IO_XSR3020
'   IO_Cisco3020 WaitForString #1 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #2 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #3 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #4 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #5 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #6 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #7 Timeout                        IO_Cisco3020
'   IO_Cisco3020 WaitForString #8 Timeout                        IO_Cisco3020
'
' Message Boxes:
'   Message                                                      Routine
'   -------                                                      -------
'   Cannot open inventory file: [fn] Password Change will close. OpenInventory
'   Current password does not match, try again.                  Logon
'   New password does not match, try again.                      Logon
'   Press OK to start                                            OnClickStart
'   Password change has finished                                 OnClickStart

' Depenecies:
'   crt
'   InternetExplorer.Application
'   ADODB.Connection
'   ADODB.RecordSet
'   Scripting.FileSystemObject
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'*************************************************
'     Global Variable Definitions
'*************************************************
'User logon credentials
  Dim slogonUsername
  Dim slogonCurrentPassword
  Dim slogonNewPassword

'Log file
  Dim slogFile
  Dim fsologFile

'Inventory list
  Dim sinvListLocation
  Dim sinvListFileName
  Dim fsoinvListConnection
  Dim fsoinvListRecordSet

'Holds information about the current device in main loop
  Dim scurrentIPAddress
  Dim scurrentDeviceType
  Dim scurrentLocation
  Dim scurrentSystemName
  Dim scurrentDescription

'Internet Explorer object for Password Change console
  Dim objIE

'*************************************************
'     Main sub procedure
'*************************************************
Sub Main
  crt.Screen.Synchronous = True
  Call Startup                                   'Setup default values for selected global variables
  Call ShowConsole                               'Put the console onto the screen
  iInv = OpenInventory                           'Open the inventory file

  'continue if inventory file opened
  If iInv <> "EXIT" Then
    Do
      'wait for Start or Exit to be clicked
      'error occurs when application is quit using the close button
      On Error Resume Next
      button = objIE.Document.All("ButtonHandler").Value

      'detect if application has been quit using the close button
      If Err.Number <> 0 Then
        Exit Do
      End If
      on error goto 0      

      'process the button click
      Select Case button
        Case "Start"        
          'the start button has been clicked
          result = onClickStart

          'the exit button has been clicked
          If result = "EXIT" Then Exit Do
        
        Case "Exit"
        'the exit button has been clicked  
        Exit Do
      End Select
    Loop
    Call CloseInventory                            'Close the inventory file
  End If

  'close IE and SecureCRT
  objIE.Quit                                     'Quit Internet Explorer
  crt.Quit                                       'Quit SecureCRT
End Sub

'*************************************************
'     Startup : Setup default values for selected global variables
'*************************************************
Sub Startup
  sinvListLocation = "c:\pwchange\inventory\"
  sinvListFileName = "inventory.csv"
  sLogFile         = "c:\pwchange\log\" & Month(Date) & "-" & Day(Date) & "-" & Year(Date) & "@" & Hour(Time) & "-" & Minute(Time) & ".log"
End Sub


'*************************************************
'     ShowConsole : Put the console onto the screen
'*************************************************
Sub ShowConsole
  'minimize the SecureCRT window
  crt.Window.Show 2

  'create Internet Explorer object and wait until about:blank loads 
  set objIE = CreateObject("InternetExplorer.Application")
  objIE.Offline = True
  objIE.navigate "about:blank"

  Do
    crt.Sleep 100
  Loop While objIE.Busy

  'set background font and color to silver    
  objIE.Document.body.Style.FontFamily = "Sans-Serif"
  'objIE.Document.body.Style.FontFamily = "Comic Sans MS"

  objIE.Document.body.Style.BackgroundColor = "#C0C0C0"

  'add title, horizontal bar and textbox
  objIE.Document.Body.innerHTML = _
    "<center><b>Network Services Password Change</b></center>" & _
    "<hr></hr>" & _
    "<textarea READONLY name='TextArea' ID='TextArea' cols='90' rows='20' wrap='off'></textarea>" & _
    "<button name='Start' AccessKey='S' onclick=document.all('ButtonHandler').value='Start';><u>S</u>tart</button>" & _
    "<button name='Stop' AccessKey='T' onclick=document.all('ButtonHandler').value='Stop';>S<u>t</u>op</button>" & _
    "<button name='Exit' AccessKey='E' onclick=document.all('ButtonHandler').value='Exit';><u>E</u>xit</button>" & _
    "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'><br>" & _    
    "<font size='2' face='Comic Sans MS'><br>Processing: <label id='DeviceCurrent'>0</label> of <label id='DeviceTotal'>0</label> with <label id='DeviceLeft'>0</label> remaining" & _
    "<br>&nbsp;&nbsp;Passed: <label id='Passed'>0</label>" & _
    "<br>&nbsp;&nbsp;Failed: <label id='Failed'>0</label>" & _
    "<br>&nbsp;&nbsp;Not Supported: <label id='NotSupported'>0</label></font>"

  'configure Internet Explorer as desired
  objIE.MenuBar = False
  objIE.StatusBar = False
  objIE.AddressBar = False
  objIE.Toolbar = False
  objIE.height = 560
  objIE.width = 800  
  objIE.document.Title = "Password Change"
  objIE.Visible = True
  objIE.Resizable = False

  'disable the stop button
  objIE.Document.All("Stop").Disabled = "disabled"
End Sub

'*************************************************
'     OpenInventory : Open the inventory file
'*************************************************
Function OpenInventory
  Const adOpenStatic = 3
  Const adLockOptimistic = 3
  Const adCmdText = &H0001

  'create objects
  set fsoinvListConnection = CreateObject("ADODB.Connection")
  set fsoinvListRecordSet = CreateObject("ADODB.Recordset")

  'establish connection to the text file
  fsoinvListConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & sinvListLocation & ";" & _
    "Extended Properties=""text;HDR=YES;FMT=Delimited"""

  'read in the contents of the entire file
  on error resume next   
  fsoinvListRecordSet.Open "SELECT * FROM " & sinvListFileName, fsoinvListConnection, adOpenStatic, adLockOptimistic, adCmdText
  If Err.Number <> 0 Then
    Call crt.Dialog.MessageBox("Cannot open inventory file: "&sinvListLocation & sinvListFileName & ". Password Change will close.", "Inventory Error", 16)
    OpenInventory = "EXIT"
  End If

  on error goto 0
End Function

'*************************************************
'     CloseInventory : Close the inventory file
'*************************************************
Sub CloseInventory
  'close objects
  fsoinvListRecordSet.close
  fsoinvListConnection.close
End Sub

'*************************************************
'     Logon : Prompt for user logon credentials
'*************************************************
Sub Logon
  'prompt user for username
  slogonUsername = crt.Dialog.Prompt("Enter the username:", "Logon", "admin", False)

  'prompt user for current password and confirm
   Do
     'promt user for passwords
     slogonCurrentPassword1 = crt.Dialog.Prompt("Enter the current password:", "Logon", "", True)   
     slogonCurrentPassword2 = crt.Dialog.Prompt("Confirm the current password:", "Logon", "", True)

     'confirm that passwords match
     If slogonCurrentPassword1 <> slogonCurrentPassword2 Then
       'passwords do not match, warn the user and continue with the current password loop
       Call crt.Dialog.MessageBox("Current password does not match, try again.", "Logon Error", 32)
     Else
       'passwords match, save password and exit the current password loop
       slogonCurrentPassword = sLogonCurrentPassword1
       Exit Do
     End If
   Loop

  'prompt user for new password and confirm
   Do
     'prompt user for passwords
     slogonNewPassword1 = crt.Dialog.Prompt("Enter the new password:", "Logon", "", True)   
     slogonNewPassword2 = crt.Dialog.Prompt("Confirm the new password:", "Logon", "", True)

     'confirm that passwords match
     If slogonNewPassword1 <> slogonNewPassword2 Then
       'passwords do not match, warn the user and continue with the new password loop
       Call crt.Dialog.MessageBox("New password does not match, try again.", "Logon Error", 32)
     Else
       'passwords match, save password and exit the new password loop
       slogonNewPassword = slogonNewPassword1
       Exit Do
     End If
   Loop
End Sub

'*************************************************
'     GetInventoryRecord : Retrieve row from the inventory list
'*************************************************
Sub GetInventoryRecord
  'retrieve four fields from the row
  scurrentIPAddress = fsoinvListRecordSet.Fields.Item("IP Address")
  scurrentDeviceType = fsoinvListRecordSet.Fields.Item("Device Type")
  scurrentLocation = fsoinvListRecordSet.Fields.Item("Location")
  scurrentSystemName = fsoinvListRecordSet.Fields.Item("System Name")
  scurrentDescription = fsoinvListRecordSet.Fields.Item("Description")

  'move recordset object to the next row
  fsoinvListRecordSet.MoveNext
End Sub

'*************************************************
'     AddConsoleMessage : Add a message to the console
'*************************************************
Sub AddConsoleMessage(Message)
  currentText = objIE.Document.All("TextArea").Value
  objIE.Document.All("TextArea").Value = currentText & Message
End Sub

'*************************************************
'     AddLogMessage : Add a message to the log file
'*************************************************
'add a message to the log file
Sub AddLogMessage(Message)
  'open log file for appending
  set fsologFile = CreateObject("Scripting.FileSystemObject")
  set objwriteStuff = fsologFile.OpenTextFile(slogFile, 8, True)
 
  'write message without a CR and LR and close the file
  objwriteStuff.Write(Message)
  objwriteStuff.Close
End Sub

'*************************************************
'     onClickStart : Start button has been clicked
'*************************************************
Function onClickStart
  'clear button handler value
  objIE.Document.All("ButtonHandler").Value = ""

  'disable the start button
  objIE.Document.All("Start").Disabled = "disabled"

  'change focus from IE to SecureCRT
  crt.Window.Activate
  crt.Window.Show 1
  crt.Window.Show 2

  'prompt for user logon credentials
  call Logon

  'ask user if they are sure they want to continue
  result = crt.Dialog.MessageBox("Press OK to start", "Start", 0 Or 1 Or 32)
  If result = 1 Then
    'enable the stop button
    objIE.Document.All("Stop").Disabled = 0
    
    'call the main loop
    mainLoopResult = MainLoop

    'disable the stop button
    objIE.Document.All("Stop").Disabled = "disabled"

    'let user know the password change has finished
    Call crt.Dialog.MessageBox("Password change has finished.", "Password Change", 32)  
  End If

  'return the next action from the main loop
  onClickStart = mainLoopResult
End Function

'*************************************************
'     Main Loop : 
'
' Return Codes:
' ------------
' Not supported
' Password changed. (from device specific function)
' AuthenticateSSH2 function failure codes
'*************************************************
Function MainLoop
  'add message to the console and log file    
  AddConsoleMessage(Now & Chr(13) )  
  AddLogMessage(Now & chr(13) & Chr(10))
  AddConsoleMessage("Inventory [" & sinvListLocation & sinvListFileName & "]" & Chr(13) )  
  AddLogMessage("Inventory [" & sinvListLocation & sinvListFileName & "]" & chr(13) & Chr(10))
  AddConsoleMessage("Log [" & sLogFile & "]" & Chr(13) & Chr(13) )  
  AddLogMessage("Log [" & sLogFile & "]" & chr(13) & Chr(10) )

  'add logon credentials to log file
  AddLogMessage("Username: [" & slogonUsername & "]" & chr(13) & Chr(10) )
  AddLogMessage("Existing Password: [" & slogonCurrentPassword & "]" & chr(13) & Chr(10) )
  AddLogMessage("New Password: [" & slogonNewPassword & "]" & chr(13) & Chr(10) & chr(13) & Chr(10))

  'get the number of devices in inventory
  ideviceCount = fsoinvListRecordSet.RecordCount

  'setup the console counters
  idevicesLeft = ideviceCount
  objIE.Document.getElementById("DeviceTotal").innerHTML = ideviceCount 
  objIE.Document.getElementById("DeviceCurrent").innerHTML = 1
  objIE.Document.getElementById("DeviceLeft").innerHTML = idevicesLeft
  objIE.Document.getElementById("Passed").innerHTML = 0
  objIE.Document.getElementById("Failed").innerHTML = 0
  objIE.Document.getElementById("NotSupported").innerHTML = 0

  'retreive and process each row in the inventory list
  Do Until fsoinvListRecordSet.EOF
    'check for button clicks
    Select Case objIE.Document.All("ButtonHandler").Value
      Case "Stop"
        'return
        Exit Function
      Case "Exit"
        'indicate the application should exit and return
        MainLoop = "EXIT"
        Exit Function
    End Select

    'retreive inventory row
    Call GetInventoryRecord

    'update console counters
    ideviceCtr = ideviceCtr + 1
    idevicesLeft = idevicesLeft - 1
    objIE.Document.getElementById("DeviceCurrent").innerHTML = ideviceCtr
    objIE.Document.getElementById("DeviceLeft").innerHTML = idevicesLeft

    'add message to the console and log file    
    AddConsoleMessage("Device " & cStr(ideviceCtr) & " of "& ideviceCount & " : " & scurrentIPAddress & " : " & scurrentDeviceType & " : " &  scurrentsystemName )  
    AddLogMessage("Device " & cStr(ideviceCtr) & " of " & ideviceCount &" : " & scurrentIPAddress & " : " & scurrentDeviceType & " : " &  scurrentsystemName & chr(13) & Chr(10))

    'change password on device
    result = ChangePassword

    'process the return code
    Select Case result
      Case "Password Changed"

        'success, add message to the console and log file
        AddConsoleMessage(" : Password changed." & chr(13) )
        AddLogMessage(chr(13) & Chr(10) & "Password changed." & chr(13) & Chr(10) & chr(13) & Chr(10))
        idevicePassed = idevicePassed + 1
        objIE.Document.getElementById("Passed").innerHTML = idevicePassed

      Case "Device is not supported."
        'device is not supported, add message to the console and log file
        AddConsoleMessage(" : The device is not supported." & chr(13) )
        AddLogMessage(chr(13) & Chr(10) & "The device is not supported." & chr(13) & Chr(10) & chr(13) & Chr(10))
        ideviceNotSupported = ideviceNotSupported + 1
        objIE.Document.getElementById("NotSupported").innerHTML = ideviceNotSupported

      Case Else

        'error has occured, add message to the console and log fiile
        AddConsoleMessage(" : " & result & chr(13) )
        AddLogMessage(chr(13) & Chr(10) & result & chr(13) & Chr(10) & chr(13) & Chr(10))
        ideviceFailed = ideviceFailed + 1
        objIE.Document.getElementById("Failed").innerHTML = ideviceFailed

     End Select
  Loop
End Function

'*************************************************
'     ChangePassword : Change password on device
'
' Return Codes:
' ------------
' Password changed.
' Device not supported supported.
' AuthenticateSSH2 function failure codes
' IO_Matrix function failure codes
' IO_ER16 function failure codes
' IO_1H582_51 function failure codes
' IO_XSR3020 function failure codes
' IO_Cisco3020 function failure codes
'*************************************************
Function ChangePassword
  'check for non supported devices
  Select Case scurrentDeviceType
    Case "1H582-51"
    Case "Matrix N7 Platinum"
    Case "Matrix N5 Platinum"
    Case "Matrix X16"
    Case "ER-16"
    Case "XSR-3020"
    Case "1H582-25"
    Case "XSR-1805"
    Case "G3G124-24"
    Case "C3G124-24"
    Case "Cisco"
      If InStr(scurrentDescription,"CBS30X0") = 0 Then
        ChangePassword = "Device is not supported."
        Exit Function
      End If
    Case Else
      ChangePassword = "Device is not supported."
      Exit Function
  End Select

  'turn on SecureCRT logging in append mode
  crt.session.LogFileName = slogFile
  crt.session.log 1, 1

  'authenticate with device using SSH2
  result = AuthenticateSSH2

  'upon authentication failure return with error message
  If result <> "Connected" Then
    ChangePassword = result
  
    'turn SecureCRT logging off
    crt.session.log 0

    'disconnect session and return
    crt.Session.Disconnect

    Exit Function
  End If

  'set function to return success unless a failure occurs below
  ChangePassword = "Password Changed"

  'run the appropriate IO function for the current device
  Select Case scurrentDeviceType
 
    'run the IO function for select devices   
    Case "1H582-51", "1H582-25"
      'run the IO function
      result = IO_1H582_51()

      'upon failure change the function return code from success to the error message
      If result <> "Complete" Then ChangePassword = result

    'run the IO function for select devices
    Case "Matrix N7 Platinum", "Matrix N5 Platinum", "Matrix X16", "G3G124-24", "C3G124-24"
      'run the IO script
      result = IO_Matrix()

      'upon failure change the function return code from success to the error message
      If result <> "Complete" Then ChangePassword = result

    Case "XSR-1805", "XSR-3020"
      'run the IO function
      result = IO_XSR3020

      'upon failure change the function return code from success to the error message
      If result <> "Complete" Then ChangePassword = result

    Case "ER-16"
      'run the IO function
      result = IO_ER16

      'upon failure change the function return code from success to the error message
      If result <> "Complete" Then ChangePassword = result

    Case "Cisco"
      'run the IO function
      result = IO_Cisco3020

      'upon failure change the function return code from success to the error message
      If result <> "Complete" Then ChangePassword = result
 
   End Select

  'disconnect session
  crt.Session.Disconnect

  'turn SecureCRT logging off
  crt.session.log 0
End Function

'*************************************************
'     AuthenticateSSH2 : Authenticate with device using SSH2
'
' Return Codes:
' ------------
' Connected
' No IP connectivity or inactive SSH2 server.
' SSH2 authentication failed due to bad username or password.
' SSH2 authentication timed out.
' AuthenticateSSH2 function debug code error.
'
' Note:
' ----
' Upon successful return the prompt character is on the screen
'*************************************************
Function AuthenticateSSH2
  'connect to device
  on error resume next  
  crt.session.Connect "/SSH2 /ACCEPTHOSTKEYS " & scurrentIPAddress, False
  on error goto 0
  errCode = crt.GetLastError
  errMessage = crt.GetLastErrorMessage
  crt.ClearLastError

  'error connecting to the device, return with the error message
  if errCode <> 0 Then
    AuthenticateSSH2 = "No IP connectivity or inactive SSH2 server."
    Exit Function
  End If 

  'send username and password
  crt.Screen.WaitForString "sername:", 5
  crt.Screen.Send slogonUsername & Chr(13)
  crt.Screen.WaitForString "assword:", 5
  crt.Screen.Send slogonCurrentPassword  & Chr(13)

  'check authentication result
  result = crt.Screen.WaitForStrings(">","Password authentication failed.",10)
  Select Case result
    Case 1: 'found prompt, successfully authenticated
      AuthenticateSSH2 = "Connected"
      Exit Function
    Case 2: 'authentication failed due to bad username or password
      AuthenticateSSH2 = "SSH2 authentication failed due to bad username or password."
      Exit Function
    Case 0: 'authentication timed out
      AuthenticateSSH2 = "SSH2 authentication timed out."
      Exit Function
    Case Else: 'this should never occur
      AuthenticateSSH2 = "Authenticate function debug code error."
      Exit Function
  End Select
End Function

'*************************************************
' IO_Matrix : IO for Matrix N7 Platinum, Matrix X16, G3G124-24. C3G124-24
'
' Return Codes:
' ------------
' Complete
' IO_Matrix WaitForString #1 Timeout
' IO_Matrix WaitForString #2 Timeout
' IO_Matrix WaitForString #3 Timeout
' IO_Matrix WaitForString #4 Timeout
'*************************************************
Function IO_Matrix
  'execute set password command
  crt.Screen.Send "set password" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("enter old password:", 5) <> True Then
    IO_Matrix = "IO_Matrix WaitForString #1 Timeout"
    Exit Function
  End If
  
  'send username
  crt.Screen.Send slogonCurrentPassword & chr(13)

  'wait for next input prompt and return with error if timeout expires 
  If crt.Screen.WaitForString("enter new password:", 5) <> True Then
    IO_Matrix = "IO_Matrix WaitForString #2 Timeout"
    Exit Function
  End If
  
  'send password
  crt.Screen.Send slogonNewPassword & chr(13)
  
  'wait for last input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("re-enter new password:", 5) <> True Then
    IO_Matrix = "IO_Matrix WaitForString #3 Timeout"
    Exit Function
  End If
  
  'send password
  crt.Screen.Send slogonNewPassword & chr(13)

  'wait for prompt character and return with error if timeout expires
  If crt.Screen.WaitForString(">", 5) <> True Then
    IO_Matrix = "IO_Matrix WaitForString #4 Timeout"
    Exit Function
  End If

  'indicate success and exit
  IO_Matrix = "Complete"
End Function

'*************************************************
' IO_1H582_51 : IO for Enterasys Matrix E1 1H582_51
'
' Return Codes:
' ------------
' Complete
' IO_1H582_51 WaitForString #1 Timeout
' IO_1H582_51 WaitForString #2 Timeout
' IO_1H582_51 WaitForString #3 Timeout
' IO_1H582_51 WaitForString #4 Timeout
'*************************************************
Function IO_1H582_51
  'execute set password command
  crt.Screen.Send "set password" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString(" Old Password:", 5) <> True Then
    IO_1H582_51 = "IO_1H582_51 WaitForString #1 Timeout"
    Exit Function
  End If

  'send current password
  crt.Screen.Send slogonCurrentPassword & chr(13)

  'wait for next input prompt and return with error if timeout expires
  If crt.Screen.WaitForString(" New Password:", 5) <> True Then
    IO_1H582_51 = "IO_1H582_51 WaitForString #2 Timeout"
    Exit Function
  End If

  'send password
  crt.Screen.Send slogonNewPassword & chr(13)

  'wait for last input prompt and return with error if timeout expires
  If crt.Screen.WaitForString(" Retype New Password:", 5) <> True Then
    IO_1H582_51 = "IO_1H582_51 WaitForString #3 Timeout"
    Exit Function
  End If

  'send password
  crt.Screen.Send slogonNewPassword & chr(13)

  'wait for prompt character and return with error if timeout expires
  If crt.Screen.WaitForString(">", 5) <> True Then
    IO_1H582_51 = "IO_1H582_51 WaitForString #4 Timeout"
    Exit Function
  End If

  'indicate success and exit
  IO_1H582_51 = "Complete"
End Function

'*************************************************
' IO_ER16 : IO for Enterasys X-Pedition ER16
'
' Return Codes:
' ------------
' Complete
' IO_ER16 WaitForString #1 Timeout
' IO_ER16 WaitForString #2 Timeout
' IO_ER16 WaitForString #3 Timeout
' IO_ER16 WaitForString #4 Timeout
' IO_ER16 WaitForString #5 Timeout
' IO_ER16 WaitForString #6 Timeout
' IO_ER16 WaitForString #7 Timeout
' IO_ER16 WaitForString #8 Timeout
' IO_ER16 WaitForString #9 Timeout
'*************************************************
Function IO_ER16
  'execute enable command
  crt.Screen.Send "enable" & chr(13)

  'wait for priv exec mode prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #1 Timeout"
    Exit Function
  End If

  'execute configure command
  crt.Screen.Send "configure" & chr(13)

  'wait for priv exec mode prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #2 Timeout"
    Exit Function
  End If

  'execute set password command
  crt.Screen.Send "system set password login" & chr(13)

  'wait for new password prompt and return with error if timeout expires
  If crt.Screen.WaitForString("New Password:", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #3 Timeout"
    Exit Function
  End If

  'send password
  crt.Screen.Send slogonNewPassword & chr(13)

  'wait for verify password prompt and return with error if timeout expires
  If crt.Screen.WaitForString("Verify New Password:", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #4 Timeout"
    Exit Function
  End If

  'send password
  crt.Screen.Send slogonNewPassword & chr(13)

  'wait for priv exec mode prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #5 Timeout"
    Exit Function
  End If

  'exit priv exec mode
  crt.Screen.Send "exit" & chr(13)

  'wait for question and return with error if timeout expires
  If crt.Screen.WaitForString("Do you want to make the changes Active [yes]?", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #6 Timeout"
    Exit Function
  End If

  'answer yes to question
  crt.Screen.Send "yes" & chr(13)

  'wait for confirm message
  If crt.Screen.WaitForString("system login password has been changed", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #7 Timeout"
    Exit Function
  End If

  'save active configuration
  crt.Screen.Send "copy active to startup" & chr(13)

  'wait for question
  If crt.Screen.WaitForString("Are you sure you want to overwrite the Startup configuration [no]?", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #8 Timeout"
    Exit Function
  End If

  'answer yes to question
  crt.Screen.Send "yes" & chr(13)

  'wait for confirm message
  If crt.Screen.WaitForString("aFile copied successfully", 5) <> True Then
    IO_ER16 = "IO_ER16 WaitForString #9 Timeout"
    Exit Function
  End If

  'indicate success and exit
  IO_ER16 = "Complete"
End Function

'*************************************************
' IO_XSR3020 : IO for XSR3020
'
' Return Codes:
' ------------
' Complete
' IO_XSR3020 WaitForString #1 Timeout
' IO_XSR3020 WaitForString #2 Timeout
' IO_XSR3020 WaitForString #3 Timeout
' IO_XSR3020 WaitForString #4 Timeout
' IO_XSR3020 WaitForString #5 Timeout
' IO_XSR3020 WaitForString #6 Timeout
'*************************************************
Function IO_XSR3020 
  'enter priv exec mode
  crt.Screen.Send "enable" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_XSR3020  = "IO_XSR3020 WaitForString #1 Timeout"
    Exit Function
  End If

  'enter configuration mode
  crt.Screen.Send "configure" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_XSR3020  = "IO_XSR3020 WaitForString #2 Timeout"
    Exit Function
  End If

  'enter password change command
  crt.Screen.Send "username admin password secret 0 " & slogonNewPassword & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_XSR3020  = "IO_XSR3020 WaitForString #3 Timeout"
    Exit Function
  End If

  'exit configuration mode
  crt.Screen.Send "exit" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_XSR3020  = "IO_XSR3020 WaitForString #4 Timeout"
    Exit Function
  End If

  'copy running to startup
  crt.Screen.Send "copy running-config startup-config" & chr(13)

  'wait for question and return with error if timeout expires
  If crt.Screen.WaitForString("(y/n)", 5) <> True Then
    IO_XSR3020  = "IO_XSR3020 WaitForString #5 Timeout"
    Exit Function
  End If

  'send yes response
  crt.Screen.Send "y" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_XSR3020  = "IO_XSR3020 WaitForString #6 Timeout"
    Exit Function
  End If

  'indicate success and exit
  IO_XSR3020 = "Complete"
End Function

'*************************************************
' IO_Cisco3020 : IO for Cisco 3020
'
' Return Codes:
' ------------
' Complete
' IO_Cisco3020 WaitForString #1 Timeout
' IO_Cisco3020 WaitForString #2 Timeout
' IO_Cisco3020 WaitForString #3 Timeout
' IO_Cisco3020 WaitForString #4 Timeout
' IO_Cisco3020 WaitForString #5 Timeout
' IO_Cisco3020 WaitForString #6 Timeout
' IO_Cisco3020 WaitForString #7 Timeout
' IO_Cisco3020 WaitForString #8 Timeout
'*************************************************
Function IO_Cisco3020
  'enter priv exec mode
  crt.Screen.Send "enable" & chr(13)

  'wait for priv exec password prompt and return with error if timeout expires
  If crt.Screen.WaitForString("Password:", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #1 Timeout"
    Exit Function
  End If

  'send current password
  crt.Screen.Send slogonCurrentPassword & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #2 Timeout"
    Exit Function
  End If

  'enter configuration mode
  crt.Screen.Send "configure t" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #3 Timeout"
    Exit Function
  End If

  'change local account password
  crt.Screen.Send "username " & slogonUsername & " password " & slogonNewPassword & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #4 Timeout"
    Exit Function
  End If

  'change priv exec password
  crt.Screen.Send "enable secret " & slogonNewPassword & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #5 Timeout"
    Exit Function
  End If

  'exit configuration mode
  crt.Screen.Send "exit" & chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #6 Timeout"
    Exit Function
  End If

  'copy running to startup
  crt.Screen.Send "copy running-config startup-config" & chr(13)

  'wait for question and return with error if timeout expires
  If crt.Screen.WaitForString("Destination filename [startup-config]?", 5) <> True Then
    IO_XSR3020  = "IO_Cisco3020 WaitForString #7 Timeout"
    Exit Function
  End If

  'send enter as a response
  crt.Screen.Send chr(13)

  'wait for input prompt and return with error if timeout expires
  If crt.Screen.WaitForString("#", 5) <> True Then
    IO_Cisco3020  = "IO_Cisco3020 WaitForString #8 Timeout"
    Exit Function
  End If

  'indicate success and exit
  IO_Cisco3020 = "Complete"
End Function