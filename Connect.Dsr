VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} ConnectionString 
   ClientHeight    =   11460
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10620
   _ExtentX        =   18733
   _ExtentY        =   20214
   _Version        =   393216
   Description     =   "Creates a connection string and saves it to clipboard."
   DisplayName     =   "ADO Connection String Builder"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "ConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VBInstance               As VBIDE.VBE
Dim mcbMenuCommandBar           As Office.CommandBarControl
Dim mfrmConnectionString        As New frmConnectionString
Public WithEvents MenuHandler   As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Sub Hide()
    On Error Resume Next
    mfrmConnectionString.Hide
End Sub

Sub Show()
Dim strConn As String
Dim strMess As String

    On Error Resume Next
    
    If mfrmConnectionString Is Nothing Then
        Set mfrmConnectionString = New frmConnectionString
    End If
    
    Set mfrmConnectionString.VBInstance = VBInstance
    Set mfrmConnectionString.Connect = Me

    strConn = BuildConnectionString
    
    If strConn = "" Then
        MsgBox "A connection string could not be established.", , "Connection Error"
    Else
        'strConn = FormatConnectionString(strConn)   '<-- Used to return the conncetion sting on multiple lines
        strMess = "The following connection string has been placed on clipboard."
        strMess = strMess & vbCrLf & vbCrLf
        strMess = strMess & Chr(34) & strConn & Chr(34)
        'strMess = strMess & strConn '<-- Used to return the conncetion sting on multiple lines
        
        MsgBox strMess, , "Connection String"
        
        Clipboard.Clear
        Clipboard.SetText Chr(34) & strConn & Chr(34)
        'Clipboard.SetText strConn   '<-- Used to return the conncetion sting on multiple lines
    End If
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    '   save the vb instance
    Set VBInstance = Application
    
    '   this is a good place to set a breakpoint and
    '   test various addin objects, properties and methods
    '   Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        '   Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("ADO Connection String Builder")
        '   sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    Unload mfrmConnectionString
    Set mfrmConnectionString = Nothing

End Sub

'   this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    '   see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        '   not available so we fail
        Exit Function
    End If
    
    '   add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    '   set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
End Function

Private Function BuildConnectionString() As String
Dim objLink As MSDASC.DataLinks
Dim strConn As String

    On Error Resume Next
    
    '   Display the dialog
    Set objLink = New MSDASC.DataLinks
    strConn = objLink.PromptNew
    
    If Err.Number = 0 Then
        '   Create a Connection object on this connection string
        BuildConnectionString = strConn
    Else
        '   User canceled the operation
        BuildConnectionString = ""
    End If
    
    Set objLink = Nothing
End Function

Private Function FormatConnectionString(ByVal strConnect As String) As String
'   Returns the connection string on multiple lines
Dim strParts() As String
Dim strBuild As String
Dim i As Integer

    strParts = Split(strConnect, ";")
    
    For i = 0 To UBound(strParts) - 1
        strParts(i) = Chr(34) & strParts(i) & Chr(59) & Chr(34) & " & _"
    Next i
    
    For i = 0 To UBound(strParts) - 1
        strBuild = strBuild & strParts(i) & vbCrLf
    Next i
    
    strBuild = strBuild & Chr(34) & strParts(UBound(strParts)) & Chr(34)
    
    FormatConnectionString = strBuild
End Function
