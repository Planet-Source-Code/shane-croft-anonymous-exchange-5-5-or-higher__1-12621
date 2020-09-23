VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewOnly 
   Caption         =   "View Only"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "frmViewOnly.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3780
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4683
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   7421
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Email"
         Object.Width           =   7233
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Exchange Server Name:"
      Height          =   255
      Left            =   1860
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   3495
   End
End
Attribute VB_Name = "frmViewOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
FormBusy
Dim oConnection As Object
Dim oRecordset As Object
Dim oCommand As Object
Dim sMail As String

Set oConnection = CreateObject("ADODB.Connection")
Set oRecordset = CreateObject("ADODB.Recordset")
Set oCommand = CreateObject("ADODB.Command")

ListView1.ListItems.Clear
oConnection.Provider = "ADsDSOObject"  'The ADSI OLE-DB provider
oConnection.Properties("User ID") = ""
oConnection.Properties("Password") = ""
oConnection.Properties("Encrypt Password") = False
oConnection.Open "ADs Provider"

strQuery = "<GC://" & Text1.Text & ":389>;(objectClass=*);cn,mail;subtree"

oCommand.ActiveConnection = oConnection
oCommand.CommandText = strQuery
oCommand.Properties("Page Size") = 99
Set oRecordset = oCommand.Execute

While Not oRecordset.EOF
    If FixNull(oRecordset.Fields("cn").Value) <> "" Then
        Set Item = ListView1.ListItems.Add(, , FixNull(oRecordset.Fields("cn").Value))
        sMail = FixNull(oRecordset.Fields("mail").Value)
        Item.SubItems(1) = sMail
    End If
   oRecordset.MoveNext
   DoEvents
   
   
Wend

CleanUp:
Set oRecordset = Nothing
Set oCommand = Nothing
Set oConnection = Nothing
FormReady
Exit Sub

err:
Screen.MousePointer = vbDefault
MsgBox Error

GoTo CleanUp

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'don't permit unload while busy
If Command1.Enabled = False Then Cancel = True

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total: " & ListView1.ListItems.Count
End Sub
Private Sub FormBusy()
Command1.Enabled = False
Screen.MousePointer = vbHourglass
End Sub
Private Sub FormReady()
Command1.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Public Function FixNull(vMayBeNull As Variant) As String
   On Error Resume Next
   FixNull = vbNullString & vMayBeNull
End Function
