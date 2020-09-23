VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Exchange"
   ClientHeight    =   4800
   ClientLeft      =   5040
   ClientTop       =   2880
   ClientWidth     =   5685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5685
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4545
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9525
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Print Current Database"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   3960
      Width           =   3375
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Import, Store, and Print Information"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3480
      Width           =   3375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Import && Store Information"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "View Information Only"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Database"
      Height          =   1815
      Left            =   735
      TabIndex        =   2
      Top             =   600
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Email Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2295
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Exchange Server Name:"
      Height          =   255
      Left            =   375
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim MSAccess As Access.Application

Private Sub Combo1_Click()
On Error Resume Next
    rs.FindFirst "[ID] = " & Combo1.ItemData(Combo1.ListIndex)
    
    Text2.Text = rs.Fields("Email") & ""
End Sub

Private Sub Command1_Click()
On Error GoTo err
Dim oConnection As Object
Dim oRecordset As Object
Dim oCommand As Object
Dim x As Long

If Text1.Text = "" Then
MsgBox "You must enter a Exchange Server Name."
Exit Sub
End If

If Option1.Value = True Then
frmViewOnly.Text1.Text = frmMain.Text1.Text
frmViewOnly.Show vbModal

Exit Sub
End If

Screen.MousePointer = vbHourglass
Command1.Enabled = False
Command2.Enabled = False

If Option2.Value = True Then
StatusBar1.Panels(1).Text = "Status: Clearing Current Database..."
DoEvents
db.Execute "DELETE * FROM Master"

StatusBar1.Panels(1).Text = "Status: Connecting to Exchange Server..."

x = 0

Set oConnection = CreateObject("ADODB.Connection")
Set oRecordset = CreateObject("ADODB.Recordset")
Set oCommand = CreateObject("ADODB.Command")

Combo1.Clear
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


If Trim(oRecordset.Fields("cn").Value) & "" <> "" Then
rs.AddNew
DoEvents
x = x + 1
StatusBar1.Panels(1).Text = "Status: Storing New Entry " & x

    rs.Fields("Name") = "" & oRecordset.Fields("cn")
    rs.Fields("Email") = "" & oRecordset.Fields("mail")
    rs.Update
End If


   oRecordset.MoveNext
 
Wend
StatusBar1.Panels(1).Text = "Status: Disconnecting..."
DoEvents
StatusBar1.Panels(1).Text = "Status: Reloading Current Database..."
Call Form_Load
DoEvents
x = 0
StatusBar1.Panels(1).Text = "Status: Done"
End If

If Option3.Value = True Then
StatusBar1.Panels(1).Text = "Status: Clearing Current Database..."

db.Execute "DELETE * FROM Master"
StatusBar1.Panels(1).Text = "Status: Connecting to Exchange Server..."

x = 0

Set oConnection = CreateObject("ADODB.Connection")
Set oRecordset = CreateObject("ADODB.Recordset")
Set oCommand = CreateObject("ADODB.Command")

Combo1.Clear
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

DoEvents

If Trim(oRecordset.Fields("cn").Value) & "" <> "" Then
    x = x + 1
    StatusBar1.Panels(1).Text = "Status: Storing New Entry " & x
    rs.AddNew



    rs.Fields("Name") = "" & oRecordset.Fields("cn")
    rs.Fields("Email") = "" & oRecordset.Fields("mail")


DoEvents


rs.Update
    End If
DoEvents
   oRecordset.MoveNext
Wend
StatusBar1.Panels(1).Text = "Status: Disconnecting..."
DoEvents
StatusBar1.Panels(1).Text = "Status: Reloading Current Database..."
Call Form_Load
DoEvents
x = 0
DoEvents

StatusBar1.Panels(1).Text = "Status: Printing..."
Set MSAccess = New Access.Application

MSAccess.OpenCurrentDatabase (App.Path & "\Database.mdb")

MSAccess.DoCmd.OpenReport "Master Report", acViewNormal
MSAccess.CloseCurrentDatabase
Set MSAccess = Nothing
DoEvents
StatusBar1.Panels(1).Text = "Status: Done"
End If

If Option4.Value = True Then
StatusBar1.Panels(1).Text = "Status: Printing..."
Set MSAccess = New Access.Application

MSAccess.OpenCurrentDatabase (App.Path & "\Database.mdb")

MSAccess.DoCmd.OpenReport "Master Report", acViewNormal
MSAccess.CloseCurrentDatabase
Set MSAccess = Nothing
StatusBar1.Panels(1).Text = "Status: Done..."
DoEvents
End If

Set oRecordset = Nothing
Set oCommand = Nothing
Set oConnection = Nothing

Screen.MousePointer = vbDefault
Command1.Enabled = True
Command2.Enabled = True

Exit Sub

err:

Screen.MousePointer = vbDefault


MsgBox Error
Command1.Enabled = True
Command2.Enabled = True

Me.Enabled = True
End Sub

Private Sub Command2_Click()


Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

    Set db = OpenDatabase(App.Path & "\Database.mdb")

    Set rs = db.OpenRecordset("SELECT * FROM Master " & "ORDER BY [Name]")
 
    ' Populate the list box
    Do Until rs.EOF
        If rs.Fields("Name") & "" <> "" Then
            Combo1.AddItem rs.Fields("Name")
            Combo1.ItemData(Combo1.NewIndex) = rs.Fields("ID")
        End If
        rs.MoveNext
        DoEvents
    Loop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Command1.Enabled = False Then Cancel = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rs.Close
    db.Close
End Sub
