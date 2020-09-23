VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6825
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin MSFlexGridLib.MSFlexGrid ag 
      Height          =   2535
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "down"
      Height          =   1095
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "up"
      Height          =   1095
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' setup odbc for testkkk.mdb
' dsn name = testkkk



Public db As ADODB.Connection
Public rs As ADODB.Recordset
Public mcounter As Integer
Public keycounter As Integer
Public reach_line As Integer


Private Sub Command1_Click()

If (ag.TopRow - reach_line) > 0 Then
ag.TopRow = ag.TopRow - reach_line
Else
ag.TopRow = 0
End If
End Sub

Private Sub Command2_Click()

If (ag.TopRow + reach_line) < keycounter Then
ag.TopRow = ag.TopRow + reach_line
End If
End Sub

Private Sub Form_Load()
' keycounter is counter of total records in database
' setup number of line for each page
reach_line = 12
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.ConnectionString = "dsn=testkkk"
db.Open
rs.Open "testkkk", db, adOpenStatic
rs.MoveFirst
ag.Rows = 0
ag.Clear
Do Until rs.EOF
ag.AddItem rs.Fields(1).Value & vbTab & rs.Fields(2).Value
rs.MoveNext
Loop
keycounter = rs.RecordCount
rs.Close
db.Close
End Sub
