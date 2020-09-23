VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DAO/JET Database driver test form, protected database"
   ClientHeight    =   1800
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5448
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5448
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Last"
      Height          =   528
      Left            =   4236
      TabIndex        =   4
      Top             =   1224
      Width           =   1152
   End
   Begin VB.CommandButton Command1 
      Caption         =   "First"
      Height          =   528
      Left            =   36
      TabIndex        =   3
      Top             =   1224
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   792
      Left            =   36
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   372
      Width           =   5352
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DAO version:"
      Height          =   192
      Left            =   108
      TabIndex        =   2
      Top             =   60
      Width           =   948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   192
      Left            =   1128
      TabIndex        =   1
      Top             =   72
      Width           =   456
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MYTB.MoveFirst
Form1.Text1.Text = MYTB.Fields("Test")
End Sub

Private Sub Command2_Click()
MYTB.MoveLast
Form1.Text1.Text = MYTB.Fields("Test")
End Sub

Private Sub Form_Load()
Dim x&
x = OpenMYDB
If x <> -1 Then
MsgBox "There was error num:" + Str$(x) + " while opening the database!"
Exit Sub
End If
MYTB.MoveFirst
Text1.Text = MYTB.Fields("Test")
Label1.Caption = MyData.Version
End Sub

Private Sub Form_Unload(Cancel As Integer)
MYTB.Close
MYDB.Close
MYWR.Close
Set MYTB = Nothing
Set MYDB = Nothing
Set MYWR = Nothing
Set MyData = Nothing
Set Form1 = Nothing 'VB on SP5 level need's this
End Sub
