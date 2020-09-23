VERSION 5.00
Begin VB.Form traceForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracer"
   ClientHeight    =   1875
   ClientLeft      =   8865
   ClientTop       =   660
   ClientWidth     =   5670
   Icon            =   "traceForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5670
   Begin VB.CommandButton cmdTrace 
      Caption         =   "Trace"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Coordinates"
      Height          =   975
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "e.g.  pi/2, 0"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Y"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Equation to Trace"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.ListBox lstEqu 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "traceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim value(1) As String
Dim vars(1) As String
Dim eqation As String
Dim equ As String
Dim Y As Single, XE As Double

Private Sub cmdTrace_Click()
On Error Resume Next
equation = lstEqu.List(lstEqu.ListIndex)
equ = equation
value(0) = Val(Evaluate(txtX.Text, False))
vars(0) = "x"
Y = Round(Val(Evaluate(equ, False, vars, value)), 3)
traceForm.txtY.Text = Y
End Sub

Private Sub cmdTrace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
equation = lstEqu.List(lstEqu.ListIndex)
equ = equation
XE = Val(Evaluate(txtX.Text, False))
value(0) = XE
vars(0) = "x"
Y = Round(Val(Evaluate(equ, False, vars, value)), 3)

PictureForm.DrawStyle = 2
PictureForm.Line (XE, YMax)-(XE, YMin), QBColor(0)
If ((Y < Maximum) And (Y > -Maximum)) Then PictureForm.Line (XMax, Y)-(XMin, Y), QBColor(0)
PictureForm.DrawStyle = 0

End Sub

Private Sub cmdTrace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureForm.Cls
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
