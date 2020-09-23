VERSION 5.00
Begin VB.Form frmGrapher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function Grapher            E-mail      Wolf at wolf_knight_x@yahoo.com"
   ClientHeight    =   2910
   ClientLeft      =   705
   ClientTop       =   7245
   ClientWidth     =   8910
   FillStyle       =   0  'Solid
   Icon            =   "Function Grapher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   8910
   Begin VB.ListBox lstEquations 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      ItemData        =   "Function Grapher.frx":0442
      Left            =   240
      List            =   "Function Grapher.frx":0444
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Functions (Double-Click to delete)"
      Height          =   1455
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Functions"
      Height          =   495
      Left            =   3600
      TabIndex        =   33
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3600
      Picture         =   "Function Grapher.frx":0446
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   1800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboEquation 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Function Grapher.frx":0888
      Left            =   600
      List            =   "Function Grapher.frx":08B9
      TabIndex        =   12
      Top             =   240
      Width           =   3495
   End
   Begin VB.Frame fraSize 
      Caption         =   "Options"
      Height          =   2895
      Left            =   4920
      TabIndex        =   20
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtXScale 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtXMax 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtStep 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtXMin 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtYScale 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkRandom 
         Caption         =   "Randomize Graph Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Use Limited Range"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkAxes 
         Caption         =   "Draw Axes"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cboSimult 
         Height          =   315
         ItemData        =   "Function Grapher.frx":093A
         Left            =   2520
         List            =   "Function Grapher.frx":0944
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox cboLine 
         Height          =   315
         ItemData        =   "Function Grapher.frx":0963
         Left            =   2520
         List            =   "Function Grapher.frx":096D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtYMax 
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtYMin 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "X"
         Height          =   1335
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         Begin VB.Label Label4 
            Caption         =   "Scale"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Min"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Max"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Y"
         Height          =   1335
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1815
         Begin VB.Label Label6 
            Caption         =   "Scale"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Min"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Max"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Step (Twips)"
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame fraFormula 
      Caption         =   "Formula"
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtEquation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Text            =   "5*exp(-(x^2+x^2)/5)"
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label lblYEqu 
         Alignment       =   2  'Center
         Caption         =   "Y="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdgraph 
      Caption         =   "Graph"
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   8955
      TabIndex        =   16
      Top             =   3120
      Width           =   9015
   End
End
Attribute VB_Name = "frmGrapher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim value(1) As String
Dim vars(1) As String
Dim expression As String
Dim X As Single
Dim Y1 As Single
Dim Y2 As Single
Dim Scales
Dim Red As Double, Green As Double, Blue As Double
Dim times As Double

Private Sub cboEquation_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdAdd_Click
End Sub

Private Sub cmdAdd_Click()
lstEquations.AddItem (cboEquation.Text)
End Sub

Private Sub cmdGraph_Click()
On Error Resume Next

Dim i
traceForm.lstEqu.Clear

For i = 0 To lstEquations.ListCount - 1
    traceForm.lstEqu.AddItem lstEquations.List(i)
Next

PictureForm.AutoRedraw = True
PictureForm.Cls

times = Time * Date

XMin = Evaluate(txtXMin.Text, False)
XMax = Evaluate(txtXMax.Text, False)
YMin = Evaluate(txtYMin.Text, False)
YMax = Evaluate(txtYMax.Text, False)
XScale = Evaluate(txtXScale.Text, False)
YScale = Evaluate(txtYScale.Text, False)
Step = (Abs(XMin) + Abs(XMax)) / PictureForm.Width * Evaluate(txtStep.Text, False)

PictureForm.Scale (XMin, YMax)-(XMax, YMin)

If chkAxes.value = 1 Then Call scaled

frmGrapher.MousePointer = 11
i = -1

While (i < lstEquations.ListCount - 1)

Randomize
Red = Int(255 * Rnd())
Green = Int(255 * Rnd())
Blue = Int(255 * Rnd())

i = i + 1
If ((i = 1) And (cboSimult.Text = "Simultaneous")) Then
    frmGrapher.MousePointer = 0
    If chkAxes.value = 1 Then Call scaled
    traceForm.lstEqu.Selected(0) = True
    Exit Sub
End If

For X = XMin To XMax + Step Step Step
        
'q = (i Mod (lstEquations.ListCount + 1)) - 1
'q = i
expression = lstEquations.List(i)

If cboSimult.Text = "Simultaneous" Then
    For a = 0 To lstEquations.ListCount - 1
        'Randomize (times + a)
        Red = Int(255 * Rnd(-1 * (times + a)))
        Green = Int(255 * Rnd(-1 * (times + a) + Red))
        Blue = Int(255 * Rnd(-1 * (times + a) + Green + Red))
        
        If chkRandom.value <> 1 Then
            PictureForm.ForeColor = QBColor((a + 1) Mod 15)
        Else
            PictureForm.ForeColor = RGB(Red, Green, Blue)
        End If
        Call graph(lstEquations.List(a), X, Step)
    Next
Else
    If chkRandom.value <> 1 Then
        PictureForm.ForeColor = QBColor((i + 1) Mod 15)
    Else
        PictureForm.ForeColor = RGB(Red, Green, Blue)
    End If
    Call graph(expression, X, Step)
End If

Next

Wend
frmGrapher.MousePointer = 0
traceForm.lstEqu.Selected(0) = True
If chkAxes.value = 1 Then Call scaled
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReset_Click()
txtXMax.Text = "3*pi"
txtStep.Text = "50"
txtXMin.Text = "-2*pi/3"
txtXScale.Text = "pi/2"
txtYMax.Text = "4"
txtYMin.Text = "-4"
txtYScale.Text = "1"
cboEquation.Text = "sin(x)"
cboLine.Text = "Line"
cboSimult.Text = "Consecutive"
chkAxes.value = 1
chkLimit.value = 1
chkLimit.value = 1
lstEquations.Clear
lstEquations.AddItem "sin(x)"
lstEquations.AddItem "cos(x)"

XMin = Evaluate(txtXMin.Text, False)
XMax = Evaluate(txtXMax.Text, False)
YMin = Evaluate(txtYMin.Text, False)
YMax = Evaluate(txtYMax.Text, False)
XScale = Evaluate(txtXScale.Text, False)
YScale = Evaluate(txtYScale.Text, False)
Step = (Abs(XMin) + Abs(XMax)) / PictureForm.Width * Evaluate(txtStep.Text, False)

PictureForm.Scale (XMin, YMax)-(XMax, YMin)

traceForm.txtX.Text = 0
traceForm.txtY.Text = 0

End Sub

Private Sub Command1_Click()
lstEquations.Clear
traceForm.lstEqu.Clear
End Sub

Private Sub Command2_Click()
frmFunctions.Visible = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdAdd_Click
End Sub

Private Sub Form_Load()
Graphed = False
cmdReset_Click
PictureForm.Visible = True
traceForm.Visible = True
frmFunctions.Visible = True
Show
PictureForm.Show
traceForm.Show
Call scaled
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstEquations_Click()
cboEquation.Text = lstEquations.List(lstEquations.ListIndex)
End Sub

Private Sub lstEquations_DblClick()
lstEquations.RemoveItem lstEquations.ListIndex
End Sub

Function scaled()
PictureForm.AutoRedraw = True
PictureForm.ForeColor = vbBlack
PictureForm.Line (0, YMax)-(0, YMin)
PictureForm.Line (XMax, 0)-(XMin, 0)

If XScale > 0 And YScale > 0 Then
    
    Scales = (Abs(YMin) + Abs(YMax)) / PictureForm.Height * 50
    For i = 0 To XMax Step XScale
        PictureForm.Line (i, Scales)-(i, -(Scales))
    Next
    For i = 0 To XMin Step -XScale
        PictureForm.Line (i, Scales)-(i, -(Scales))
    Next

    Scales = (Abs(XMin) + Abs(XMax)) / PictureForm.Width * 50
    For i = 0 To YMax Step YScale
        PictureForm.Line (Scales, i)-(-(Scales), i)
    Next
    For i = 0 To YMin Step -YScale
        PictureForm.Line (Scales, i)-(-(Scales), i)
    Next
        
End If
PictureForm.AutoRedraw = False
End Function

Function graph(expression As String, X, Step)
Maximum = (Abs(YMax) + Abs(YMin)) * 1.5
Graphed = True
PictureForm.AutoRedraw = True
value(0) = X
vars(0) = "x"
Y1 = Val(Evaluate(expression, False, vars, value))
    
If ((Y1 < Maximum) And (Y1 > -Maximum)) Or chkLimit.value = 0 Then
    If cboLine.Text = "Line" Then
        value(0) = X - Step
        Y2 = Val(Evaluate(expression, False, vars, value))
        If Not (((Y2 < YMin) And (Y2 > YMax)) Or ((Y1 < YMin) And (Y1 > YMax))) Or chkLimit.value = 0 Then PictureForm.Line (X, Y1)-(X - Step, Y2)
    Else
        PictureForm.PSet (X, Y1)
    End If
End If
PictureForm.AutoRedraw = False
End Function
