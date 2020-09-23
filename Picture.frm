VERSION 5.00
Begin VB.Form PictureForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Grapher"
   ClientHeight    =   5070
   ClientLeft      =   720
   ClientTop       =   840
   ClientWidth     =   6900
   Icon            =   "Picture.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   5070
   ScaleWidth      =   6900
   Begin VB.Timer main 
      Interval        =   100
      Left            =   2880
      Top             =   2280
   End
End
Attribute VB_Name = "PictureForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim value(1) As String
Dim vars(1) As String
Dim equation As String, equ As String
Dim Y As Single
Dim cleared As Boolean

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then
    If cleared = False Then
        PictureForm.Cls
        cleared = True
        Exit Sub
    End If
    Exit Sub
End If

If Graphed = True Then
    If cleared Then cleared = False
    
    equation = traceForm.lstEqu.List(traceForm.lstEqu.ListIndex)
    equ = equation
    value(0) = Val(Evaluate(traceForm.txtX.Text, False))
    vars(0) = "x"
    Y = Round(Val(Evaluate(equ, False, vars, value)), 3)
    traceForm.txtY.Text = Y
    
    PictureForm.AutoRedraw = False
    traceForm.txtX.Text = Round(X, 3)
    PictureForm.Cls
    PictureForm.DrawStyle = 2
    PictureForm.Line (X, YMax)-(X, YMin), QBColor(0)
    If ((Y < Maximum) And (Y > -Maximum)) Then PictureForm.Line (XMax, Y)-(XMin, Y), QBColor(0)
    PictureForm.DrawStyle = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
