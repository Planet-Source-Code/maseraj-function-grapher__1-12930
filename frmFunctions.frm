VERSION 5.00
Begin VB.Form frmFunctions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Functions"
   ClientHeight    =   4710
   ClientLeft      =   10530
   ClientTop       =   4245
   ClientWidth     =   2655
   Icon            =   "frmFunctions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   2655
   Begin VB.Frame Frame1 
      Caption         =   "Functions (Double-Click to Insert)"
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.ListBox List1 
         Height          =   4350
         ItemData        =   "frmFunctions.frx":0442
         Left            =   120
         List            =   "frmFunctions.frx":04DF
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_DblClick()
frmGrapher.cboEquation.Text = frmGrapher.cboEquation.Text & List1.List(List1.ListIndex)
End Sub
