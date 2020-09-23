VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   2880
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9480
      TabIndex        =   3
      Top             =   3735
      Width           =   9480
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   5
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9480
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gradiant Picture"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   1
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gradiant Picture"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Top             =   -15
         Width           =   3495
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Vote Me"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   7
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gradiant Form"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, Green, Blue
Public Enum enumOrientation
    Orientation_Horizontal = 0
    Orientation_Vertical = 1
End Enum

Public Function Gradient(Frm As Object, Orientation As enumOrientation, SClr As ColorConstants, EClr As ColorConstants)
Frm.AutoRedraw = True: Frm.ScaleMode = 3 '2 is interesting,too
Analyze (SClr): SRed = Red: SGreen = Green: SBlue = Blue
Analyze (EClr): ERed = Red: EGreen = Green: EBlue = Blue
DifR = ERed - SRed: DifG = EGreen - SGreen: DifB = EBlue - SBlue
Select Case Orientation
  Case Is = 0: Fora = Frm.ScaleHeight
  Case Is = 1: Fora = Frm.ScaleWidth
End Select
For Yi = 0 To Fora
SRed = SRed + (DifR / Fora): If SRed < 0 Then SRed = 0
SGreen = SGreen + (DifG / Fora): If SGreen < 0 Then SGreen = 0
SBlue = SBlue + (DifB / Fora): If SBlue < 0 Then SBlue = 0
Select Case Orientation
  Case Is = 0: Frm.Line (0, Yi)-(Frm.ScaleWidth, Yi), RGB(SRed, SGreen, SBlue), B
  Case Is = 1: Frm.Line (Yi, 0)-(Yi, Frm.ScaleHeight), RGB(SRed, SGreen, SBlue), B
End Select
Next
End Function

Public Function Analyze(CConst As ColorConstants)
Dim rr, gr, br As Long
rr = 1: gr = 256: br = 65536
Dim rest As Long
rest = CConst \ br
Blue = rest
CConst = CConst Mod br
If Blue < 0 Then Blue = 0
rest = CConst \ gr
Green = rest
CConst = CConst Mod gr
If Green < 0 Then Green = 0
rest = CConst \ rr
Red = rest
CConst = CConst Mod rr
If Red < 0 Then Red = 0
End Function

Private Sub Form_Resize()
Gradient Picture1, 1, vbCyan, vbYellow
Gradient Picture2, 0, vbRed, vbCyan
Gradient Me, 1, vbGreen, vbRed
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
Label5.Caption = Format(Date, "Long date")
End Sub
