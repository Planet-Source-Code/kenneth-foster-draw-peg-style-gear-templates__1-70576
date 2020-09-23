VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw Peg Type Gears"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   11850
      TabIndex        =   26
      Top             =   7290
      Width           =   1125
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Point of Contact  - End only"
      Height          =   300
      Left            =   10455
      TabIndex        =   25
      Top             =   3120
      Width           =   2565
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Print Teeth and Ratio Labels on gear. If wheel size < 1 then no labels will show."
      Height          =   630
      Left            =   10455
      TabIndex        =   23
      Top             =   3465
      Width           =   2595
   End
   Begin VB.CommandButton cmdStartGear 
      Caption         =   "Start another gear"
      Height          =   435
      Left            =   10710
      TabIndex        =   18
      Top             =   6000
      Width           =   2100
   End
   Begin VB.CommandButton cmdSetPos 
      Caption         =   "Set Position"
      Height          =   435
      Left            =   10710
      TabIndex        =   14
      Top             =   5400
      Width           =   2100
   End
   Begin VB.VScrollBar VS2 
      Height          =   9750
      LargeChange     =   10
      Left            =   9810
      Max             =   9000
      Min             =   500
      SmallChange     =   5
      TabIndex        =   13
      Top             =   45
      Value           =   1000
      Width           =   405
   End
   Begin VB.HScrollBar HS2 
      Height          =   375
      LargeChange     =   10
      Left            =   30
      Max             =   9000
      Min             =   500
      SmallChange     =   5
      TabIndex        =   12
      Top             =   9780
      Value           =   1000
      Width           =   9780
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show End Peg Marks"
      Height          =   360
      Left            =   10455
      TabIndex        =   8
      Top             =   2775
      Width           =   1920
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Draw Change"
      Height          =   405
      Left            =   10905
      TabIndex        =   7
      Top             =   2340
      Width           =   1470
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   435
      Left            =   10440
      TabIndex        =   5
      Top             =   7290
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9480
      Left            =   285
      ScaleHeight     =   9420
      ScaleWidth      =   9420
      TabIndex        =   4
      Top             =   270
      Width           =   9480
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   11385
      TabIndex        =   1
      Text            =   "60"
      Top             =   2040
      Width           =   555
   End
   Begin VB.HScrollBar HS1 
      Height          =   270
      Left            =   10830
      Max             =   4300
      Min             =   300
      TabIndex        =   0
      Top             =   870
      Value           =   2000
      Width           =   1605
   End
   Begin VB.Label Label14 
      Caption         =   "There is no erase or undo function."
      Height          =   270
      Left            =   10305
      TabIndex        =   24
      Top             =   8715
      Width           =   2655
   End
   Begin VB.Label Label13 
      Caption         =   "Position Gear with Scrollbars."
      Height          =   210
      Left            =   10785
      TabIndex        =   22
      Top             =   5145
      Width           =   2100
   End
   Begin VB.Line Line3 
      X1              =   10410
      X2              =   13020
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label Label12 
      Caption         =   "STEP 3"
      Height          =   210
      Left            =   10425
      TabIndex        =   21
      Top             =   6825
      Width           =   675
   End
   Begin VB.Line Line2 
      X1              =   10395
      X2              =   12975
      Y1              =   5085
      Y2              =   5085
   End
   Begin VB.Label Label11 
      Caption         =   "STEP 2"
      Height          =   210
      Left            =   10425
      TabIndex        =   20
      Top             =   4800
      Width           =   690
   End
   Begin VB.Line Line1 
      X1              =   10395
      X2              =   13020
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label10 
      Caption         =   "STEP 1"
      Height          =   210
      Left            =   10455
      TabIndex        =   19
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape1 
      Height          =   930
      Left            =   10455
      Top             =   9075
      Width           =   2550
   End
   Begin VB.Label Label9 
      Caption         =   " Ratio Between Marks---"
      Height          =   360
      Left            =   10530
      TabIndex        =   17
      Top             =   9690
      Width           =   2475
   End
   Begin VB.Label Label8 
      Caption         =   " Number of Teeth---"
      Height          =   345
      Left            =   10515
      TabIndex        =   16
      Top             =   9420
      Width           =   1875
   End
   Begin VB.Label Label7 
      Caption         =   " Wheel Size---"
      Height          =   240
      Left            =   10530
      TabIndex        =   15
      Top             =   9165
      Width           =   1860
   End
   Begin VB.Label Label6 
      Caption         =   "Approx. Inches"
      Height          =   240
      Left            =   11190
      TabIndex        =   11
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Except for the Number of Teeth, all numbers are relative."
      Height          =   450
      Left            =   10320
      TabIndex        =   10
      Top             =   8220
      Width           =   2745
   End
   Begin VB.Label Label4 
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10335
      TabIndex        =   9
      Top             =   7950
      Width           =   780
   End
   Begin VB.Label Label3 
      Caption         =   "Wheel Size"
      Height          =   210
      Left            =   11160
      TabIndex        =   6
      Top             =   615
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11400
      TabIndex        =   3
      Top             =   1185
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Teeth"
      Height          =   225
      Left            =   11100
      TabIndex        =   2
      Top             =   1800
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'**                          Draw Peg Gear Template
'**                               Version 1.0.0
'**                               By Ken Foster
'**                                 May  2008
'**                     Freeware--- no copyrights claimed
'*******************************************************************


Option Explicit

Public X As Long
Public Y As Long
Public Z As Integer
Public degree As Double
Public radiusX As Long
Public radiusY As Long
Public convert As Double
Public GearCenterX As Integer
Public GearCenterY As Integer

Const PI = 3.14159265358979

Private Sub cmdSave_Click()
   Picture1.Picture = Picture1.Image
   SavePicture Picture1, App.Path & "\gearSaved.bmp"
   MsgBox "Saved in " & App.Path & "\gearSaved.bmp"
End Sub

Private Sub Form_Load()
   Picture1.AutoRedraw = True
End Sub

Private Sub Form_Resize()
   Picture1.Cls
   GearCenterX = (Picture1.Width / 2) - 50
   GearCenterY = (Picture1.Height / 2) + 200
   HS2.Value = GearCenterX
   VS2.Value = GearCenterY
   DrawGear
End Sub

Private Sub DrawGear()
   Dim lngDeg As Double
   Dim xz As Integer
   
   radiusX = HS1.Value
   radiusY = HS1.Value
   
   Picture1.Cls
   Picture1.Refresh
   For Z = 1 To Text1.Text
      lngDeg = 360 / Text1.Text
      degree = Z * lngDeg
      convert = PI / 180                           'from Radian to Degree
      X = GearCenterX - (Sin(-degree * convert) * radiusX)
      Y = GearCenterY - (Sin((90 + (degree)) * convert) * radiusY)
      
      Picture1.CurrentX = X - 40
      Picture1.CurrentY = Y - 110
      
      'draw marks
      If Check1.Value = Unchecked Then    'side peg marks
      Picture1.ForeColor = vbBlack
      Picture1.Print "+"
      Picture1.Circle (GearCenterX, GearCenterY), HS1.Value + 150   'draw outer circle
   Else                                                    'end peg marks
      Picture1.ForeColor = vbBlack
      Picture1.Line (GearCenterX, GearCenterY)-(Picture1.CurrentX + 40, Picture1.CurrentY + 110)  'draw marks
      Picture1.Circle (GearCenterX, GearCenterY), HS1.Value           'draw outer circle
      Picture1.ForeColor = vbWhite      'hides next outer circle
      Picture1.FillStyle = 0
      Picture1.FillColor = vbWhite          'fill circle with this color
      Picture1.Circle (GearCenterX, GearCenterY), HS1.Value - 100  'this circle is filled with white to cover the black lines, saves ink when printing and looks better too
      Picture1.FillStyle = 1
   End If
Next Z

'draw center circle
Picture1.ForeColor = vbBlack
Picture1.Circle (GearCenterX, GearCenterY), 150
Label2.Caption = Format((HS1.Value / 614.29), "##.00")            'approx inches (very approx)

'print labels
Picture1.CurrentX = 80
Picture1.CurrentY = 100
Label7.Caption = "Wheel Size--- " & Label2.Caption
Label8.Caption = " Number of Teeth--- " & Text1.Text
Label9.Caption = " Ratio Between Marks--- " & Format((Label2.Caption) / (Text1.Text), "##.##00")

'draw crosshairs in center
Picture1.Line (GearCenterX - 230, GearCenterY)-(GearCenterX + 250, GearCenterY)
Picture1.Line (GearCenterX, GearCenterY - 250)-(GearCenterX, GearCenterY + 250)
End Sub

Private Sub Check1_Click()               'determines if end or side marks are to be drawn
   DrawGear
   If Check3.Value = Checked And Check1.Value = Checked Then       'point of contact
      Picture1.ForeColor = vbRed
      Picture1.Circle (GearCenterX, GearCenterY), HS1.Value + 100
      Picture1.ForeColor = vbBlack
   Else
      Picture1.ForeColor = vbWhite
      Picture1.Circle (GearCenterX, GearCenterY), HS1.Value + 100
      Picture1.ForeColor = vbBlack
   End If
   Check2_Click
End Sub

Private Sub Check2_Click()
   If HS1.Value < 614 Then Exit Sub                    'if wheel size is less than 1 then do not show labels
   If Check2.Value = Checked Then
      Picture1.CurrentX = HS2.Value - 100
      Picture1.CurrentY = VS2.Value - 430
      Picture1.Print Text1.Text
      Picture1.CurrentX = HS2.Value - 250
      Picture1.CurrentY = VS2.Value + 230
      Picture1.Print Format((Label2.Caption) / (Text1.Text), "##.##00")
   Else
      Picture1.ForeColor = vbWhite
      Picture1.CurrentX = HS2.Value - 100
      Picture1.CurrentY = VS2.Value - 430
      Picture1.Print Text1.Text
      Picture1.CurrentX = HS2.Value - 250
      Picture1.CurrentY = VS2.Value + 230
      Picture1.Print Format((Label2.Caption) / (Text1.Text), "##.##00")
      Picture1.ForeColor = vbBlack
   End If
End Sub

Private Sub Check3_Click()
   Check1_Click
   Check2_Click
End Sub

Private Sub cmdPrint_Click()
   Picture1.Picture = Picture1.Image
   Printer.PaintPicture Picture1.Picture, 0, 0
   Printer.EndDoc
End Sub

Private Sub cmdSet_Click()
    DrawGear
    Check3_Click
End Sub

Private Sub cmdSetPos_Click()
   Picture1.Picture = Picture1.Image
End Sub

Private Sub cmdStartGear_Click()
   GearCenterX = (Picture1.Width / 2) - 50
   GearCenterY = (Picture1.Height / 2) + 200
   HS2.Value = GearCenterX
   VS2.Value = GearCenterY
End Sub

Private Sub HS1_Change()
   HS1_Scroll
End Sub

Private Sub HS1_Scroll()
   DrawGear
   Check1_Click
   Check2_Click
End Sub

Private Sub HS2_Change()
   HS2_Scroll
End Sub

Private Sub HS2_Scroll()
   GearCenterX = HS2.Value
   DrawGear
   Check1_Click
   Check2_Click
End Sub

Private Sub VS2_Change()
   VS2_Scroll
End Sub

Private Sub VS2_Scroll()
   GearCenterY = VS2.Value
   DrawGear
   Check1_Click
   Check2_Click
End Sub
