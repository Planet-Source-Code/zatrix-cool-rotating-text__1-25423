VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rotate Text Zatrix@load.com"
   ClientHeight    =   2820
   ClientLeft      =   3090
   ClientTop       =   2040
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Text"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DrawRotatedText(ByVal txt As String, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal font_name As String, ByVal size As Long, _
    ByVal weight As Long, ByVal escapement As Long, _
    ByVal use_italic As Boolean, ByVal use_underline As Boolean, _
    ByVal use_strikethrough As Boolean)

Const CLIP_LH_ANGLES = 16   ' Needed for tilted fonts.
Const PI = 3.14159625
Const PI_180 = PI / 180#

Dim newfont As Long
Dim oldfont As Long

    newfont = CreateFont(size, 0, _
        escapement, escapement, weight, _
        use_italic, use_underline, _
        use_strikethrough, 0, 0, _
        CLIP_LH_ANGLES, 0, 0, font_name)
    
    ' Select the new font.
    oldfont = SelectObject(hdc, newfont)
    
    ' Display the text.
    CurrentX = X
    CurrentY = Y
    Print txt

    ' Restore the original font.
    newfont = SelectObject(hdc, oldfont)
    
    ' Free font resources (important!)
    DeleteObject newfont
End Sub
Private Sub Command1_Click()
Timer1 = True
End Sub

Private Sub Timer1_Timer()
Const FW_NORMAL = 400   ' Normal font weight.

Static angle As Long    ' Angle in degrees.

    ' Start from scratch.
    Cls
    
    DrawRotatedText "ZATRiX", 100, 100, _
        "Times New Roman", 20, _
        FW_NORMAL, angle * 10, _
        False, False, False

    angle = angle + 5
End Sub
