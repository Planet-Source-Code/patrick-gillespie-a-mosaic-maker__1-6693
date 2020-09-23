VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mosaic Maker"
   ClientHeight    =   7650
   ClientLeft      =   735
   ClientTop       =   495
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8895
   Begin VB.PictureBox HoldFadedPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9180
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7860
      TabIndex        =   2
      Top             =   660
      Width           =   975
   End
   Begin VB.PictureBox ThePic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   8100
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox NewPic 
      AutoRedraw      =   -1  'True
      Height          =   7515
      Left            =   60
      ScaleHeight     =   502.02
      ScaleMode       =   0  'User
      ScaleWidth      =   516.162
      TabIndex        =   0
      Top             =   60
      Width           =   7725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mosaic Maker
' by PAT or JK (Patrick Gillespie)
' http://www.patorjk.com/
' 3.19.00

' This is an example on how to change an icon of Drew Carey's Face
' to into one big picture of his face made up of picture's of his face.
' This can be cool for visual effects, or it can be something to do
' when you're bored. I got the idea to make this after seeing a poster
' that was of a picture made out of smaller pictures and thought it'd
' be cool if people had their own program to create pictures like that.

Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub Command1_Click()
    Dim rows As Integer, cols As Integer
    Dim Color1 As Long, Color2 As Long
    Dim TheColor As Long, i As Integer, i2 As Integer
    ' just in case something goes wrong
    On Error Resume Next
    ' ok, here we are going to loop through the pixels in drew's face
    ' comparing them to other pixels in the picture so we can fade
    ' the picture a little to make it look like part of one certain
    ' spot in the picture.
    
    ' make sure the pic holder is the same size as the picture with the pic
    HoldFadedPic.Width = ThePic.Width
    HoldFadedPic.Height = ThePic.Height
    For i = 0 To (ThePic.ScaleWidth - 1) * 16 Step 16
        For i2 = 0 To (ThePic.ScaleHeight - 1) * 16 Step 16
            For cols = 0 To ThePic.ScaleWidth - 1
                For rows = 0 To ThePic.ScaleHeight - 1
                    Color1 = GetPixel(ThePic.hdc, rows, cols)
                    Color2 = GetPixel(ThePic.hdc, i / 16, i2 / 16)
                    If Color1 <> Color2 Then
                        ' get the color pixel we want
                        TheColor = GetFadedColor(Color1, Color2, 3, 4)
                    Else
                        ' the two colors are the same so we don't
                        ' need to waste time with the fading sub
                        TheColor = Color1
                    End If
                    ' set the pixel in our holding picture box
                    Call SetPixel(HoldFadedPic.hdc, rows, cols, TheColor)
                Next
            Next
        ' set the picture's image to it's picture
        Set HoldFadedPic.Picture = HoldFadedPic.Image
        ' paint a the picture onto the big picture
        NewPic.PaintPicture HoldFadedPic.Picture, i, i2, 16, 16
        ' eh, just to be nice
        DoEvents
        Next
    Next
    MsgBox "done"
End Sub

Private Sub Form_Load()
    ' make sure everything is in pixels
    NewPic.ScaleMode = 3
    ThePic.ScaleMode = 3
    HoldFadedPic.ScaleMode = 3
End Sub

Public Function GetFadedColor(c1 As Long, c2 As Long, FN As Integer, FS As Integer) As Long
    Dim i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, cx1!, cx2!, cx3!
    
    ' get the red, green, and blue values out of the different
    ' colors
    red1% = (c1 And 255)
    green1% = (c1 \ 256 And 255)
    blue1% = (c1 \ 65536 And 255)
    red2% = (c2 And 255)
    green2% = (c2 \ 256 And 255)
    blue2% = (c2 \ 65536 And 255)
    
    ' get the step of the color changing
    pat1 = (red2% - red1%) / FS
    pat2 = (green2% - green1%) / FS
    pat3 = (blue2% - blue1%) / FS

    ' set the cx variables at the starting colors
    cx1 = red1%
    cx2 = green1%
    cx3 = blue1%

    ' loop till you reach the faze you are at in the fading
    For i% = 1 To FN
        cx1 = cx1 + pat1
        cx2 = cx2 + pat2
        cx3 = cx3 + pat3
    Next
    GetFadedColor = RGB(cx1, cx2, cx3)
End Function
