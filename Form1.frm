VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Msimg32 Test"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTransparent 
      Caption         =   "Transparent"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "AlphaBlend"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdRect 
      Caption         =   "Rectangle"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTriangle 
      Caption         =   "Triangle"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ClearForm"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   5040
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drawing functions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gradient Fill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
    
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Const GRADIENT_FILL_TRIANGLE As Long = &H2
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GradientFillTri Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Private Function LongToUShort(ULong As Long) As Integer
   LongToUShort = CInt(ULong - &H10000)
End Function

Private Sub cmdClear_Click()
   Cls
End Sub

Private Sub cmdRect_Click()
Dim vert(1) As TRIVERTEX
Dim gRect As GRADIENT_RECT
With vert(0)
    .x = 0
    .y = 0
    .Red = 0
    .Green = &HFF&
    .Blue = 0
    .Alpha = 0
End With
With vert(1)
    .x = Me.ScaleWidth
    .y = Me.ScaleHeight
    .Red = 0
    .Green = LongToUShort(&HFF00&)
    .Blue = LongToUShort(&HFF00&)
    .Alpha = 0
End With
gRect.UpperLeft = 1
gRect.LowerRight = 0
GradientFillRect Me.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
Me.Refresh
End Sub

Private Sub cmdTriangle_Click()
'!!!Play with colors!!!
Dim vert(3) As TRIVERTEX
Dim gTri(1) As GRADIENT_TRIANGLE
With vert(0)
    .x = 0
    .y = 0
    .Red = 0&
    .Green = LongToUShort(&HFF00&) '0
    .Blue = 0&
    .Alpha = 0&
End With
With vert(1)
    .x = Me.ScaleWidth
    .y = 0
    .Red = 0 'LongToUShort(&HFF00&)
    .Green = 0&
    .Blue = LongToUShort(&HFF00&)
    .Alpha = 0&
End With
With vert(2)
    .x = Me.ScaleWidth
    .y = Me.ScaleHeight
    .Red = LongToUShort(&HFF00&)
    .Green = 0&
    .Blue = 0 'LongToUShort(&HFF00&)
    .Alpha = 0&
End With
With vert(3)
    .x = 0
    .y = Me.ScaleHeight
    .Red = 0 'LongToUShort(&HFF00&)
    .Green = LongToUShort(&HFF00&)
    .Blue = LongToUShort(&HFF00&)
    .Alpha = 0&
End With
gTri(0).Vertex1 = 0
gTri(0).Vertex2 = 1
gTri(0).Vertex3 = 2

gTri(1).Vertex1 = 0
gTri(1).Vertex2 = 2
gTri(1).Vertex3 = 3

GradientFillTri Me.hdc, vert(0), 4, gTri(0), 2, GRADIENT_FILL_TRIANGLE
Me.Refresh
End Sub

Private Sub cmdTransparent_Click()
  Dim clr As Long, w As Long, h As Long
  w = Picture1.ScaleWidth
  h = Picture1.ScaleHeight
  clr = Picture1.Point(0, 0)
  Call TransparentBlt(Me.hdc, ScaleWidth - w, 0, w, h, Picture1.hdc, 0, 0, w, h, clr)
  Me.Refresh
End Sub

Private Sub cmdAlpha_Click()
  Dim w As Long, h As Long
  w = Picture1.ScaleWidth
  h = Picture1.ScaleHeight
  Call AlphaBlend(Me.hdc, 0, 0, w, h, Picture1.hdc, 0, 0, w, h, 50)
  Refresh
End Sub

Private Sub Form_Load()
   Me.ScaleMode = vbPixels
   Picture1.ScaleMode = vbPixels
   Picture1.Visible = False
End Sub

