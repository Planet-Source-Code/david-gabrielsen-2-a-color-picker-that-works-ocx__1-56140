VERSION 5.00
Begin VB.UserControl ctrColorPicker 
   CanGetFocus     =   0   'False
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   2940
   Begin VB.Label Label3 
      Caption         =   "A Thousand Thank Yous to Branco Medeiros who made the HSB-color model available for VB!!"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "By David Gabrielsen"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "ColorPicker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "ctrColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public retval As Long
Public Function ShowColor(Optional Title As String, Optional DefaultColor As Long) As Long
    strTitle = Title
    lngDefaultColor = DefaultColor
    Load frmMain
    frmMain.Show 1, Me
    ShowColor = frmMain.Color
End Function


