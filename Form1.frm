VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SideBox"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ctlSideMenu ctlSideMenu4 
      Height          =   720
      Left            =   2025
      TabIndex        =   3
      Top             =   2520
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   1270
      Align           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "MS Sans Serif"
   End
   Begin Project1.ctlSideMenu ctlSideMenu3 
      Height          =   4920
      Left            =   9135
      TabIndex        =   2
      Top             =   45
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   8678
      Align           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      FontSize        =   12
      FontName        =   "MS Sans Serif"
   End
   Begin Project1.ctlSideMenu ctlSideMenu2 
      Height          =   1500
      Left            =   2520
      TabIndex        =   1
      Top             =   2475
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "MS Sans Serif"
   End
   Begin Project1.ctlSideMenu ctlSideMenu1 
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   2461
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      FontSize        =   8.25
      ForeColor       =   16576
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ctlSideMenu1.AddItem "Option One"
    ctlSideMenu1.AddItem "Option Two"
    ctlSideMenu1.AddItem "A longer option string."
    ctlSideMenu1.AddItem "Option Three"
    ctlSideMenu1.AddItem "Option Four"
    ctlSideMenu1.AddItem "Option Five"

    ctlSideMenu2.AddItem "Option One"
    ctlSideMenu2.AddItem "Option Two"
    ctlSideMenu2.AddItem "A longer option string."
    ctlSideMenu2.AddItem "Option Three"
    ctlSideMenu2.AddItem "Option Four"
    ctlSideMenu2.AddItem "Option Five"
    
    ctlSideMenu3.AddItem "Option One"
    ctlSideMenu3.AddItem "Option Two"
    ctlSideMenu3.AddItem "A longer option string."
    ctlSideMenu3.AddItem "Option Three"
    ctlSideMenu3.AddItem "Option Four"
    ctlSideMenu3.AddItem "Option Five"
End Sub
