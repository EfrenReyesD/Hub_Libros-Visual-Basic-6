VERSION 5.00
Begin VB.Form frmViewUserProfile 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LbViewApelliUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lb2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido : "
      BeginProperty Font 
         Name            =   "Playball"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   3120
      Width           =   1320
   End
   Begin VB.Label LbViewNameUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lb2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre : "
      BeginProperty Font 
         Name            =   "Playball"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lbNameUserProfile 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Playball"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3705
      TabIndex        =   0
      Top             =   480
      Width           =   2205
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   480
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmViewUserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

