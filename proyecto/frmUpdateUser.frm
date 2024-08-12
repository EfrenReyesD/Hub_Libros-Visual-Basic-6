VERSION 5.00
Begin VB.Form frmUpdateUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación de Usuario"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInsertUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "Modificar Usuario"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   4095
   End
   Begin VB.ComboBox cbBxOption 
      Height          =   315
      Left            =   6840
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtConfimPassInsert 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6720
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtPassInsert 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6720
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtUsuarioInsert 
      Height          =   285
      Left            =   6480
      MaxLength       =   16
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtApellidoInsert 
      Height          =   285
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtNombreInsert 
      Height          =   285
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de usuario:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   12
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5280
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5400
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Credenciales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lbUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lbUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lbTitleUpdateUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar Usuario"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmUpdateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbTitleInsertUser_Click(Index As Integer)

End Sub
