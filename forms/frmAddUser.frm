VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Nuevo Usuario"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9660
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAddUser.frx":1084A
   ScaleHeight     =   6420
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUrlPerfilInsert 
      Height          =   285
      Left            =   5760
      MaxLength       =   255
      TabIndex        =   16
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txtConfimPassInsert 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7080
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtPassInsert 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7080
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtUsuarioInsert 
      Height          =   285
      Left            =   7080
      MaxLength       =   16
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox cbBxOption 
      Height          =   315
      Left            =   7080
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtApellidoInsert 
      Height          =   285
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdInsertUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "Registrar Nuevo Usuario"
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
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   4095
   End
   Begin VB.TextBox txtNombreInsert 
      Height          =   285
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Url Foto Perfil"
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
      Index           =   0
      Left            =   3600
      TabIndex        =   15
      Top             =   4680
      Width           =   2535
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
      Left            =   5280
      TabIndex        =   11
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   5280
      Top             =   1440
      Width           =   4215
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
      Left            =   5640
      TabIndex        =   9
      Top             =   3480
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
      Left            =   5520
      TabIndex        =   8
      Top             =   2880
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
      Left            =   5760
      TabIndex        =   7
      Top             =   2280
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
      Left            =   6000
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
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
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lbTitleInsertUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Nuevo Usuario"
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
      TabIndex        =   3
      Top             =   240
      Width           =   3495
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
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()



End Sub

Private Sub cmdInsertUser_Click()
'Validar ingreso de informacion de todos los campos

If Trim(txtNombreInsert) = "" Then MsgBox "Debe llenar el campo nombre", vbInformation, "Alerta": txtNombreInsert.SetFocus: Exit Sub
If Trim(txtApellidoInsert) = "" Then MsgBox "Debe llenar el campo Apellido", vbInformation, "Alerta": txtApellidoInsert.SetFocus: Exit Sub
If Trim(txtUsuarioInsert) = "" Then MsgBox "Agrega un Usuario", vbInformation, "Alerta": txtUsuarioInsert.SetFocus: Exit Sub
If Trim(txtPassInsert) = "" Then MsgBox "Agrega una Contraseña", vbInformation, "Alerta": txtPassInsert.SetFocus: Exit Sub
If Trim(txtConfimPassInsert) = "" Then MsgBox "Confirma tu Contraseña", vbInformation, "Alerta": txtConfimPassInsert.SetFocus: Exit Sub
If Trim(cbBxOption) = "" Then MsgBox "Elige el tipo de Usuario", vbInformation, "Alerta": cbBxOption.SetFocus: Exit Sub
If Trim(txtUrlPerfilInsert) = "" Then MsgBox "Ingresa URL de foto de perfil", vbInformation, "Alerta": txtUrlPerfilInsert.SetFocus: Exit Sub
'MsgBox "Usuario registrado correctamente", vbInformation, "Confirmado"

If txtPassInsert.Text <> txtConfimPassInsert.Text Then
    MsgBox "La contraseña no coincide"
Else
    'Validar si el usuario ya existe
    With RsTablaUsers
    .Requery
    .Find "UserName='" & Trim(txtUsuarioInsert.Text) & "'"
    
    Dim typeUser As Integer
    If cbBxOption = "Admin" Then
        typeUser = 1
    Else
        typeUser = 0
    End If
    
        
    
    If .EOF Then 'si no encontro nada
    'Significa que se debe agregar el usuario como nuevo
    .AddNew
    !UserName = txtUsuarioInsert.Text
    !Password = txtPassInsert.Text
    !Admin = typeUser
    !FirstName = txtNombreInsert.Text
    !LastName = txtApellidoInsert.Text
    !ProfilePictureUrl = txtUrlPerfilInsert.Text
    .Update
    
    Set frmListaUsuarios.DataGridListUsers.DataSource = RsViewUsers
    
    Else
    MsgBox "El Usuario ya Existe", vbInformation, "Alerta": txtUsuarioInsert.SetFocus: Exit Sub
    
    End If
    
    End With
    
    Set frmListaUsuarios.DataGridListUsers.DataSource = RsViewUsers
    frmListaUsuarios.EstiloGrid
    Unload Me
    RsViewUsers.Requery
    frmListaUsuarios.DataGridListUsers.Refresh
    frmListaUsuarios.EstiloGrid
    
End If





End Sub


Private Sub Form_Load()
cbBxOption.AddItem "Cliente"
cbBxOption.AddItem "Admin"
 
End Sub

