VERSION 5.00
Begin VB.Form frmlogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio Sesión"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "Login"
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
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "123456"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   1
      Text            =   "Efren"
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lbContraseña 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lbUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lbInfLogin 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su usuario y contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lbbienvenido 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   3555
      Left            =   0
      Picture         =   "frmlogin.frx":1084A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5685
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()

'validar el ingreso de informacion
If Trim(txtUser.Text) = "" Then MsgBox "Ingrese usuario", vbInformation, "Aviso": txtUser.SetFocus: Exit Sub
If Trim(txtPass.Text) = "" Then MsgBox "Ingrese Contraseña", vbInformation, "Aviso": txtPass.SetFocus: Exit Sub

'buscar el usuario
With RsTablaUsers
.Requery
.Find "UserName = '" & Trim(txtUser.Text) & "'"


'si no se encontro nada
If .EOF Then
    MsgBox "No se encontro el usuario", vbInformation, "Aviso"
    txtUser.Text = ""
    txtUser.SetFocus
    Exit Sub
Else
'Si encontro al usuario y hay que validar la contraseña
    If !Password = Trim(txtPass.Text) Then
        controlUser = !UserId
        controlNameUser = !UserName
        
        frmPrincipal.Show 'Muestro el formulario principal
        frmPrincipal.mnuMiCuenta.Caption = controlNameUser
        If !Admin = 0 Then
            
            frmPrincipal.mnuAdmin.Visible = False
            frmPrincipal.mnudeleteLibro.Visible = False
            frmListLibros.cmdDeleteBook.Visible = False
            
        Else
            frmPrincipal.mnuAdmin.Visible = True
            frmListLibros.cmdDeleteBook.Visible = True
        End If
        Unload Me 'Cierro el formulario login
    Else
        'no es correcta la clave
        MsgBox "La contraseña es incorrecta", vbInformation, "Aviso"
        txtPass.Text = ""
        txtPass.SetFocus
        Exit Sub
         
    End If

End If



End With



End Sub

Private Sub Form_Load()
    AbrirTablaUsers
    ' Configurar las propiedades del control Image
    Image1.Stretch = True
    ' Image1.Picture = LoadPicture("ruta_de_tu_imagen.jpg")
    
    ' Ajustar el tamaño del control Image para que cubra todo el formulario
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight
    
    'abrir nuestra tabla users
    
    
End Sub


