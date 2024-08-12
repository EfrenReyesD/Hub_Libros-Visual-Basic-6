VERSION 5.00
Begin VB.Form frmViewUserProfile1 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangePassword 
      BackColor       =   &H80000002&
      Caption         =   "Cambiar Contraseña"
      BeginProperty Font 
         Name            =   "Courgette"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   5055
   End
   Begin VB.Label LbViewDateCreateProfile 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   5295
   End
   Begin VB.Label LbViewApellidoUser 
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
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label LbViewNombreUser 
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
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lb1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido: "
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
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lb1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   990
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbUserNameProfile 
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
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   1965
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgProfileUser 
      Height          =   3975
      Left            =   480
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmViewUserProfile1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Unload frmListBooksNG
Unload frmListaUsuarios


Me.Caption = controlNameUser


'cargar la consulta sql
AbrirTablaUsers

With RsTablaUsers
.MoveFirst
Do While Not .EOF
    'Aqui realizará las modificaciones si es que ya existe el registro
    If !UserId = controlUser Then
     
        lbUserNameProfile.Caption = !UserName
        LbViewNombreUser.Caption = !FirstName
        LbViewApellidoUser.Caption = !LastName
        
        LbViewDateCreateProfile.Caption = !CreatedAt
        
        
        
        LoadImageFromUrl RsTablaUsers!ProfilePictureUrl
        
    End If
    .MoveNext
Loop

End With





End Sub




Private Sub LoadImageFromUrl(ByVal imageUrl As String)
    On Error GoTo ErrorHandler
    
    ' Asumiendo que estás usando un control Image
    imgProfileUser.Picture = LoadPicture(imageUrl)
        imgProfileUser.Stretch = True
    
    ' Si no ocurre ningún error, se sale del procedimiento
    Exit Sub

ErrorHandler:
    ' Aquí manejas el error, por ejemplo, mostrando una imagen de reserva
    imgProfileUser.Picture = LoadPicture("C:\Users\Efren\Desktop\MEGA\proyectos visual basic\hubLibros\backgrounds\bookbg.jpg")
    imgProfileUser.Stretch = True
    MsgBox "No se pudo cargar la imagen. Mostrando imagen de reserva.", vbExclamation, "Error de carga de imagen"
End Sub

