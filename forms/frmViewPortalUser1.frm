VERSION 5.00
Begin VB.Form frmViewPortalBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista de Libro"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmViewPortalUser1.frx":0000
   ScaleHeight     =   7590
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReadBook 
      BackColor       =   &H80000002&
      Caption         =   "Leer"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox txtViewDescription 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5160
      Width           =   4455
   End
   Begin VB.Label lb4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion :"
      BeginProperty Font 
         Name            =   "Playball"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label LbViewGenre 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label LbViewAuthor 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label LbViewTitle 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lb3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Genero : "
      BeginProperty Font 
         Name            =   "Playball"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   975
      TabIndex        =   3
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label lb2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autor : "
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
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      Width           =   1020
   End
   Begin VB.Label lb1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo : "
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
      Left            =   1080
      TabIndex        =   1
      Top             =   3480
      Width           =   990
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImgViewBook 
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmViewPortalUser1.frx":59A71
      Height          =   3135
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmViewPortalBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdReturnListBooks_Click()
Unload Me

End Sub

Private Sub cmdReadBook_Click()
Unload frmListaUsuarios
Unload Me
frmReadPdf.Show
'***************boton abrir pdf*******************+
'cargar la consulta sql
abrirAllBooks

With RsAllBooks
.MoveFirst
Do While Not .EOF
    'Aqui realizará las modificaciones si es que ya existe el registro
    If !BookId = controlIdBookSelectView Then
     frmReadPdf.FoxitCtlViewPdf.OpenFile (!PdfUrl)
        
    End If
    .MoveNext
Loop

End With



End Sub

Private Sub Form_Load()
'cargar la consulta sql
abrirAllBooks

With RsAllBooks
.MoveFirst
Do While Not .EOF
    'Aqui realizará las modificaciones si es que ya existe el registro
    If !BookId = controlIdBookSelectView Then
     
        LbViewTitle.Caption = !Title
        LbViewAuthor.Caption = !author
        LbViewGenre.Caption = !Genre
        txtViewDescription.Text = !Description
        txtViewDescription.TabStop = False
        ImgViewBook.Picture = LoadPicture(!CoverImage)
        ImgViewBook.Stretch = True
        
    End If
    .MoveNext
Loop

End With


End Sub
