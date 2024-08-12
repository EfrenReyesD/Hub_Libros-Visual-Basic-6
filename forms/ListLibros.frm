VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListLibros 
   Caption         =   "Todos los Libros"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ListLibros.frx":0000
   ScaleHeight     =   10530
   ScaleMode       =   0  'User
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFiltrarBooks 
      Height          =   405
      Left            =   15000
      MaxLength       =   16
      TabIndex        =   7
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cbFiltrarBooks 
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12240
      TabIndex        =   6
      Text            =   "Titulo"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdOpenInfBook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "Ver Libro"
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
      Index           =   1
      Left            =   15600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9480
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGridListBooks 
      Height          =   7935
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   13996
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label cmdAddBook 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Añadir Libro"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Index           =   2
      Left            =   9360
      TabIndex        =   9
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label cmdDeleteBook 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Eliminar Libro"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   12120
      TabIndex        =   8
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label cmdAddRead 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marcar como Leído"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Index           =   0
      Left            =   6600
      TabIndex        =   4
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label cmdAddNG 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marcar como NG"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label cmdAddFavorite 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Añadir a Favorito"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label lbListBooks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido a la Biblioteca de los Mejores Libros"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13335
   End
End
Attribute VB_Name = "frmListLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAddFavorite_Click(Index As Integer)
Dim EncontroRegistro As Boolean
EncontroRegistro = False
' Modificar los campos del Recordset
With RsUserBooks
.MoveFirst
Do While Not .EOF
    'Aqui realizará las modificaciones si es que ya existe el registro
    If !userId = controlUser And !BookId = DataGridListBooks.Columns(0).Text Then
        EncontroRegistro = True 'agregamos una variable llave para saber si encontro el registro o no
        Select Case !IsFavorite
            Case True
                !IsFavorite = False
                If !IsDisliked = True Then
                    !IsDisliked = False
                End If
            Case False
                !IsFavorite = True
                If !IsDisliked = True Then
                    !IsDisliked = False
                End If
        End Select
    End If
    .MoveNext
Loop

End With

'anexaremos codigo para crear un nuevo registro de status si aun no se encuentra
If (EncontroRegistro = False) Then
    RsUserBooks.AddNew
    RsUserBooks!userId = controlUser
    RsUserBooks!BookId = DataGridListBooks.Columns(0).Text
    RsUserBooks!IsFavorite = True
    RsUserBooks.Update
End If
updateGridListBooks
     
End Sub
Private Sub updateGridListBooks()
    RsViewBooks.Requery
    Set DataGridListBooks.DataSource = RsViewBooks
    EstiloGridBooks
    DataGridListBooks.Refresh
End Sub

Private Sub cmdAddNG_Click(Index As Integer)
Dim EncontroRegistro As Boolean
EncontroRegistro = False
With RsUserBooks
.MoveFirst
Do While Not .EOF

    If !userId = controlUser And !BookId = DataGridListBooks.Columns(0).Text Then
        EncontroRegistro = True 'agregamos una variable llave para saber si encontro el registro o no
        Select Case !IsDisliked
            Case True
                !IsDisliked = False
                If !IsFavorite = True Then
                    !IsFavorite = False
                End If
            Case False
                !IsDisliked = True
                If !IsFavorite = True Then
                    !IsFavorite = False
                End If
        End Select
    End If
    .MoveNext
Loop
End With
'anexaremos codigo para crear un nuevo registro de status si aun no se encuentra
If (EncontroRegistro = False) Then
    RsUserBooks.AddNew
    RsUserBooks!userId = controlUser
    RsUserBooks!BookId = DataGridListBooks.Columns(0).Text
    RsUserBooks!IsDisliked = True
    RsUserBooks.Update
End If
updateGridListBooks
End Sub

Private Sub cmdAddRead_Click(Index As Integer)
Dim EncontroRegistro As Boolean
EncontroRegistro = False
With RsUserBooks
.MoveFirst
Do While Not .EOF
    If !userId = controlUser And !BookId = DataGridListBooks.Columns(0).Text Then
        EncontroRegistro = True 'agregamos una variable llave para saber si encontro el registro o no
        Select Case !IsRead
            Case True
                !IsRead = False
            Case False
                !IsRead = True
        End Select
    End If
    .MoveNext
Loop
End With

'anexaremos codigo para crear un nuevo registro de status si aun no se encuentra
If (EncontroRegistro = False) Then
    RsUserBooks.AddNew
    RsUserBooks!userId = controlUser
    RsUserBooks!BookId = DataGridListBooks.Columns(0).Text
    RsUserBooks!IsRead = True
    RsUserBooks.Update
End If
updateGridListBooks

End Sub

Private Sub cmdOpenInfBook_Click(Index As Integer)
controlIdBookSelectView = DataGridListBooks.Columns(0).Text


frmViewPortalBook.Show vbModal
End Sub

Private Sub Form_Load()
Unload frmReadPdf
Unload frmListaUsuarios
abrirUserBooks
viewBooks (controlUser)

Set DataGridListBooks.DataSource = RsViewBooks
EstiloGridBooks

'add items al combobox
cbFiltrarBooks.AddItem "Titulo"

End Sub


Sub EstiloGridBooks()
    Dim ColumnsWidth As Integer
    'Tamaño de columnas
    ColumIdWidth = 800
    ColumnsWidth = ((DataGridListBooks.Width - 325) - ColumIdWidth) / (DataGridListBooks.Columns.Count - 1)
    
    
    'Tamaños
    Dim i As Integer
    DataGridListBooks.Columns(0).Width = ColumIdWidth
    For i = 1 To (DataGridListBooks.Columns.Count - 1)
        DataGridListBooks.Columns(i).Width = ColumnsWidth
    Next i
    
    'caption
    Dim namesColumns As Variant
    namesColumns = Array("Id Libro", "Titulo", "Autor", "Genero", "Status")
    
    For i = 0 To (DataGridListBooks.Columns.Count - 1)
        DataGridListBooks.Columns(i).Caption = namesColumns(i)
    
    Next i
    
    'Alineacion
    For i = 0 To (DataGridListBooks.Columns.Count - 1)
        DataGridListBooks.Columns(i).Alignment = dbgCenter
        
    Next i
    
    DataGridListBooks.HeadFont.Bold = True
    
    
    
End Sub







Private Sub txtFiltrarBooks_Change()
    'RsViewUserFilter.Requery
    DataGridListBooks.Refresh
    EstiloGridBooks
    
    Dim filtro As String
    Dim columna As String
    
    columna = cbFiltrarBooks.Text
    filtro = txtFiltrarBooks.Text
    
    If filtro = "" Then
        RsViewBooks.Requery
        DataGridListBooks.Refresh
        Set DataGridListBooks.DataSource = RsViewBooks
        EstiloGridBooks
    End If
    
    
    'Cambiaremos el nombre de la columna a como aparece en la base de datos
    
    
    If columna <> "" Then
        If filtro <> "" Then
            viewBooksFilter controlUser, filtro
            Set DataGridListBooks.DataSource = RsViewBooksFilter
            EstiloGridBooks
        End If
    
    End If
    
    
    
    
End Sub

