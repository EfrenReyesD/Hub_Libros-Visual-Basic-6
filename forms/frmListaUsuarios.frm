VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaUsuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Listado de Usuarios"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000014&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "frmListaUsuarios.frx":0000
   ScaleHeight     =   10530
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitListUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "Salir"
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
      Left            =   15720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9360
      Width           =   1455
   End
   Begin VB.TextBox txtFiltrar 
      Height          =   405
      Left            =   12840
      MaxLength       =   16
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cbFiltrar 
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
      Left            =   10200
      TabIndex        =   3
      Text            =   "Nombre"
      Top             =   360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGridListUsers 
      Height          =   7935
      Left            =   600
      TabIndex        =   2
      Top             =   1080
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
   Begin VB.Label cmdDeleteUser 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Eliminar Usuario"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label cmdInsertUser 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Usuario"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label lbListUsers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lista General de Usuarios"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   -360
      TabIndex        =   1
      Top             =   240
      Width           =   13335
   End
End
Attribute VB_Name = "frmListaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDeleteUser_Click(Index As Integer)

'Verificar si la tabla esta vacia
With RsViewUsers
If .RecordCount = 0 Then Exit Sub
End With

'Comparar que tipo de usuario es y si es admin no podra eliminar
If DataGridListUsers.Columns(4).Text = "Admin" Then

MsgBox "No puedes eliminar a " + DataGridListUsers.Columns(2) + " porque es administrador", vbInformation, "Error"
Else

'Variables para el manejo de eliminacion del usuario
Dim respuesta As Integer
Dim IdUser As Integer

respuesta = MsgBox("¿Estas seguro de eliminar a " + DataGridListUsers.Columns(2) + " ?", vbYesNo + vbQuestion, "Confirmación")

If respuesta = vbYes Then
    ' MsgBox "Has elegido 'Sí'.", vbInformation, "Resultado"
    IdUser = DataGridListUsers.Columns(0).Text
    MsgBox IdUser
    With RsViewUsers
    .Requery
    .Find "UserId = '" & IdUser & "'" 'Ya tengo ubicado al usuario
    .Delete
    .Requery
    'Set DataGridListUsers.DataSource = rs
    EstiloGrid
    'IdUser = 0
    End With
    
    
ElseIf respuesta = vbNo Then
    MsgBox "Has elegido 'No'.", vbInformation, "Resultado"
End If

End If




End Sub

Private Sub cmdExitListUsers_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdInsertUser_Click(Index As Integer)
frmAddUser.Show vbModal



End Sub

Private Sub cmdUpdateUser_Click(Index As Integer)
frmUserUpdate.Show vbModal
End Sub

Private Sub Form_Load()
Unload frmListLibros
Unload frmReadPdf
Me.WindowState = vbMaximized
viewUsers
'llenar la grilla de usuarios y dar estilos
Set DataGridListUsers.DataSource = RsViewUsers

'llamar la función
EstiloGrid
'add items al combobox
cbFiltrar.AddItem "ID Usuario"
cbFiltrar.AddItem "Usuario"
cbFiltrar.AddItem "Nombre"
cbFiltrar.AddItem "Apellido"

End Sub

Sub EstiloGrid()
    Dim ColumnsWidth As Integer
    Dim ColumIdWidth As Integer
    'Tamaño de columnas
    ColumIdWidth = 1000 'Especificamos el ancho de la columna ID de la tabla.
    ColumnsWidth = ((DataGridListUsers.Width - 325) - ColumIdWidth) / (DataGridListUsers.Columns.Count - 1)
    
    
    'Tamaños
    Dim i As Integer
    DataGridListUsers.Columns(0).Width = ColumIdWidth
    For i = 1 To (DataGridListUsers.Columns.Count - 1)
        DataGridListUsers.Columns(i).Width = ColumnsWidth
    Next i
    
    'caption
    Dim namesColumns As Variant
    namesColumns = Array("ID Usuario", "Usuario", "Nombre", "Apellido", "Tipo Usuario")
    
    For i = 0 To (DataGridListUsers.Columns.Count - 1)
        DataGridListUsers.Columns(i).Caption = namesColumns(i)
    
    Next i
    
    'Alineacion
    For i = 0 To (DataGridListUsers.Columns.Count - 1)
        DataGridListUsers.Columns(i).Alignment = dbgCenter
        
    Next i
    
    DataGridListUsers.HeadFont.Bold = True
    
    
    
End Sub

Private Sub txtFiltrar_Change()
    'RsViewUserFilter.Requery
    DataGridListUsers.Refresh
    EstiloGrid
    
    Dim filtro As String
    Dim columna As String
    
    columna = cbFiltrar.Text
    filtro = txtFiltrar.Text
    
    If filtro = "" Then
        RsViewUserFilter.Requery
        DataGridListUsers.Refresh
        Set DataGridListUsers.DataSource = RsViewUsers
        EstiloGrid
    End If
    
    
    'Cambiaremos el nombre de la columna a como aparece en la base de datos
    Select Case columna
        Case "ID Usuario"
            columna = "UserId"
        Case "Usuario"
            columna = "UserName"
        Case "Nombre"
            columna = "FirstName"
        Case "Apellido"
            columna = "LastName"
    End Select
    

    
    
    If columna <> "" Then
        If filtro <> "" Then
            viewUsersFilter columna, filtro
            Set DataGridListUsers.DataSource = RsViewUserFilter
            EstiloGrid
        End If
    
    End If
    
    
    
    
End Sub
