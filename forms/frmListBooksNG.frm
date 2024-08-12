VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListBooksNG 
   Caption         =   "Libros NG"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOpenInfBookNG 
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
      Left            =   15720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9480
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGridListNG 
      Height          =   7935
      Left            =   600
      TabIndex        =   0
      Top             =   1200
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Libros Disliked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmListBooksNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub EstiloGridBooksNG()
    Dim ColumnsWidth As Integer
    'Tamaño de columnas
    ColumIdWidth = 800
    ColumnsWidth = ((DataGridListNG.Width - 325) - ColumIdWidth) / (DataGridListNG.Columns.Count - 1)
    
    
    'Tamaños
    Dim i As Integer
    DataGridListNG.Columns(0).Width = ColumIdWidth
    For i = 1 To (DataGridListNG.Columns.Count - 1)
        DataGridListNG.Columns(i).Width = ColumnsWidth
    Next i
    
    'caption
    Dim namesColumns As Variant
    namesColumns = Array("Id Libro", "Titulo", "Autor", "Genero")
    
    For i = 0 To (DataGridListNG.Columns.Count - 1)
        DataGridListNG.Columns(i).Caption = namesColumns(i)
    
    Next i
    
    'Alineacion
    For i = 0 To (DataGridListNG.Columns.Count - 1)
        DataGridListNG.Columns(i).Alignment = dbgCenter
        
    Next i
    
    DataGridListNG.HeadFont.Bold = True
    
    
    
End Sub
