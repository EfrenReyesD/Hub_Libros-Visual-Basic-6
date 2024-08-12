VERSION 5.00
Object = "{3A8BD65E-9922-4162-A649-83F2D5326BBE}#1.0#0"; "FoxitReaderBrowserAx.dll"
Begin VB.Form frmReadPdf 
   Caption         =   "Lectura"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClosePdf 
      BackColor       =   &H00C0C000&
      Caption         =   "Dejar de Leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9480
      Width           =   4695
   End
   Begin FOXITREADERLibCtl.FoxitCtl FoxitCtlViewPdf 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17895
      _cx             =   5080
      _cy             =   5080
      src             =   ""
   End
End
Attribute VB_Name = "frmReadPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlUpl1_UploadedFile(ByVal bstrFileName As String, ByVal lErrorCode As Long, ByVal bstrErrMsg As String)

End Sub

Private Sub cmdClosePdf_Click()
Unload Me

End Sub
