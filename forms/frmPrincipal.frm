VERSION 5.00
Begin VB.MDIForm frmPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "HUB LECTURA"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   17880
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmPrincipal.frx":1084A
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuPerfil 
      Caption         =   "Mi Perfil"
      Begin VB.Menu mnuMiCuenta 
         Caption         =   ""
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuCatalogo 
      Caption         =   "Catálogo"
      Begin VB.Menu mnuAllLibros 
         Caption         =   "Todos los libros"
      End
      Begin VB.Menu mnuFavoritos 
         Caption         =   "Favoritos"
      End
   End
   Begin VB.Menu mnuLibros 
      Caption         =   "Libros"
      Begin VB.Menu mnuLeidos 
         Caption         =   "Libros Leidos"
      End
      Begin VB.Menu mnuMalos 
         Caption         =   "Libros No Good"
      End
      Begin VB.Menu mnuAddBook 
         Caption         =   "Agregar Libro"
      End
      Begin VB.Menu mnudeleteLibro 
         Caption         =   "Eliminar Libro"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrador"
      Begin VB.Menu mnuUsers 
         Caption         =   "Usuarios"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
End Sub

Private Sub mnuAllLibros_Click()
    Dim frm As Form
    Dim formFound As Boolean
    formFound = False

    ' Buscar si el formulario ya está abierto
    For Each frm In Forms
        If TypeOf frm Is frmListLibros Then
            frm.WindowState = vbNormal  ' Restaurar primero
            frm.WindowState = vbMaximized  ' Luego maximizar
            frm.Show
            formFound = True
            Exit For
        End If
    Next frm

    ' Si no está abierto, crear una nueva instancia y mostrarlo maximizado
    If Not formFound Then
        Set frmMyForm = New frmListLibros
        'frmMyForm.MDIChild = True
        frmMyForm.Show
        frmMyForm.WindowState = vbMaximized
    End If









Unload frmListaUsuarios
Unload frmReadPdf
'frmListLibros.Show
End Sub

Private Sub mnuFavoritos_Click()
    viewBooksFavorites (controlUser)
    'RsViewBooksFavorites.MoveFirst
    Unload frmFavoritos
    Dim frm As Form
    Dim formFound As Boolean
    formFound = False

    ' Buscar si el formulario ya está abierto
    For Each frm In Forms
        If TypeOf frm Is frmFavoritos Then
            frm.WindowState = vbNormal  ' Restaurar primero
            frm.WindowState = vbMaximized  ' Luego maximizar
            frm.Show
            formFound = True
            Exit For
        End If
    Next frm

    ' Si no está abierto, crear una nueva instancia y mostrarlo maximizado
    If Not formFound Then
        Set frmMyForm = New frmFavoritos
        'frmMyForm.MDIChild = True
        frmMyForm.Show
        frmMyForm.WindowState = vbMaximized
    End If

viewBooksFavorites (controlUser)
If Not RsViewBooksFavorites.EOF Then
    Set frmFavoritos.DataGridListFavorites.DataSource = RsViewBooksFavorites
End If
frmFavoritos.EstiloGridBooksFavorites




Unload frmListaUsuarios
Unload frmListLibros
Unload frmReadPdf

frmFavoritos.Show
End Sub

Private Sub mnuLeidos_Click()
    viewBooksRead (controlUser)
    'RsViewBooksRead.MoveFirst
    Unload frmListBooksRead
    Dim frm As Form
    Dim formFound As Boolean
    formFound = False

    ' Buscar si el formulario ya está abierto
    For Each frm In Forms
        If TypeOf frm Is frmListBooksRead Then
            frm.WindowState = vbNormal  ' Restaurar primero
            frm.WindowState = vbMaximized  ' Luego maximizar
            frm.Show
            formFound = True
            Exit For
        End If
    Next frm

    ' Si no está abierto, crear una nueva instancia y mostrarlo maximizado
    If Not formFound Then
        Set frmMyForm = New frmListBooksRead
        'frmMyForm.MDIChild = True
        frmMyForm.Show
        frmMyForm.WindowState = vbMaximized
    End If

viewBooksRead (controlUser)
If Not RsViewBooksRead.EOF Then
    Set frmListBooksRead.DataGridListBooksRead.DataSource = RsViewBooksRead
End If
frmListBooksRead.EstiloGridBooksRead




Unload frmListaUsuarios
Unload frmListLibros
Unload frmReadPdf
Unload frmFavoritos

frmListBooksRead.Show
End Sub

Private Sub mnuMalos_Click()
    viewBooksNG (controlUser)
    'RsViewBooksRead.MoveFirst
    Unload frmListBooksNG
    Dim frm As Form
    Dim formFound As Boolean
    formFound = False

    ' Buscar si el formulario ya está abierto
    For Each frm In Forms
        If TypeOf frm Is frmListBooksNG Then
            frm.WindowState = vbNormal  ' Restaurar primero
            frm.WindowState = vbMaximized  ' Luego maximizar
            frm.Show
            formFound = True
            Exit For
        End If
    Next frm

    ' Si no está abierto, crear una nueva instancia y mostrarlo maximizado
    If Not formFound Then
        Set frmMyForm = New frmListBooksNG
        'frmMyForm.MDIChild = True
        frmMyForm.Show
        frmMyForm.WindowState = vbMaximized
    End If

viewBooksNG (controlUser)
If Not RsViewBooksNG.EOF Then
    Set frmListBooksNG.DataGridListNG.DataSource = RsViewBooksNG
End If
frmListBooksNG.EstiloGridBooksNG




Unload frmListaUsuarios
Unload frmListLibros
Unload frmReadPdf
Unload frmFavoritos
Unload frmListBooksRead

frmListBooksNG.Show
End Sub

Private Sub mnuMiCuenta_Click()
Unload frmListaUsuarios
Unload frmListLibros
Unload frmListBooksNG
Unload frmListBooksRead
Unload frmFavoritos
frmViewUserProfile1.Show vbModal

End Sub

Private Sub mnuSalir_Click()
Unload Me
frmlogin.Show
End Sub

Private Sub mnuUsers_Click()
frmListaUsuarios.Show
End Sub

Private Sub Picture1BgPrincipal_Click()

End Sub

Private Sub Picture1_Click()

End Sub
