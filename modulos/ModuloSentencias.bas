Attribute VB_Name = "ModuloSentencias"
Sub Main()
'conectar a la base de datos
With Base
.CursorLocation = adUseClient 'ser cliente de la BD
.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=hublibros;Data Source=LAPTOP-QRHFTO8H\SQLEXPRESS"
frmlogin.Show
End With
End Sub

'Conectores a tablas independientes

Sub AbrirTablaUsers()
With RsTablaUsers
If .State = 1 Then .Close 'si esta abierto que se cierre
.Open "Select * from Users", Base, adOpenKeyset, adLockPessimistic ' adOpenStatic, adLockBatchOptimistic
End With
End Sub

'Solo obtener ciertos datos de la db
Sub viewUsers()
With RsViewUsers
If .State = 1 Then .Close 'si esta abierto que se cierre
.Open "select UserId, UserName, FirstName, LastName, CASE WHEN Admin=1 THEN 'Admin' ELSE 'Cliente' END AS TipoUsuario  from Users", Base, adOpenKeyset, adLockPessimistic
End With
End Sub

'Obtener datos de la tabla libros
Sub abrirAllBooks()
With RsAllBooks
If .State = 1 Then .Close 'si esta abierto que se cierre
.Open "select *  from Books", Base, adOpenKeyset, adLockPessimistic
End With
End Sub
'Funcion para obtener los datos de la tabla pivote userBooks
Sub abrirUserBooks()
With RsUserBooks
If .State = 1 Then .Close 'si esta abierto que se cierre
.Open "select *  from UserBooks", Base, adOpenKeyset, adLockPessimistic
End With
End Sub

'Funcion para mostrar todos los libros
Sub viewBooks(ByVal User As Integer)
With RsViewBooks
If .State = 1 Then .Close 'si esta abierto que se cierre
.Open "SELECT b.BookId, b.Title, b.Author, b.Genre, " & _
             "CASE WHEN (ub.IsRead = 1 AND ub.IsDisliked=1) THEN 'Leído y No Le Gustó'" & _
             "WHEN (ub.IsRead = 1 AND ub.IsFavorite=1) THEN 'Leído y Favorito'" & _
             "WHEN ub.IsRead = 1 THEN 'Leído' " & _
             "WHEN ub.IsDisliked = 1 THEN 'No le gustó' " & _
             "WHEN ub.IsFavorite = 1 THEN 'Favorito'" & _
             "ELSE 'No leído' END AS Status " & _
             "FROM Books b " & _
             "LEFT JOIN UserBooks ub ON b.BookId = ub.BookId AND ub.UserId = " & User & "", Base, adOpenKeyset, adLockPessimistic
End With
End Sub


'Funcion para mostrar todos los libros FAVORITOS
Sub viewBooksFavorites(ByVal User As Integer)
    With RsViewBooksFavorites
    If .State = 1 Then .Close 'si esta abierto que se cierre
        .Open "SELECT b.BookId, b.Title, b.Author, b.Genre " & _
                     "FROM Books b " & _
                     "JOIN UserBooks ub ON b.BookId = ub.BookId " & _
                     "WHERE ub.UserId = " & User & " AND ub.IsFavorite = 1", Base, adOpenKeyset, adLockPessimistic
    End With
End Sub


'Funcion para mostrar todos los libros LEÍDOS
Sub viewBooksRead(ByVal User As Integer)
    With RsViewBooksRead
    If .State = 1 Then .Close 'si esta abierto que se cierre
        .Open "SELECT b.BookId, b.Title, b.Author, b.Genre " & _
                     "FROM Books b " & _
                     "JOIN UserBooks ub ON b.BookId = ub.BookId " & _
                     "WHERE ub.UserId = " & User & " AND ub.IsRead = 1", Base, adOpenKeyset, adLockPessimistic
    End With
End Sub






'Funcion para mostrar todos los libros NG
Sub viewBooksNG(ByVal User As Integer)
    With RsViewBooksNG
    If .State = 1 Then .Close 'si esta abierto que se cierre
        .Open "SELECT b.BookId, b.Title, b.Author, b.Genre " & _
                     "FROM Books b " & _
                     "JOIN UserBooks ub ON b.BookId = ub.BookId " & _
                     "WHERE ub.UserId = " & User & " AND ub.IsDisliked = 1", Base, adOpenKeyset, adLockPessimistic
    End With
End Sub

'Funcion que nos ayudará a filtrar Usuarios asi como se escriba en el textBOx
Sub viewUsersFilter(ByVal columna As String, ByVal filtro As String)
    With RsViewUserFilter
    If .State = 1 Then .Close 'si esta abierto que se cierre
        .Open "select UserId, UserName, FirstName, LastName, CASE WHEN Admin=1 THEN 'Admin' ELSE 'Cliente' END AS TipoUsuario  from Users WHERE [" & columna & "] LIKE '%" & filtro & "%'", Base, adOpenKeyset, adLockPessimistic
    End With
End Sub




'Funcion filtrar todos los libros
Sub viewBooksFilter(ByVal User As Integer, ByVal filtro As String)
With RsViewBooksFilter
If .State = 1 Then .Close 'si esta abierto que se cierre
.Open "SELECT b.BookId, b.Title, b.Author, b.Genre, " & _
          "CASE WHEN ub.IsRead = 1 THEN 'Leído' " & _
          "WHEN ub.IsDisliked = 1 THEN 'No le gustó' " & _
          "ELSE 'No leído' END AS Status " & _
          "FROM Books b " & _
          "LEFT JOIN UserBooks ub ON b.BookId = ub.BookId AND ub.UserId = " & User & " " & _
          "WHERE b.Title LIKE '%" & filtro & "%'", Base, adOpenKeyset, adLockPessimistic
End With
End Sub






