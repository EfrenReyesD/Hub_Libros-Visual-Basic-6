Attribute VB_Name = "ModuloDeclaraciones"
'Variable de coneccion a base de datos
Global Base As New ADODB.Connection 'declarando una variable ADO para coneccion

'Variables RecordSet

Global RsTablaUsers As New ADODB.Recordset
Global RsViewUsers As New ADODB.Recordset
Global RsAllBooks As New ADODB.Recordset
Global RsViewBooks As New ADODB.Recordset
Global RsUserBooks As New ADODB.Recordset
Global RsViewBooksFavorites As New ADODB.Recordset
Global RsViewBooksRead As New ADODB.Recordset
Global RsViewBooksNG As New ADODB.Recordset
Global RsViewUserFilter As New ADODB.Recordset
Global RsViewBooksFilter As New ADODB.Recordset



Public controlUser As String
Public controlNameUser As String
Public controlIdBookSelectView As Integer


