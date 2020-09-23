Attribute VB_Name = "Module1"
'public declarations
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rsbill As New ADODB.Recordset
Public Constring As String
Dim btb As String
'function main to be used all over to declare db connection from this module1
Sub Main()

'opens connection to the database(pharmacy1)
Constring = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Pharmacy1.mdb"


End Sub
