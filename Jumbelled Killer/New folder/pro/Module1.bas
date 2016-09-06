Attribute VB_Name = "Module1"
Public con As ADODB.Connection
Public cmd As ADODB.Recordset
Public atn As ADODB.Recordset
Public rreg As ADODB.Recordset

Public Sub connect()

Set con = New ADODB.Connection
Set cmd = New ADODB.Recordset
Set atn = New ADODB.Recordset
Set rreg = New ADODB.Recordset

con.Provider = "Microsoft.Jet.OLEDB.4.0;Password="
con.Open "Data Source=D:\THARUN KING\PROHIBITED AREA-passers not alowed beyound this\project\thapropro.mdb;Persist Security Info=True"


cmd.Open "select * from tharunpro", con, adOpenForwardOnly, adLockReadOnly

atn.Open "select * from tharunpro", con, adOpenDynamic, adLockPessimistic

rreg.Open "select * from reg", con, adOpenDynamic, adLockPessimistic



End Sub

