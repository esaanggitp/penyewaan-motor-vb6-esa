Attribute VB_Name = "Module1"
'MENDEFINISIKAN OBJEK
Public KON As New ADODB.Connection
Public rsmotor As ADODB.Recordset
Public rskaryawan As ADODB.Recordset
Public rspelanggan As ADODB.Recordset
Public rsdetail As ADODB.Recordset
Public rstrans As ADODB.Recordset
Public rstemp As ADODB.Recordset
Sub koneksi()
'MEMBUAT OBJEK
Set KON = New ADODB.Connection
Set rsmotor = New ADODB.Recordset
Set rskaryawan = New ADODB.Recordset
Set rspelanggan = New ADODB.Recordset
Set rsdetail = New ADODB.Recordset
Set rstrans = New ADODB.Recordset
Set rstemp = New ADODB.Recordset
KON.ConnectionString = "driver=mysql odbc 3.51 driver;server=127.0.0.1;uid=root;db=rental;"
KON.Open
End Sub

