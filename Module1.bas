Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rst As New ADODB.Recordset
Public strcon As String
Public strsql As String

Public Sub buka()
On Error GoTo pesan
strcon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database1.mdb;Persist Security Info=False"
If con.State = adStateOpen Then
con.Close
Set con = New ADODB.Connection
con.Open strcon
Else
con.Open strcon
End If
Exit Sub
pesan:
MsgBox "Tidak ada koneksi ke database..!", vbInformation, "Informasi"
End Sub

Public Sub tutup()
con.Close
End Sub

