Attribute VB_Name = "Module1"
Public MyData As Variant ' These are variant's now you can use As Object, but Not As Database or Recordset or Workspace here now!
Public MYDB As Variant
Public MYTB As Variant
Public Function OpenMYDB() As Long
On Error GoTo openerr
OpenMYDB = -1
Set MYDB = MyData.OpenDatabase(App.Path + "\unprotec.mdb", True, True)
Set MYTB = MYDB.OpenRecordset("TABLE1")
Exit Function
openerr:
OpenMYDB = Err.Number
Err.Clear
End Function

Public Sub Main()
On Error GoTo LoadNextEngine
Set MyData = CreateObject("DAO.DBEngine.35") 'Well Nothing else worked but this one, this is not tested macines below 32meg ram, the memory model may be different on these obselote macines, similar as in IE!
GoTo PassSecond   'No errors, first engine option is available!
LoadNextEngine:  'If first attempt fail's
Err.Clear  'Try next engine version
Set MyData = CreateObject("DAO.DBEngine.36")
PassSecond:
Form1.Show
End Sub
