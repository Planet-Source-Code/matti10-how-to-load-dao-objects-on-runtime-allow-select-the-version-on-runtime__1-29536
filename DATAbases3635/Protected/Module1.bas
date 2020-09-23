Attribute VB_Name = "Module1"
Public MyData As Variant ' These are variant's now you can use As Object, but Not As Database or Recordset or Workspace here now!
Public MYDB As Variant
Public MYTB As Variant
Public MYWR As Variant
Public Function OpenMYDB() As Long
On Error GoTo openerr
OpenMYDB = -1
Set MYDB = MYWR.OpenDatabase(App.Path + "\protectd.mdb", True, True)
Set MYTB = MYDB.OpenRecordset("TABLE1")
Exit Function
openerr:
OpenMYDB = Err.Number
Err.Clear
End Function

Public Sub Main()
Dim MyPWL$, MyUser$
'MyUser = "Admin"   'Admin account
'MyPWL = "admin"
MyUser = "MattiA"   'users account
MyPWL = "masaA"

On Error GoTo LoadNextEngine
Set MyData = CreateObject("DAO.DBEngine.35") 'the database need to be in Acc97 format to use this one
MyData.SystemDB = App.Path + "\SYSTEM.MDW" 'This DB is protected, so need a way to insert password and user!
Set MyData = CreateObject("DAO.DBEngine.35") 'Well Nothing else worked but this one, this is not tested macines below 32meg ram, the memory model may be different on these obselote macines, similar as in IE!
Set MYWR = MyData.CreateWorkspace("", MyUser, MyPWL, 2) 'Normal DAO Constant's not available whitout references, need them make them: Public Const MyDBUSEJET = 2
GoTo PassSecond   'No errors, first engine option is available!
LoadNextEngine:  'If first attempt fail's
Err.Clear  'Try next engine version
Set MyData = CreateObject("DAO.DBEngine.36")
MyData.SystemDB = App.Path + "\SYSTEM.MDW" 'This DB is protected, done very lousy, left the valid user ("") account there in case you can't set the system.mdw in Access, Beginers! Don't try to do that if you not familar whit Access!
Set MyData = CreateObject("DAO.DBEngine.36")
Set MYWR = MyData.CreateWorkspace("", MyUser, MyPWL, 2) 'Constant's not available whitout references, need them make them: Public Const MyDBUSEJET = 2
PassSecond:
Form1.Show
End Sub
