Attribute VB_Name = "modRegFree"
'@IgnoreModule ModuleWithoutFolder
Option Explicit 'simple "drop-in" regfree-module for RC6 (in the IDE it will use the registered version, but as Exe will require a \Bin\-Subfolder with all RC6-Dlls)

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetInstanceEx Lib "DirectCOM" (spFName As Long, spClassName As Long, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object

Private Const DirectComDllRelPath As String = "\Bin\DirectCOM.dll"
Private Const RCDllRelPath  As String = "\Bin\RC6.dll"

Public Property Get New_c() As cConstructor
  Static st_RC As cConstructor
  If Not st_RC Is Nothing Then Set New_c = st_RC: Exit Property
  
  If App.LogMode Then 'we run compiled - and try to ensure regfree instantiation from \Bin\
     On Error Resume Next
        LoadLibraryW StrPtr(App.path & DirectComDllRelPath)
        Set st_RC = GetInstanceEx(StrPtr(App.path & RCDllRelPath), StrPtr("cConstructor"))
        If st_RC Is Nothing Then MsgBox "Couldn't load regfree... (\Bin-folder missing?)" & vbLf & vbLf & "Will try with a registered version next..."
     On Error GoTo 0
  End If
  If st_RC Is Nothing Then Set st_RC = cGlobal.New_c 'fall back to loading a registered version
  Set New_c = st_RC
End Property

Public Property Get Cairo() As cCairo
  Static st_CR As cCairo
  If st_CR Is Nothing Then Set st_CR = New_c.Cairo
  Set Cairo = st_CR
End Property
 
Public Function New_W(ClassName As String, Optional DllName As String = "RC6Widgets.dll") As Object
   If App.LogMode <> 0 And StrComp(App.path & "\Bin\", New_c.LibPath, vbTextCompare) = 0 Then 'when running compiled - and both paths match...
      On Error Resume Next '...we try regfree instantiation from \Bin\ first
         Set New_W = New_c.RegFree.GetInstanceEx(App.path & "\Bin\" & DllName, ClassName)
      On Error GoTo 0
   End If
   If New_W Is Nothing Then Set New_W = CreateObject(Replace(DllName, "dll", "", 1, 1, vbTextCompare) & ClassName) 'the fallback to a registered version
End Function
 
