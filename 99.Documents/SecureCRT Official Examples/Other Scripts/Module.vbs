
' Subroutine Doit() is written as a module that can be shared by inclusion
' in more than one script. The file CallModule.vbs shows the VBScript syntax
' that allows this file to be included and used by other scripts.
'
Sub Doit(s)
  MsgBox s
End Sub

