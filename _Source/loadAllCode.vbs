Set objExcel = CreateObject("Excel.Application")
set objShell = CreateObject("WScript.Shell")
currentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
Set objWorkbook = objExcel.Workbooks.Open(currentDirectory & "\_CPNM.xlam")

objExcel.Application.Visible = True
objExcel.Workbooks.Add
objExcel.Application.Run "codeLoad"

saveaddress = currentDirectory & "_CPNM_new.xlam"
objWorkbook.Saveas saveaddress, 55

objExcel.Application.Quit

on error resume next
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\_CPNM.xlam"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\_CPNM.dotm"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\_CPNM.dvb"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\sourceTools.xla"
objShell.Run "git update-index --assume-unchanged " & currentDirectory & "\sourceTools.vst"
on error goto 0