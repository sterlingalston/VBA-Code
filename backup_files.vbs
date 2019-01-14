'FileSystemObject.CopyFile "c:\temp\test_export.csv", "N:\Leasing_Real Estate Admin\Business Analytics\MatthewAlston\test_export.csv",true

Option Explicit

Call BackupMatthew

Public Sub BackupMatthew()



dim fso, rootdir, targdir, tempdir, targdir2,f
dim oShell


'set Wscript = CreateObject("WScript.Shell")
set fso = CreateObject("Scripting.FileSystemObject")
set oShell = WScript.CreateObject("WScript.Shell") 
'on error resume next

rootdir = "C:\Users\matthewsalston\Documents\*"
targdir = "N:\Leasing_Real Estate Admin\Business Analytics\MatthewAlston\*"
targdir2 = "h:\*"

'targdir = "C:\Users\matthewsalston\Desktop\Dir1\*"
'targdir2 = "C:\Users\matthewsalston\Desktop\Dir2\*"

set tempdir = fso.getFolder(left(rootdir,len(rootdir)-1))

'fso.CopyFolder rootdir,targdir,true
'fso.CopyFolder rootdir,targdir2,true

oShell.Run "xcopy " & """" & rootdir & """" & " " & """" & targdir   & """" & "/Z /Y /V /S /E"
oShell.Run "xcopy " & """" & rootdir & """" & " " & """" & targdir2   & """" & "/Z /Y /V /S /E"

for each f in tempdir.subfolders
	 oShell.Run "xcopy " & """" & left(rootdir,len(rootdir)-1) & f.name &"\*" & """" & " " & """" & left(targdir,len(targdir)-1) & f.name &"\*"   & """" & "/Z /Y /V /S /E"
	 oShell.Run "xcopy " & """" & left(rootdir,len(rootdir)-1) & f.name &"\*" & """" & " " & """" & left(targdir2,len(targdir2)-1) & f.name &"\*"   & """" & "/Z /Y /V /S /E"
next 

End Sub