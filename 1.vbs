Set objArgs = WScript. Arguments  
dim t 
dim a(5)
a(1)=2
'objArgs.count  
Set fso=CreateObject("scripting.filesystemobject")
'Set f=fso.GetFolder(fso.GetAbsolutePathName("."))
Set f=fso.GetFolder(fso.GetAbsolutePathName(objArgs(0)))
WScript.Echo getFolders(f.path)
WScript.Echo a
Set folders=Nothing
Set f=nothing
Set fso=nothing

Function getFolders(path)
    dim t 
    Set ofolder=fso.GetFolder(path)
    For Each uu In ofolder.subfolders
        Set op=fso.GetFolder(uu.path)
        t=t & vbcrlf & op.path
        t=t & vbcrlf & getFolders(op.path)
    Next
  
    For Each f In ofolder.Files 
        t=t & vbcrlf &  f.Name
    Next
    getFolders = t
End function

