objArgs = WScript. Arguments  
var a=[]
//objArgs.count  
var fso=new ActiveXObject("scripting.filesystemobject")
//Set f=fso.GetFolder(fso.GetAbsolutePathName("."))
var f=fso.GetFolder(fso.GetAbsolutePathName(objArgs(0)));
WScript.Echo(getFolders(f.path));
WScript.Echo(a);
folders=null;
f=null;
fso=null;

function getFolders(path){
    var t = "";
    var arr=[];
    var op;
    var ofolder=fso.GetFolder(path);
    var subfolders=new Enumerator(ofolder.subfolders);
    for(;!subfolders.atEnd();subfolders.moveNext()){
        //arr[arr.length]=subfolders.item().path;
        //arr[arr.length]=getFolders(subfolders.item().path);
        t=t + "\n" + subfolders.item().path;
        t=t + "\n" + getFolders(subfolders.item().path);
    }
    var subfiles=new Enumerator(ofolder.Files);
    for(;!subfiles.atEnd();subfiles.moveNext()){ 
        //arr[arr.length]=subfiles.item().Name;
        t=t + "\n" +  subfiles.item().Name;
    }
    return t;
}
