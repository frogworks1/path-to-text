Set objFS = CreateObject("Scripting.FileSystemObject")
strFolder = "O:\Biostatistics\SGN-35\sg035-0003\ASH2015abstract\outputs\tlfs\testing"
strDestination = "C:\Users\tkelly\Desktop\github\frogworks1.github.io\path-to-text\output.txt"
Set objFolder = objFS.GetFolder(strFolder)

Go(objFolder)

Sub Go(objDIR)
  If objDIR <> "\System Volume Information" Then
    For Each eFolder in objDIR.SubFolders   	
        Go eFolder
    Next
    For Each strFile In objDIR.Files
        strFileName = strFile.Name
        strExtension = objFS.GetExtensionName(strFile)
        If strExtension = "sas" Then
        	objFS.CopyFile strFile , strDestination & strFileName
        End If 
    Next    
  End If  
End Sub