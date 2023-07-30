' VIF_backup.vbs
' version 2
' 17/10/17 Adam Mead
'
'
' This script is provided "as is" with no warranties expressed or implied.
' Written in VBScript, tested on Windows 7 and Windows 10
'
' Run at the command line or via task scheduler with
' cscript VIF_backup /source:<file> /destination:<dir> /delete:14
'
'
' What does this do?....
' If you have a file that you can't use normal VCS on, say a excel file (or any file that you can't merge)
' this will copy the file to a backup directory, while keeping a date stamp of when it was copied.
' in the directory you'll have a copy of the file with the date stamp and the original file.
' the copy of the original file will be to check for differences, if so it'll copy a new version to the backup directory.
' the /delete:14 will delete any files in the backup directory that are older than 14 days. You can use other numbers if 14 isn't to your liking.
' requires 2 parameters, /source:<file> /destination:<dir>
' /delete is optional
'
'
Option Explicit

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

Dim objFolder
Dim objFileCollection
Dim objFile
dim strSourceFile
dim strSourceFnOnly
dim strDestination
dim strSourcePathOnly
dim intDeleteAge
Dim modDate

Main
wscript.echo "How did we get here?"
WScript.quit -1

sub Argue()
    Dim objArgs
    dim colArgs
    Set colArgs = WScript.Arguments.Named

    if colArgs.exists("source") then
        strSourceFile=colArgs.item("source")
    else
        strSourceFile=""
    end if

    if colArgs.exists("destination") then
        strDestination=colArgs.item("destination")
    else
        strDestination=""
    end if
    if colArgs.exists("delete") then
        intDeleteAge=cint(colArgs.item("delete"))
    else
        intDeleteAge=0
    end if
    if ((wscript.arguments.count < 2) or (wscript.arguments.count > 3) or (strSourceFile="") or (strDestination="")) then
        ErrorMsg
    end if
end sub

sub ErrorMsg()
    wscript.echo ""
    wscript.echo ""
    wscript.echo "file backup will copy source file to destination directory."
    wscript.echo "needs 2 parameters:"
    wscript.echo "/source:<SourceFileName> /destination:<DestinationDirectory> /delete:<age in days>"
    wscript.echo "both source file and destination path must include the FULL path, if there is an error you'll see this. /delete is optional"
    wscript.echo ""
    WScript.quit -1    
end sub

sub CheckFileAndDest()
    if not objFS.FileExists(strSourceFile) then
        wscript.echo "Source File Error:  " & strSourceFile
        ErrorMsg
    end if 
    wscript.echo "Source File Exists: " & strSourceFnOnly
    if right(strDestination,1) <> "\" then strDestination=strDestination & "\"
    if not objFS.FolderExists(strDestination) then
        wscript.echo "Destination Error:  " & strDestination
        ErrorMsg
    end if     
    wscript.echo "Destination Exists: " & strDestination
    
    set objFile=objFS.getfile(strSourceFile)
    strSourceFnOnly=objFS.getfilename(objFile)
    strSourcePathOnly=objFS.GetParentFolderName(objFile)    
    if right(strSourcePathOnly,1) <> "\" then strSourcePathOnly=strSourcePathOnly & "\"    
end sub

Function CompareFiles(sfile, dfile)
	dim tmp
	dim objfile1
    dim objFile2

	If not objFS.FileExists(dfile) Then
		'wscript.echo "Destination master is missing, copying new"
		CompareFiles = 0
		exit function
	End If

	if tmp = 0 then
		wscript.echo "Source:		"+sfile
		wscript.echo "Destination:	"+dfile
		Set objFile1 = objFS.GetFile(sfile)
		Set objFile2 = objFS.GetFile(dfile)
		If objFile1.DateLastModified > objFile2.DateLastModified Then
			tmp = 0
		else
			tmp = 1
		End If
	end if
	CompareFiles = tmp

End Function

' Check the number of days is 1 or greater (otherwise it will just delete everything)
function DelOlder(strDirectoryPath, intDaysOld)
    Dim objDir
    Dim objFCollection
    Dim objF
    If (intDaysOld>0) Then 
        wscript.echo "Delete files more than " & intDaysOld & " days old."

        If (IsNull(strDirectoryPath)) Then 
            wscript.echo "error in path, quitting"
            WScript.quit -1
        end if
        wscript.echo "Delete from: " & strDirectoryPath
        wscript.echo ""


        set objDir = objFS.GetFolder(strDirectoryPath)
        set objFCollection = objDir.Files

        For each objF in objFCollection
            If objF.DateLastModified < (Date() - intDaysOld) Then
                    Wscript.Echo "DELETE-> " & objF.Name & " " & objF.DateLastModified
                    objF.Delete(True)
            'To delete for real, remove the ' from the line above
            End If
        Next
    end if
end function

sub Main
    
    Argue
    CheckFileAndDest
    call DelOlder(strDestination,intDeleteAge)

    if CompareFiles(strSourceFile,strDestination & strSourceFnOnly) = 0 then
        wscript.echo "Either the Source is newer or destination is missing, copying"
        objFS.CopyFile strSourceFile, strDestination, true
        modDate = DatePart("yyyy", now) & "-" & Right("0" & DatePart("m",now), 2) & "-" &  Right("0" & DatePart("d",now), 2) & "_" & Right("0" & DatePart("h",now), 2) & Right("0" & DatePart("n",now), 2)
        modDate = modDate & " " & strSourceFnOnly
        wscript.echo "Creating:	" +modDate
        objFS.CopyFile strSourceFile, strDestination & "\" & modDate, true
    else
        wscript.echo "Either the source is the same age or the destination doesn't exist, doing nothing"
    end if

    wscript.quit 1
end sub
