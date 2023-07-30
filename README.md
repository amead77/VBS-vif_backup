# VBS-vif_backup

 VIF_backup.vbs

 version 2
 
 17/10/17 Adam Mead


 This script is provided "as is" with no warranties expressed or implied.
 
 Written in VBScript, tested on Windows 7 and Windows 10

 Run at the command line or via task scheduler with
 
 cscript VIF_backup /source:<file> /destination:<dir> /delete:14


 What does this do?....
 
 If you have a file that you can't use normal VCS on, say a excel file (or any file that you can't merge)
 
 this will copy the file to a backup directory, while keeping a date stamp of when it was copied.
 
 in the directory you'll have a copy of the file with the date stamp and the original file.
 
 the copy of the original file will be to check for differences, if so it'll copy a new version to the backup directory.
 
 the /delete:14 will delete any files in the backup directory that are older than 14 days. You can use other numbers if 14 isn't to your liking.
 
 requires 2 parameters, /source:<file> /destination:<dir>
 
 /delete is optional

