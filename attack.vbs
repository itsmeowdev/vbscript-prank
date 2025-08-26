X=MsgBox("Error while opening Computer, Do you want to fix it?",4+64,"Computer") 

X=MsgBox("Unable to fix it, Do you want to scan your Computer",3+48,"Computer") 

X=MsgBox("Alert! Virus Detected. Delete Virus?",3+16,"Alert!") 

X=MsgBox("Unable to remove Virus",5+48,"Critical Error") 

X=MsgBox("Virus Activated!!",2+16,"Alert!") 

X=MsgBox("Deleting system files...",2+16,"File Deletion") 

X=MsgBox("Extracting password...",2+48,"Extracting") 

X=MsgBox("Please wait, Uploading your data to server...",1+64,"Copying") 

X=MsgBox("Can't stop, File Transfer Completed.",1+16,"Completed") 

X=MsgBox("Computer Hacked!!!",0+16,"Alert!") 

X=MsgBox("LOL!! it's a Prank",0+64,"Fooled You") 

DIM C
set C=CreateObject("Shell.Application")
C.ShellExecute "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk","https://www.youtube.com/watch?v=xvFZjo5PgG0","","",1