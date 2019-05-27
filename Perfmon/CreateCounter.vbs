'This script capture user input then create performance counter captures using four configuration templates

'TO DO:
'Set start and end times

Option Explicit


'This is required to tell the file system object to open the file for reading
Const ForReading = 1


'Declare the variables required for this script
dim WshShell, ServerName, sCounterName, oFS, oFile, oFile1
dim sInfo, sFileName, sLine, sUsr, sPwd, sLoc, sFileFormat
dim iQues, oWorkingDir


'Create a windows scripting shell object, primarily used to run command line apps and arguments.
Set WshShell = CreateObject("Wscript.shell")


'Create a file System object, so we can read and write to text files
Set oFS = CreateObject("Scripting.FileSystemObject")

oWorkingDir = oFS.GetAbsolutePathName("")

wscript.echo oWorkingDir


'This function will run programs and arguments on the command line
'it returns a errorlevel or 0 or 1
'0 is good
'1 is bad
Function CommandShell(sCli)

	Dim iErr

	iErr = Wshshell.run(sCli, 7, true)

	CommandShell = iErr

End Function


'Open config file and write server name to the start of a new template
'the new template will be used to create the performance counters
Sub WriteToConf(sInfo, sFileName)
	
	Set oFile1 = oFS.CreateTextFile(oWorkingDir & "\" & sInfo & sFileName & ".cfg")

	Set oFile = oFS.openTextFile(oWorkingDir & "\" & sFileName & "Template.cfg", ForReading)

	Do until oFile.AtEndOfStream	

		sLine = oFile.Readline
		
		sLine = sFileFormat & sLine			
				
		oFile1.Writeline(sLine)

	Loop

	oFile1.Close
	
	oFile.Close


	If sLoc = "l" Then
		CommandShell "logman create counter " & """" & sCounterName & " " & sFilename & """" & " -f bincirc -max 200 -si 00:15:00 -v -o """ & oWorkingDir & "\Localhost\" & sFileName & """ -cf """ & oWorkingDir & "\" & sInfo & sFileName & ".cfg"""
	Else
		CommandShell "logman create counter " & """" & sCounterName & " " & sFilename & """" & " -f bincirc -max 200 -si 00:15:00 -v -o """ & oWorkingDir & "\" & ServerName & "\" & sFileName & """ -cf """ & oWorkingDir & "\" & sInfo & sFileName & ".cfg"" -u " & sUsr & " " & sPwd
	End If
 
iQues = MsgBox("Do you want to start the counter " & sCounterName & " " & sFilename, 4+32+0)

Select Case iQues
	
	Case 6
		CommandShell "logman.exe start " & """" & sCounterName & " " & sFilename & """"

	Case 7
		wscript.echo "To start the counters open a command window and type logman.exe start " & """" & sCounterName & " " & sFilename & """"
		
End Select

End Sub




'********************** START *****************************************************************************

'Pass whatever the user types to the respective variable
sLoc = InputBox("Is the server local or remote; type l for local and r for remote")

sLoc = Trim(sLoc)

sLoc = LCase(sLoc)

If sLoc = "l" then

	'No need to ask for a username
	sCounterName = InputBox("What do you what the counter to be called?")
	
	ServerName = "LocalHost"

	sFileFormat = """"

Elseif sLoc = "r" then

	ServerName = InputBox("Enter the server name")

	sUsr = InputBox("Which user should run the performance counters? The format should be domain\username")
	sPwd = InputBox("Type that users password")

	sCounterName = InputBox("What do you what the counter to be called?")

	sFileFormat = """\\" & ServerName

Else

	'They entered an incorrect value
	wscript.echo "You have entered an incorrect value in the location input box; please enter a r for remote captures or a l for local captures"
	wscript.quit

End If

WriteToConf servername, "DiskIO"
WriteToConf servername, "Memory"
WriteToConf servername, "NetworkIO"
WriteToConf servername, "processor"


'************************* Future Dev *********************************************
'starttime = InputBox("When should the counter start?","","DD/MM/YYYY hh:mm:ss")
'endtime = InputBox("When should the counter end?","","DD/MM/YYYY hh:mm:ss")
'************************* Future Dev *********************************************

Set sUsr = Nothing
Set sPwd = Nothing


'********************** FINISH *****************************************************************************