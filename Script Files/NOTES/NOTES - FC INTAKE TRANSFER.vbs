Option Explicit  'this has to be on the top and be sure to use it!

'this is a function document from page 5 custom functions library, copied and pasted
DIM beta_agency 'remember to add

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF
'this is where the copy and paste from functions library ended


'added DIM to declare my info from the dialog
DIM FC_Intake_Transfer_dialog, Prism_Case_Number, Underlying_Case_Number, file_scanned_checkbox, volunary_agreement_order_checkbox,Worker_Name, ButtonPressed


'added my dialog here using copy and past from dialog edit
BeginDialog FC_Intake_Transfer_dialog, 0, 0, 206, 155, "FC Intake Transfer"
  EditBox 80, 5, 90, 15, Prism_Case_Number
  EditBox 95, 25, 85, 15, Underlying_Case_Number
  CheckBox 10, 50, 90, 10, "File Scanned to Onbase", file_scanned_checkbox
  CheckBox 10, 65, 195, 10, "Voluntary Placement Agreement/Juvenile Order received", volunary_agreement_order_checkbox
  EditBox 60, 105, 95, 15, Worker_Name
  ButtonGroup ButtonPressed
    OkButton 95, 130, 50, 15
    CancelButton 150, 130, 50, 15
  Text 10, 10, 70, 10, "FC Case Number"
  Text 10, 30, 85, 10, "Underlying Case Number"
  Text 10, 90, 75, 10, "Transferred to FC-CSS"
  Text 10, 110, 45, 10, "Completed by"
EndDialog

'added to the dialog I cut and pasted 
'Dialog FC_Intake_Transfer_dialog
'If ButtonPressed = 0 THEN StopScript

'adding a loop 
DO
	Dialog FC_Intake_Transfer_dialog
	If ButtonPressed = 0 THEN StopScript 'this is for the cancel button to actuall work
	If Prism_Case_Number = "" THEN MsgBox "Prism Case number must be Completed"
	If Underlying_Case_Number = "" THEN MsgBox "Underlying Case number must be Completed"
	If Worker_Name = "" THEN MsgBox "Worker Name must be Completed"
LOOP UNTIL Prism_Case_Number <> "" and Underlying_Case_Number <> "" and Worker_Name <> ""


'connecting to bluezone
EMConnect ""

'to pull up my prism 
EMFocus

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

'brings me to caad
CALL navigate_to_PRISM_screen ( "CAAD")

'to create a caad note
PF5
EMWriteScreen "A", 3, 29

'sets caad note type free
EMWriteScreen "free", 4, 54

'set the cursor on the caad note area
EMSetCursor 16, 4

'added so this info from the dialog will go in caad note
CALL write_variable_in_CAAD("FC Intake Case Transfer")
CALL write_bullet_and_variable_in_CAAD("FC Case Number", Prism_Case_Number & "   Underlying Case Number: " & Underlying_Case_Number)
IF file_scanned_checkbox = checked THEN call write_variable_in_CAAD ("*Filed Scanned to Onbase.")
IF volunary_agreement_order_checkbox = checked THEN call write_variable_in_CAAD ("*Voluntary Placement Agreement/Juvenile Order received.")
CALL write_variable_in_CAAD("Transferred to FC-CSS")
CALL write_variable_in_CAAD(Worker_Name)
