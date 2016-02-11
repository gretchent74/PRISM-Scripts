'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - ADJUSTMENTS.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/theVKC/Anoka-PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")			'Creates an object to get a URL
req.open "GET", url, FALSE						'Attempts to open the URL
req.send									'Sends request
IF req.Status = 200 THEN						'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText					'Executes the script code
ELSE										'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
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

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
checked = 1
unchecked = 0
CAAD_note_check = checked

'Dialogs -------------------------------------------------------------------------------------------------------------------------

'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 85, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

'Adjustment Dialog-
BeginDialog Adjustment_Dialog, 0, 0, 187, 176, "Adjustment(s)"
  Text 0, 10, 70, 10, "Adjustment Reason"
  DropListBox 70, 10, 110, 10, "Please Select One:"+chr(9)+"Arrears Management"+chr(9)+"Direct Support"+chr(9)+"Error"+chr(9)+"Forgiveness"+chr(9)+"Interest Adjustment"+chr(9)+"Order"+chr(9)+"Other"+chr(9)+"Overpayment", Reason_List
  Text 0, 30, 70, 20, "Affected Obligations"
  CheckBox 70, 30, 30, 20, "CCC", CCC_Obli_checkbox
  CheckBox 70, 50, 30, 10, "CCH", CCH_Obli_checkbox
  CheckBox 70, 60, 30, 10, "CMI", CMI_Obli_checkbox
  CheckBox 70, 70, 30, 10, "CMS", CMS_Obli_checkbox
  CheckBox 70, 80, 30, 10, "CSP", CSP_Obli_checkbox
  CheckBox 110, 40, 30, 10, "JCC", JCC_Obli_checkbox
  CheckBox 110, 50, 30, 10, "JCH", JCH_Obli_checkbox
  CheckBox 110, 60, 30, 10, "JME", JME_Obli_checkbox
  CheckBox 110, 70, 30, 10, "JMI", JMI_Obli_checkbox
  CheckBox 110, 80, 30, 10, "JMS", JMS_Obli_checkbox
  CheckBox 150, 40, 30, 20, "Other", Other_Obli_checkbox
  Text 0, 100, 50, 10, "Description"
  EditBox 40, 100, 150, 10, Descrip_Box
  Text 0, 120, 80, 20, "Total Amount Adjusted"
  EditBox 80, 120, 50, 10, Amount_Adjusted
  Text 60, 140, 60, 20, "Worker Signature"
  EditBox 120, 140, 50, 10, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 70, 160, 40, 10
    CancelButton 140, 160, 40, 10
EndDialog

'Connecting to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)

'Case number display dialog
Do
	Dialog case_number_dialog
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
Loop until case_number_valid = True


'Displays dialog for adjustments and checks for information
Do
	
	Do
		Do 	
			Do
				'Shows dialog, validates that PRISM is up and not timed out, with transmit
				Dialog Adjustment_Dialog
				If buttonpressed = 0 then stopscript
				transmit
				EMReadScreen PRISM_check, 5, 1, 36
				If PRISM_check <> "PRISM" then MsgBox "You appear to have timed out, or are out of PRISM. Navigate to PRISM and try again."
			Loop until PRISM_check = "PRISM"
			'Makes sure worker enters in signature
			If Worker_Signature = "" then MsgBox "Sign your CAAD note"
		Loop until Worker_Signature <> ""
		'Makes sure worker selects adjustment reason
		If Reason_List = "Please Select One" then MsgBox "Please enter a reason for the adjustment"
	Loop until Reason_List <> "Please Select One"
	'Make sure worker selects at least one obligation
	If CCH_Obli_checkbox <> checked and CMS_Obli_checkbox <> checked and CMI_Obli_checkbox <> checked and CCC_Obli_checkbox <> checked and CSP_Obli_checkbox <> checked and _
	  JCH_Obli_checkbox <> checked and JMS_Obli_checkbox <> checked and JME_Obli_checkbox <> checked and JCC_Obli_checkbox <> checked and JMI_Obli_checkbox <> checked and _
	  Other_Obli_checkbox <> checked Then MsgBox "You must select an obligation!"
Loop until CCH_Obli_checkbox = checked or CMS_Obli_checkbox = checked or CMI_Obli_checkbox = checked or CCC_Obli_checkbox = checked or CSP_Obli_checkbox = checked or _
  JCH_Obli_checkbox = checked or JMS_Obli_checkbox = checked or JME_Obli_checkbox = checked or JCC_Obli_checkbox = checked or JMI_Obli_checkbox = checked or Other_Obli_checkbox = checked 

'Cleaning up the case note for check boxes	

If CCH_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CCH, ")
If CMS_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CMS, ")
If CMI_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CMI, ")
If CCC_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CCC, ")
If CSP_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CSP, ")
If JCH_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JCH, ")
If JMS_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JMS, ")
If JME_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JME, ")
If JCC_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JCC, ")
If JMI_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JMI, ")
If Other_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("Other")
If right(line_for_CAAD_note, 2) = ", " then line_for_CAAD_note = left(line_for_CAAD_note, len(line_for_CAAD_note) - 2)






'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)


PF5					'Did this because you have to add a new note

EMWriteScreen "FREE", 4, 54   'adds correct caad code 

EMSetCursor 16, 4			'Because the cursor does not default to this location
call write_new_line_in_PRISM_case_note(">>>Adjustments<<<")
call write_editbox_in_PRISM_case_note("Adjustment Reason", Reason_List, 4)
call write_editbox_in_PRISM_case_note("Total Amount Adjusted", "$" & Amount_Adjusted, 4)
call write_editbox_in_PRISM_case_note("Affected Obligations", line_for_CAAD_note, 4) 
call write_editbox_in_PRISM_case_note("Description", Descrip_Box, 4)
call write_new_line_in_PRISM_case_note("---")
call write_new_line_in_PRISM_case_note(Worker_Signature)


