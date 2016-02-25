'GATHERING STATS----------------------------------------------------------------------------------------------------

'name_of_script = "ACTIONS -PIN #.vbs"
'start_time = timer


'this is a function document
DIM beta_agency 'remember to add

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO                                                                          'Declares variables to be good to option explicit users
If beta_agency = "" then                                              'For scriptwriters only
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then                 'For beta agencies and testers
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else                                                                                                                        'For most users
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")                                                               'Creates an object to get a URL
req.open "GET", url, False                                                                                                                                            'Attempts to open the URL
req.send                                                                                                                                                                                                              'Sends request
If req.Status = 200 Then                                                                                                                                                '200 means great success
                Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
                Execute req.responseText                                                                                                                          'Executes the script code
ELSE                                                                                                                                                                                                                       'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
                MsgBox                "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
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

'-------is this needed

checked = 1
unchecked = 0

'DIALOG---------------------------------------------------------------------------------------------
DIM PIN, PRISM_case_number, CP_PIN_Notice_check, NCP_PIN_Notice_check, CP_MNOnlineEDAK_check, NCP_MNOnlineEDAK_check,CAAD_note, ButtonPressed, Initials_CAAD, err_msg

BeginDialog PIN, 0, 0, 191, 140, "Pin and/or Participant Number Request"
  EditBox 55, 5, 100, 15, PRISM_case_number
  CheckBox 15, 40, 70, 10, "PIN DORD F0999", CP_PIN_Notice_check
  CheckBox 15, 50, 70, 10, "MNOnline EDAK", CP_MNOnlineEDAK_check
  CheckBox 120, 40, 65, 10, "PIN Dord F0999", NCP_PIN_Notice_check
  CheckBox 120, 50, 65, 10, "MNOnline EDAK", NCP_MNOnlineEDAK_check
  CheckBox 5, 75, 130, 10, "Check here if you want a CAAD note.", CAAD_note
  EditBox 80, 90, 35, 15, Initials_CAAD
  ButtonGroup ButtonPressed
    OkButton 75, 120, 50, 15
    CancelButton 135, 120, 50, 15
  Text 5, 10, 45, 10, "Case Number"
  Text 5, 30, 45, 10, "Letters to CP"
  Text 110, 30, 50, 10, "Letters to NCP"
  Text 5, 95, 70, 10, "Initials for CAAD note:"
EndDialog

'END DIALOG----------------------------------------------------------------------------------------------------


'connecting to bluezone
EMConnect ""

'to pull up my prism 
EMFocus

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

'brings me to the CAPS screen to auto fill prims case number in dialog
CALL navigate_to_PRISM_screen ("CAPS")
EMReadScreen PRISM_case_number, 13, 4, 8 




'THE LOOP--------------------------------------------------------------------------
Do	
	err_msg = ""
	Dialog PIN				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF CP_PIN_Notice_check = 0 AND NCP_PIN_Notice_check = 0 AND NCP_MNOnlineEDAK_check = 0 AND CP_MNOnlineEDAK_check = 0 THEN err_msg = err_msg & vbNewline & "You must select at least one document to send out!"
		IF CAAD_note = 1 and Initials_CAAD = "" THEN MsgBox "Please sign your CAAD Note."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

Loop until err_msg = ""

'END LOOP-------------------------------------------------------------------------

'CUSTOM FUNCTIONS***************************************************************************************************************

' This is a custom function to change the format of a participant name.  The parameter is a string with the 
' client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".  

FUNCTION change_client_name_to_FML(client_name)
	client_name = trim(client_name)
	length = len(client_name)
	position = InStr(client_name, ", ")
	last_name = Left(client_name, position-1)
	first_name = Right(client_name, length-position-1)	
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
END FUNCTION

'This is a custom function to fix data that we are reading from PRISM that includes underscores.  The parameter is a string for the 
'variable to be searched.  The function searches the variable and removes underscores.  Then, the fix case function is called to format
'the string in the correct case.  Finally, the data is trimmed to remove any excess spaces.	
FUNCTION fix_read_data (search_string) 
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
END FUNCTION

'END CUSTOM FUNCTIONS-------------------------------------------------------
DIM worker_name, worker_phone
'Getting worker info for case note
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
EMReadScreen worker_phone, 12, 8, 35
PF3

'Cleaning up worker info
worker_name = trim(worker_name)
call fix_case(worker_name, 1)
worker_name = change_client_name_to_FML(worker_name)

DIM NCP_MCI, CP_MCI, NCP_F, NCP_M, NCP_L, NCP_Name, ncp_street_line_1, ncp_street_line_2, ncp_city_state_zip, ncp_address
'readS mci for word and dord docs
EMReadScreen NCP_MCI, 10, 8, 11
EMReadScreen CP_MCI, 10, 4, 8 

'NCP Name
call navigate_to_PRISM_screen("NCDE")
EMWriteScreen NCP_MCI, 4, 7
EMReadScreen NCP_F, 12, 8, 34
EMReadScreen NCP_M, 12, 8, 56
EMReadScreen NCP_L, 17, 8, 8

NCP_name = fix_read_data(NCP_F) & " " & fix_read_data(NCP_M) & " " & fix_read_data(NCP_L)	
NCP_name = trim(NCP_name)

'NCP Address
'Navigating to NCDD to pull address info
call navigate_to_PRISM_screen("NCDD")
EMReadScreen ncp_street_line_1, 30, 15, 11
EMReadScreen ncp_street_line_2, 30, 16, 11
EMReadScreen ncp_city_state_zip, 49, 17, 11

'Cleaning up address info
ncp_street_line_1 = replace(ncp_street_line_1, "_", "")
call fix_case(ncp_street_line_1, 1)
ncp_street_line_2 = replace(ncp_street_line_2, "_", "")
call fix_case(ncp_street_line_2, 1)
if trim (ncp_street_line_2) <> "" then
	ncp_address = ncp_street_line_1 & chr(13) & ncp_street_line_2
else
	ncp_address = ncp_street_line_1
end if
ncp_city_state_zip = replace(replace(replace(ncp_city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(ncp_city_state_zip, 2)

DIM CP_F, CP_M, CP_L, CP_Name, cp_street_line_1, cp_street_line_2, cp_city_state_zip, cp_address
'CP Name											
call navigate_to_PRISM_screen("CPDE")
EMWriteScreen CP_MCI, 4, 7
EMReadScreen CP_F, 12, 8, 34
EMReadScreen CP_M, 12, 8, 56
EMReadScreen CP_L, 17, 8, 8

CP_name = fix_read_data(CP_F) & " " & fix_read_data(CP_M) & " " & fix_read_data(CP_L)	
CP_name = trim(CP_Name)

'CP Address
'Navigating to CPDD to pull address info
call navigate_to_PRISM_screen("CPDD")
EMReadScreen cp_street_line_1, 30, 15, 11
EMReadScreen cp_street_line_2, 30, 16, 11
EMReadScreen cp_city_state_zip, 49, 17, 11

'Cleaning up address info
cp_street_line_1 = replace(cp_street_line_1, "_", "")
call fix_case(cp_street_line_1, 1)
cp_street_line_2 = replace(cp_street_line_2, "_", "")
if trim (cp_street_line_2) <> "" then
	cp_address = cp_street_line_1 & chr(13) & cp_street_line_2
else
	cp_address = cp_street_line_1
end if
call fix_case(cp_street_line_2, 1)
cp_city_state_zip = replace(replace(replace(cp_city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(cp_city_state_zip, 2)



'creating the word doc if selected in the dialog
'--------------------------------------WORD DOC-----------------------------------NEED TO FINISH******************************************************
DIM word_doc_open, objWord, objDoc, objSelection

'creating the workd application object (if any of the Word docs are selected and making it visible)
If _
	CP_MNOnlineEDAK_check = checked or _
	NCP_MNOnlineEDAK_check = checked THEN
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End If
	
'Opens cp mnonline word document
IF CP_MNOnlineEDAK_check = 1 THEN
 	set objDoc = objWord.Documents.Add("G:\Scripts CS\Word Docs\Active\mnonline edak.docx")	'Opens the specific Word doc 
	With objDoc
		.Formfields ("client_name").Result = CP_name
		.Formfields ("address_1").Result = cp_address
		.Formfields ("address_2").Result = cp_city_state_zip
		.Formfields ("case_number").Result = PRISM_case_number
		.Formfields ("client_name2").Result = CP_name
		'.Formfields ("pin_sent").Result = IF CP_PIN_Notice_check = 1 
		'.Formfields ("pin_not_sent").Result = IF CP_MNOnlineEDAK_check
		.Formfields ("mci").Result = CP_MCI
		.Formfields ("worker_name").Result = worker_name 
		.Formfields ("worker_phone").Result = worker_phone
		'.Formfields ("worker_fax").Result = 
		'.Formfields ("worker_email").Result = 
 	END WITH
END IF

'Opens cp mnonline work document
IF NCP_MNOnlineEDAK_check = 1 THEN
	set objDoc = objWord.Documents.Add("G:\Scripts CS\Word Docs\Active\mnonline edak.docx")	'Opens the specific Word doc 
	With objDoc
		.Formfields ("client_name").Result = NCP_name
		.Formfields ("address_1").Result = ncp_address
		.Formfields ("address_2").Result = ncp_city_state_zip
		.Formfields ("case_number").Result = PRISM_case_number
		.Formfields ("client_name2").Result = NCP_name
		'.Formfields ("pin_sent").Result = 
		'.Formfields ("pin_not_sent").Result = 
		.Formfields ("mci").Result = NCP_MCI
		.Formfields ("worker_name").Result = worker_name 
		.Formfields ("worker_phone").Result = worker_phone
		'.Formfields ("worker_fax").Result = 
		'.Formfields ("worker_email").Result = 
 	END WITH
END IF


'---------------------------------END WORD DOC------------------------------------

'-----CAAD NOTE--------------------------------------------------------------------

IF CAAD_note = checked THEN 
'bring to CAAD screen to create a CAAD note
	CALL navigate_to_PRISM_screen ("CAAD")																					
	PF5
	EMWriteScreen "A", 3, 29
	EMWriteScreen "free", 4, 54
	EMSetCursor 16, 4

'this will add information to the CAAD note
	IF CP_PIN_Notice_check = 1 AND CP_MNOnlineEDAK_check = 0 THEN CALL write_variable_in_CAAD ("*Sent cp pin number.")
	IF CP_PIN_Notice_check = 0 AND CP_MNOnlineEDAK_check = 1 THEN CALL write_variable_in_CAAD ("*Sent cp Mnonline Edak with Participant number.")
	IF CP_PIN_Notice_check = 1 AND CP_MNOnlineEDAK_check = 1 THEN CALL write_variable_in_CAAD ("*Sent cp pin number and Mnonline edak with participant number.")
	IF NCP_PIN_Notice_check = 1 AND NCP_MNOnlineEDAK_check = 0 THEN CALL write_variable_in_CAAD ("*Sent ncp pin number.")
	IF NCP_PIN_Notice_check = 0 AND NCP_MNOnlineEDAK_check = 1 THEN CALL write_variable_in_CAAD ("*Sent ncp Mnonline Edak with Participant number.")
	IF NCP_PIN_Notice_check = 1 AND NCP_MNOnlineEDAK_check = 1 THEN CALL write_variable_in_CAAD ("*Sent ncp pin number and Mnonline edak with participant number.")
	CALL write_variable_in_CAAD(Initials_CAAD)
	transmit
	PF3
END IF


  
'------END OF CAAD NOTE--------------------------------------------------------------


'SENDING DORD F0999 PIN--------------------------------------------------------------	
IF NCP_PIN_Notice_check = 1 THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "          ", 4, 15
	EMWriteScreen "  ", 4, 26 
	EMWriteScreen "F0999", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	EMWriteScreen NCP_MCI, 4, 15
	transmit
END IF

IF CP_PIN_Notice_check = 1 THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "          ", 4, 15
	EMWriteScreen "  ", 4, 26 

	EMWriteScreen "F0999", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	EMWriteScreen CP_MCI, 4, 15
	transmit
END IF

'reminder to print and mail documents
IF CP_MNOnlineEDAK_check or NCP_MNOnlineEDAK_check = checked THEN MsgBox ( "IMPORTANT!!  IMPORTANT!!" & vbNewline & vbNewline & "REMEMBER TO PRINT and MAIL WORD DOCUMENTS! " )

script_end_procedure("")

