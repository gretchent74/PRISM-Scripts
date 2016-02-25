'GATHERING STATS----------------------------------------------------------------------------------------------------

'name_of_script = "ACTIONS - Interstate Transfer.vbs"
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


DIM Interstate_Transfer_Note, scanned_correctly, INCM_updated, Acknowledgement, State_Order, Responding_State ,initials_caad, Interstate_Worker_droplistbox, ButtonPressed
DIM err_msg

BeginDialog Interstate_Transfer_Note, 0, 0, 236, 205, "Interstate Transfer Note"
  Text 5, 5, 85, 10, "Interstate Case Transfer"
  CheckBox 10, 20, 225, 10, "Interstate Documents Scanned Correctly to Onbase as UIFSA OUT", scanned_correctly
  CheckBox 10, 40, 150, 10, "INCM updated with local office information", INCM_updated
  Text 5, 65, 110, 10, "Date Acknowledgement Received"
  EditBox 120, 60, 50, 15, Acknowledgement
  Text 5, 90, 80, 10, "State that issued Order"
  EditBox 90, 85, 50, 15, State_Order
  Text 5, 115, 60, 10, "Responding State"
  EditBox 70, 110, 50, 15, Responding_State
  Text 5, 140, 130, 10, "Case Transferred to Interstate Worker"
  DropListBox 135, 140, 95, 45,  "Select one..."+chr(9)+"Heidi F"+chr(9)+"Mary SC"+chr(9)+"Pam G", Interstate_Worker_droplistbox
  Text 5, 165, 70, 10, "Initials for CAAD note"
  EditBox 80, 160, 40, 15, initials_caad
  ButtonGroup ButtonPressed
    OkButton 115, 185, 50, 15
    CancelButton 180, 185, 50, 15
EndDialog


'connecting to bluezone
EMConnect ""

'to pull up my prism 
EMFocus

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

'adding a loop
Do 
               Dialog Interstate_Transfer_Note
                If ButtonPressed = 0 THEN StopScript 'this if for cancel button to work
                If Acknowledgement = "" THEN err_msg = err_msg & vbcr & "Date acknowledgement received must be Completed"
                If State_Order = "" THEN err_msg = err_msg & vbcr & "State that issued Order must be Completed"
		    If Responding_State = "" THEN err_msg = err_msg & vbcr & "Responding state must be Completed"
		    If initials_caad = "" THEN err_msg = err_msg & vbcr & "Initials for CAAD note"
		    If Interstate_Worker_droplistbox = "Select one..." THEN err_msg = err_msg & vbcr & "Interstate worker must be Completed"
  		    If err_Msg <> "" THEN 
	               	err_Msg = "NOTICE" & vbcr & err_msg & vbCr & vbCr & "Please resolve for this script to continue"
				MsgBox err_msg
		    END IF		

		    
LOOP UNTIL Acknowledgement <> "" and _
	State_Order <> "" and _
	Responding_State <> "" and _
	initials_caad <> "" and _
	Interstate_Worker_droplistbox <> "Select one..."



'brings me to caad
CALL navigate_to_PRISM_screen ("CAAD")

'to create a caad note
PF5
EMWriteScreen "A", 3, 29

'sets caad note type free
EMWriteScreen "free", 4, 54

'set the cursor on the caad note area
EMSetCursor 16, 4

'added so this info from the dialog will go in caad note

CALL write_variable_in_CAAD("Interstate Case Transfer")
CALL write_bullet_and_variable_in_CAAD("Date Acknowledgement Received", Acknowledgement)
CALL write_bullet_and_variable_in_CAAD("State that issued Order", State_Order & "    Responding State: " & Responding_State)
IF scanned_correctly = checked THEN CALL write_variable_in_CAAD ("* Interstate Documents Scanned Correctly to Onbase as UIFSA OUT.")
IF INCM_updated = checked THEN CALL write_variable_in_CAAD ("* INCM updated with local office information.")
CALL write_bullet_and_variable_in_CAAD ("Case Transferred to Interstate Worker", Interstate_Worker_droplistbox)
CALL write_variable_in_CAAD(initials_caad)

'saves the CAAD note
transmit

'exits back out of the CAAD note
PF3


	If Interstate_Worker_droplistbox = "Heidi F" THEN 
'goes to caas and transfer case to Heidi 037 001 SMR 24
	CALL navigate_to_PRISM_screen("CAAS")
	EMWriteScreen "M", 3, 29
	EMWriteScreen "037", 9, 20
	EMWriteScreen "001", 10, 20
	EMWriteScreen "SMR", 11, 20
	EMWriteScreen "24",12, 20
END IF

	If Interstate_Worker_droplistbox = "Mary SC" THEN 
'goes to caas and transfer case to Mary SC 037 001 SMR 12
	CALL navigate_to_PRISM_screen("CAAS")
	EMWriteScreen "M", 3, 29
	EMWriteScreen "037", 9, 20
	EMWriteScreen "001", 10, 20
	EMWriteScreen "SMR", 11, 20
	EMWriteScreen "12",12, 20
END IF

	If Interstate_Worker_droplistbox = "Pam G" THEN 
'goes to caas and transfer case to Pam G 037 001 SMR 22
	CALL navigate_to_PRISM_screen("CAAS")
	EMWriteScreen "M", 3, 29
	EMWriteScreen "037", 9, 20
	EMWriteScreen "001", 10, 20
	EMWriteScreen "SMR", 11, 20
	EMWriteScreen "22",12, 20
END IF

'saves and exits caas emsendkey and emwaitready 0,0
transmit

CALL navigate_to_PRISM_screen("CAPS")

MsgBox ("Make sure PRISM APP Enabler is OPEN and then CONTROL P! Remember to add any special information to the CAAD Note!")

script_end_procedure("")
