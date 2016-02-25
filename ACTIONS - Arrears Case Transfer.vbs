'Option Explicit 'this has to be on the top, always
'Option Explicit

'GATHERING STATS----------------------------------------------------------------------------------------------------

'name_of_script = "ACTIONS - Arrears CAse Transfer.vbs"
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


DIM Arrears_Transfer_Note_Dialog, Completed_Arrears_Transfer_Procedure, Arrears_Case_Reason_droplistbox, Date_Ended, NonAccrual_Amount, Court_Ordered_Payback, Interest_droplistbox, Active_Contempt, Exp_Date, Active_Warrant, DL_droplistbox, Currently_Incarcerated, Ant_Release_Date, CPOD_Exists, Interstate, Interstate_droplistbox, Other_State, Arrears_Judgments, initials_caad, ButtonPressed
DIM err_msg

BeginDialog Arrears_Transfer_Note_Dialog, 0, 0, 256, 290, "Arrears Transfer Note"
  Text 5, 5, 75, 10, "Arrears Case Transfer"
  CheckBox 5, 20, 145, 10, "Completed Arrears Transfer Procedure", Completed_Arrears_Transfer_Procedure
  Text 5, 40, 75, 10, "Arrears Case Reason:"
  DropListBox 80, 35, 105, 45, "Select one..."+chr(9)+"Parental Rights Terminated"+chr(9)+"Emancipated"+chr(9)+"Reconciled Family"+chr(9)+"Child Adopted"+chr(9)+"NPA-Case closed by CP"+chr(9)+"Child Death"+chr(9)+"Other", Arrears_Case_Reason_droplistbox
  Text 5, 60, 65, 10, "Charging Stopped:"
  EditBox 70, 55, 50, 15, Date_Ended
  Text 5, 85, 105, 10, "Monthly Non-Accrual Amount:"
  EditBox 105, 80, 50, 15, NonAccrual_Amount
  CheckBox 165, 85, 120, 10, "Court Ordered Payback", Court_Ordered_Payback
  Text 5, 105, 65, 10, "Interest Stopped:"
  DropListBox 65, 100, 70, 45, "Select one..."+chr(9)+"No"+chr(9)+"Administravely"+chr(9)+"Court Ordered", Interest_droplistbox
  CheckBox 5, 125, 70, 10, "Active Contempt", Active_Contempt
  Text 75, 125, 55, 10, "Expiration Date:"
  EditBox 135, 120, 50, 15, Exp_Date
  CheckBox 5, 140, 60, 10, "Active Warrant", Active_Warrant
  Text 5, 160, 40, 10, "DL Status"
  DropListBox 45, 155, 105, 20, "Select one..."+chr(9)+"No Payment Agreement"+chr(9)+"Payment Agreement"+chr(9)+"Suspended"+chr(9)+"Suspended by Court Order"+chr(9)+"Suspended for Failure to Comply", DL_droplistbox
  CheckBox 5, 175, 90, 10, "Currently Incarcerated", Currently_Incarcerated
  Text 95, 175, 90, 10, "Anticipated Release Date:"
  EditBox 185, 170, 50, 15, Ant_Release_Date
  CheckBox 5, 190, 70, 10, "CPOD Exists", CPOD_Exists
  CheckBox 5, 210, 45, 10, "Interstate:", Interstate
  DropListBox 55, 205, 60, 45, "Select one..."+chr(9)+"Initiating"+chr(9)+"Responding", Interstate_droplistbox
  Text 125, 210, 25, 10, "State:"
  EditBox 150, 205, 65, 15, Other_State
  CheckBox 5, 235, 150, 10, "Judgment and Arrears Clean Up Completed", Arrears_Judgments
  Text 5, 260, 70, 10, "Initials for CAAD note"
  EditBox 80, 255, 40, 15, initials_caad
  ButtonGroup ButtonPressed
    OkButton 135, 270, 50, 15
    CancelButton 195, 270, 50, 15
EndDialog


'connecting to bluezone
EMConnect ""

'to pull up my prism 
EMFocus

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

'adding a loop
Do 
            err_msg = ""   
		Dialog Arrears_Transfer_Note_Dialog
                If ButtonPressed = 0 THEN StopScript 'this if for cancel button to work
                If Date_Ended = "" THEN err_msg = err_msg & vbNewline & "Date Current Charging Stopped must be Completed"
                If NonAccrual_Amount = "" THEN err_msg = err_msg & vbNewline & "Monthly Non-Accrual Amount must be Completed"
		    If initials_caad = "" THEN err_msg = err_msg & vbNewline & "Initials for CAAD note"
		    If Arrears_Case_Reason_droplistbox = "Select one..." THEN err_msg = err_msg & vbNewline & "Arrears case reason must be Completed"
                If Interest_droplistbox = "Select one..." THEN err_msg = err_msg & vbNewline & "Interest information must be Completed"
                If DL_droplistbox = "Select one..." THEN err_msg = err_msg & vbNewline & "DL Status must be Completed"
  		    If err_Msg <> "" THEN 
	               	Msgbox "***NOTICE***" & vbcr & err_msg & vbNewline & vbNewline & "Please resolve for this script to continue"
			END IF		

LOOP UNTIL err_msg = ""



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

CALL write_variable_in_CAAD("Arrears Case Transfer")
If Completed_Arrears_Transfer_Procedure = checked THEN call write_variable_in_CAAD ("* Completed Arrears Transfer Procedure.")
'arrears case reason dropboxlist
Call write_bullet_and_variable_in_CAAD ("Arrears Case Reason", Arrears_Case_Reason_droplistbox)
Call write_bullet_and_variable_in_CAAD("Charging Stopped", Date_Ended & "  Monthly NonAccrual Amount: " & NonAccrual_Amount)
If Court_Ordered_Payback = checked THEN call write_variable_in_CAAD ("* Court Ordered Payback.")
If Active_Contempt = checked THEN call write_variable_in_CAAD ("* Active_Contempt.  " & " Expiration Date: " & Exp_Date)
If Active_Warrant = checked THEN call write_variable_in_CAAD ("* Active_Warrant.")
'interest droplistbox
Call write_bullet_and_variable_in_CAAD("Interest Stopped", Interest_droplistbox)
'dl status droplistbox
Call write_bullet_and_variable_in_CAAD("DL Status", DL_droplistbox)
If Currently_Incarcerated = checked THEN call write_variable_in_CAAD ("* Currently Incarcerated." & "  Anticipated Release Date: " &   Ant_Release_Date)

If Interstate = checked THEN call write_variable_in_CAAD  ("* Interstate: "  & Interstate_droplistbox     & " " &  Other_State)
If CPOD_Exists = checked THEN call write_variable_in_CAAD ("* CPOD Exists.")
If Arrears_Judgments = checked THEN call write_variable_in_CAAD ("* Judgment and Arrears Clean Up Completed.")
Call write_variable_in_CAAD(initials_caad)


'saves the CAAD note
transmit

'exits back out of the CAAD note
PF3

'goes to caas and transfer case to arrears team and
CALL navigate_to_PRISM_screen("CAAS")

'puts the m in action field
EMWriteScreen "M", 3, 29

'adds the county dakota
EMWriteScreen "037", 9, 20

'adds the office
EMWriteScreen "001", 10, 20

'adds the arrears team LJH
EMWriteScreen "LJH", 11, 20

'add the position 14 for arrears team
EMWriteScreen "14",12, 20

'saves and exits caas emsendkey and emwaitready 0,0
transmit

'goes to the dord screen to send letter
CALL navigate_to_PRISM_screen("DORD")

'puts the A in action field
EMWriteScreen "A", 3, 29

'adds form document
EMWriteScreen "F0104", 6, 36

EMWriteScreen "ncp", 11, 51

transmit

'shift f2, to get to user lables 
PF14
 
EMWriteScreen "u", 20,14

transmit

EMSetCursor 7, 5

EMWriteScreen "S", 7, 5
EMWriteScreen "S", 8, 5

EMSendKey "<enter>"

EMWriteScreen "Your case has been transferred to the Arrears Team.  Please", 16, 15

PF3

EMWriteScreen "direct all future correspondence to the Arrears Team.", 16, 15

EMSendKey "<enter>"

PF3

EMWriteScreen "M", 3, 29

transmit


'goes to the dord screen to send letter
CALL navigate_to_PRISM_screen("DORD")

EMWriteScreen "C", 3, 29
EMSendKey "<enter>"

'puts the A in action field
EMWriteScreen "A", 3, 29

'adds form document
EMWriteScreen "F0104", 6, 36

EMWriteScreen "cpp", 11, 51

transmit

'shift f2, to get to user lables 
PF14
 
EMWriteScreen "u", 20,14

transmit

EMSetCursor 7, 5

EMWriteScreen "S", 7, 5
EMWriteScreen "S", 8, 5

transmit

EMWriteScreen "Your case has been transferred to the Arrears Team.  Please", 16, 15

PF3

EMWriteScreen "direct all future correspondence to the Arrears Team.", 16, 15

transmit 

PF3

EMWriteScreen "M", 3, 29

transmit

CALL navigate_to_PRISM_screen("CAPS")

MsgBox ("Make sure PRISM APP Enabler is OPEN and then CONTROL P!" & vbNewline & vbNewline & "Remember to send an email to CS ARREARS TEAM!")

script_end_procedure("")
