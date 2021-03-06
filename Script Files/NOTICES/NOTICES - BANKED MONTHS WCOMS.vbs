'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - BANKED MONTHS WCOMS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
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
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Dialogs
BeginDialog case_number_dlg, 0, 0, 211, 80, "Case Number Dialog"
  EditBox 70, 10, 60, 15, case_number
  EditBox 70, 30, 30, 15, approval_month
  EditBox 160, 30, 30, 15, approval_year
  ButtonGroup ButtonPressed
    OkButton 45, 55, 50, 15
    CancelButton 105, 55, 50, 15
  Text 10, 15, 55, 10, "Case Number: "
  Text 10, 35, 55, 10, "Approval Month:"
  Text 105, 35, 50, 10, "Approval Year:"
EndDialog


BeginDialog banked_months_menu_dialog, 0, 0, 356, 140, "Banked Months WCOMs"
  ButtonGroup ButtonPressed
    PushButton 10, 25, 90, 10, "All Banked Months Used", banked_months_used_button
    PushButton 10, 50, 90, 10, "Banked Months Notifier", banked_months_notifier
    PushButton 10, 75, 90, 10, "Closing for E/T Non-Coop", e_t_non_coop_button
    CancelButton 300, 120, 50, 15
  Text 110, 25, 230, 20, "-- Use this script when a client's SNAP is closing because they used all their eligible banked months."
  Text 110, 50, 230, 20, "-- Use this script to add a WCOM to a notice notifying the client they may be eligible for banked months."
  Text 110, 75, 235, 25, "-- Use this script to add a WCOM to a client's closing notice to inform them they are closing on banked months for Employment Services Non-Coop."
  GroupBox 5, 10, 345, 90, "WCOM"
EndDialog





'--- The script -----------------------------------------------------------------------------------------------------------------

EMConnect ""


call MAXIS_case_number_finder(case_number)
approval_month = DatePart("M", (DateAdd("M", 1, date)))
IF len(approval_month) = 1 THEN 
	approval_month = "0" & approval_month
ELSE
	approval_month = Cstr(approval_month)
END IF
approval_year = Right(DatePart("YYYY", (DateAdd("M", 1, date))), 2)

DO
	err_msg = ""
	dialog case_number_dlg
	cancel_confirmation
	IF case_number = "" THEN err_msg = "* Please enter a case number" & vbNewLine
	IF len(approval_month) <> 2 THEN err_msg = err_msg & "* Please enter your month in MM format." & vbNewLine
	IF len(approval_year) <> 2 THEN err_msg = err_msg & "* Please enter your year in YY format." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""	

CALL check_for_MAXIS(false)

DIALOG banked_months_menu_dialog
	cancel_confirmation
	
	'This is the WCOM for when the client has used all their banked months.
	IF ButtonPressed = banked_months_used_button THEN
		call navigate_to_MAXIS_screen("spec", "wcom")
		
		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit
		
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"
		
		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your financial worker if you have questions.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "
		
		wcom_type = "all banked months"
		
	'This is the WCOM for when the client is closing for ABAWD and is being notified that they could be eligible for banked months.
	ELSEIF ButtonPressed = banked_months_notifier THEN 
		call navigate_to_MAXIS_screen("spec", "wcom")
		
		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit
		
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"
		
		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "
		
		wcom_type = "banked months notifier"

	'This is the WCOM for when the client is closing on banked months for E&T Non-Coop
	ELSEIF ButtonPressed = e_t_non_coop_button THEN 
	
		DO
			hh_member = InputBox("Please enter the name of the client that is closing for E&T Non-Coop...")
			confirmation_msg = MsgBox("Please confirm to add the client's name to the WCOM: " & vbCr & vbCr & hh_member & " is closing on banked months for SNAP E&T Non-Cooperation." & vbCr & vbCr & "Is this correct? Press YES to continue. Press NO to re-enter the client's name. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirmation_msg = vbCancel THEN stopscript
		LOOP UNTIL confirmation_msg = vbYes
		
		call navigate_to_MAXIS_screen("spec", "wcom")
		
		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit
		
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"
		
		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP case is closing because " & hh_member & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your financial worker before your SNAP closes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above AND all other eligibility factors you may be eligible for SNAP. If you have questions please contact your financial worker.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "
		
		wcom_type = "non coop"
	END IF

'Outcome ---------------------------------------------------------------------------------------------------------------------

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	'Case note
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("---WCOM added regarding banked months---")
	IF wcom_type = "all banked months" THEN 
		CALL write_variable_in_CASE_NOTE("* WCOM added because client all eligible banked months have been used.")
	ELSEIF wcom_type = "non coop" THEN
		CALL write_variable_in_CASE_NOTE("* Banked months ending for SNAP E & T non-coop.")
	ELSEIF wcom_type = "banked months notifier" THEN 
		CALL write_variable_in_CASE_NOTE("* Client has used ABAWD counted months and MAY be eligible for banked months. Eligibility questions should be directed to financial worker.")
	END IF
	
	call write_variable_in_CASE_NOTE("---")
	IF worker_signature <> "" THEN 
		call write_variable_in_CASE_NOTE(worker_signature)
	ELSE
		worker_signature = InputBox("Please sign your case note...")
		CALL write_variable_in_CASE_NOTE(worker_signature)
	END IF
END IF

script_end_procedure("")
