'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EMERGENCY SCREENING.vbs"
start_time = timer

'Declared variables for the FuncLib
'DIM name_of_script, start_time, FuncLib_URL, run_locally, default_directory, use_master_branch, req, fso, row

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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'Declared variables for main script
'DIM emergency_screening_dialog, case_number, HH_members, eviction_check, utility_disconnect_check
'DIM homelessness_check, security_deposit_check, affordable_housing_yes, affordable_housing_no
'DIM EMER_HSR_manual_button, affordbable_housing, meets_residency, net_income, ButtonPressed, worker_signature
'DIM err_msg, footer_month, footer_year, begin_search_month, begin_search_year, EMER_type, EMER_amt_issued
'DIM EMER_elig_start_date, EMER_elig_end_date, monthly_standard, EMER_available_date, emer_issued, col
'DIM last_page_check, crisis, EMER_last_used_dates, screening_determination, Screening_options, script_repository

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog emergency_screening_dialog, 0, 0, 286, 235, "Emergency Screening dialog"
  EditBox 60, 15, 65, 15, case_number
  ComboBox 255, 15, 25, 15, "1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20", HH_members
  CheckBox 15, 55, 40, 10, "Eviction", eviction_check
  CheckBox 65, 55, 70, 10, "Utility disconnect", utility_disconnect_check
  CheckBox 140, 55, 60, 10, "Homelessness", homelessness_check
  CheckBox 210, 55, 65, 10, "Security deposit", security_deposit_check
  ComboBox 210, 75, 70, 15, "Select one..."+chr(9)+"Affordable"+chr(9)+"Not affordable", affordbable_housing
  ComboBox 230, 95, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", meets_residency
  EditBox 150, 120, 130, 15, net_income
  EditBox 75, 145, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 145, 50, 15
    CancelButton 230, 145, 50, 15
    PushButton 135, 170, 145, 15, "HSR Manual Emergency Assistance page ", EMER_HSR_manual_button
  GroupBox 5, 40, 275, 30, "Crisis (Check all that apply. If none, do not check any):"
  Text 140, 20, 105, 10, "Number of EMER HH members:"
  Text 10, 100, 220, 10, "Has anyone in the HH been residing in MN for more than 30 days?"
  Text 35, 210, 200, 10, "Information to be added to help HSR's answer the questions."
  Text 10, 80, 150, 10, "Is the household's living situation affordable?"
  Text 10, 20, 45, 10, "Case number:"
  Text 10, 125, 125, 10, "What is the household's NET income?"
  GroupBox 5, 195, 275, 35, "Info about net income/affordability:"
  Text 10, 150, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(case_number)

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
'creating current month as footer month/year'
footer_month = datepart("m", date)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", date)
footer_year = right(footer_year, 2)

'creating month variable 13 months prior to current footer month/year to search for EMER programs issued
begin_search_month = dateadd("m", -13, date)
begin_search_year = datepart("yyyy", begin_search_month)
begin_search_year = right(begin_search_year, 2)
begin_search_month = datepart("m", begin_search_month)
If len(begin_search_month) = 1 then begin_search_month = "0" & begin_search_month
'End of date calculations----------------------------------------------------------------------------------------------

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog emergency_screening_dialog
		cancel_confirmation
		'Opening the the HSR manual to the NOMI page
		IF buttonpressed = EMER_HSR_manual_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Emergency_Assistance_Policy.aspx")
		If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If HH_members = "" or IsNumeric(HH_members) = False then err_msg = err_msg & vbNewLine & "* Enter the number of household members."
		If affordbable_housing = "Select one..." then err_msg = err_msg & vbNewLine & "* Answer the affordable living situation question."
		If meets_residency = "Select one..." then err_msg = err_msg & vbNewLine & "* Answer the MN residency question."
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		If net_income = "" or IsNumeric(net_income) = False then err_msg = err_msg & vbNewLine & "* Enter the household's net income."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
LOOP until ButtonPressed = -1

'Checking for an active MAXIS session
Call check_for_MAXIS(True)
'navigating to INQX'
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43	'entering current footer month/year
EMWriteScreen footer_year, 20, 46
Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_search_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_search_year, 6, 41
EMWriteScreen footer_month, 6, 53		'entering current footer month/year
EMWriteScreen footer_year, 6, 56
EMWriteScreen "x", 9, 50		'selecting EA
EMWriteScreen "x", 11, 50		'selecting EGA
transmit

'searching for EA/EG issued on the INQD screen
DO
	row = 6
	DO
		EMReadScreen emer_issued, 1, row, 16		'searching for EMER programs as they start with E
		IF emer_issued = "E" then
			'reading the EMER information for EMER issuance
			EMReadScreen EMER_type, 2, row, 16
			EMReadScreen EMER_amt_issued, 7, row, 39
			EMReadScreen EMER_elig_start_date, 8, row, 62
			EMReadScreen EMER_elig_end_date, 8, row, 73
			exit do
		ELSE
			row = row + 1
		END IF
	Loop until row = 18				'repeats until the end of the page
		PF8
		EMReadScreen last_page_check, 21, 24, 2
		If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

'creating variables and conditions for EMER screening
EMER_available_date = dateadd("d", 1, EMER_elig_end_date)	'creating emer available date that is 1 day past the EMER_elig_end_date
EMER_last_used_dates = EMER_elig_start_date & " - " & EMER_elig_end_date	'combining dates into new variable

If emer_issued <> "E" then	'creating variables for cases that have not had EMER issued in current 13 months
 	EMER_last_used_dates = "n/a"
	EMER_available_date = "Currently available"
END IF

'Logic to enter what the "crisis" variable is from the check boxes indicated
If eviction_check = 1 then crisis = crisis & "eviction, "
If utility_disconnect_check = 1 then crisis = crisis & "utility disconnect, "
If homelessness_check = 1 then crisis = crisis & "homelessness, "
If security_deposit_check = 1 then crisis = crisis & "security deposit, "
If eviction_check = 0 and utility_disconnect_check = 0 and homelessness_check = 0 and security_deposit_check = 0 then
  crisis = "no crisis given"
Else
  crisis = trim(crisis)
  crisis = left(crisis, len(crisis) - 1)
End if

'determining  200% FPG (using last year's amounts) per HH member---handles up to 20 members
If HH_members = "1" then monthly_standard = "1915"
If HH_members = "2" then monthly_standard = "2585"
If HH_members = "3" then monthly_standard = "3255"
If HH_members = "4" then monthly_standard = "3925"
If HH_members = "5" then monthly_standard = "4595"
If HH_members = "6" then monthly_standard = "5265"
If HH_members = "7" then monthly_standard = "5935"
If HH_members = "8" then monthly_standard = "6605"
If HH_members = "9" then monthly_standard = "7275"
If HH_members = "10" then monthly_standard = "7945"
If HH_members = "11" then monthly_standard = "8615"
If HH_members = "12" then monthly_standard = "9285"
If HH_members = "13" then monthly_standard = "9955"
If HH_members = "14" then monthly_standard = "10625"
If HH_members = "15" then monthly_standard = "11295"
If HH_members = "16" then monthly_standard = "11965"
If HH_members = "17" then monthly_standard = "12635"
If HH_members = "18" then monthly_standard = "13305"
If HH_members = "19" then monthly_standard = "13975"
If HH_members = "20" then monthly_standard = "14645"

'determining if client is potentially elig for EMER or not'
If crisis <> "no crisis given" AND affordbable_housing = "Affordable" AND meets_residency = "Yes" AND net_income < monthly_standard AND _
net_income <> "0" AND EMER_last_used_dates = "n/a" then
	screening_determination = "potentially eligible for emergency programs."
Else
	screening_determination = "NOT be eligible for emergency programs because: "
END IF

'if client is not elig, reason(s) for not being elig will be listed in the msgbox
If crisis = "no crisis given" then screening_determination = screening_determination & vbNewLine & "* No crisis meeting program requirements."
If affordbable_housing = "Not affordable" then screening_determination = screening_determination & vbNewLine & "* The household's living situation is not affordable."
IF meets_residency = "No" then screening_determination = screening_determination & vbNewLine & "* No one in the household has met 30 day residency requirements."
If net_income > monthly_standard then screening_determination = screening_determination & vbNewLine & "* Net income exceeds program guidelines."
IF net_income = "0" then screening_determination = screening_determination & vbNewLine & "* HH does not have current/ongoing income."
If EMER_last_used_dates <> "n/a" then screening_determination = screening_determination & vbNewLine & "* Emergency funds were used within the last year from the eligibility period."

'Msgbox with screening results. Will give the user the option to cancel the script, case note the results, or use the EMER notes script
Screening_options = MsgBox ("Based on the information provided, this HH appears to " & screening_determination & vbNewLine & vbNewLine &"The last date emergency funds were used was: " & EMER_last_used_dates & "." & _
vbNewLine & "Emergency programs will be available to the HH again on: " & EMER_available_date & "." & vbNewLine & vbNewLine & "Would you like to start the NOTES - EMERGENCY script?" , vbYesNoCancel, "Screening results dialog")
IF Screening_options = vbCancel then script_end_procedure("")	'ends the script
IF Screening_options = vbYes then call run_from_GitHub(script_repository & "/NOTES/NOTES - EMERGENCY.vbs")	'run the NOTES EMER script
IF Screening_options = vbNO then
	case_note_option = Msgbox("Would you like to case note the emergency screening information?", vbYesNoCancel, "Screening case note")
		If case_note_option = vbCancel then script_end_procedure("")
		If case_note_option = VbNo then script_end_procedure("Information regarding the emergency screening has not been recorded in case note.")
		If case_note_option = vbYes then 'just the case note option
	'The case note
	Call start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("--//--Emergency Programs Screening--//--")
	Call write_bullet_and_variable_in_CASE_NOTE("Number of HH members", HH_members)
	Call  write_bullet_and_variable_in_CASE_NOTE("Crisis/Type of Emergency", crisis)
	Call write_bullet_and_variable_in_CASE_NOTE("Living situation is", affordbable_housing)
	Call write_bullet_and_variable_in_CASE_NOTE("Does any member of the HH meet 30 day residency requirements", meets_residency)
	Call write_bullet_and_variable_in_CASE_NOTE("Net income for HH", net_income)
	Call write_variable_in_CASE_NOTE("---")
	Call write_bullet_and_variable_in_CASE_NOTE("Last date EMER programs were used", EMER_last_used_dates)
	Call write_variable_in_CASE_NOTE("* Date EMER programs will be available to HH: " & EMER_available_date)
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	END IF
END IF

script_end_procedure("")
