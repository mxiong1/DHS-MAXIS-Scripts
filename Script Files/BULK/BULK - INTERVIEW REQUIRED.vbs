worker_county_code = "x127"
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - INTERVIEW REQUIRED.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer

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
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 39			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

BeginDialog appointment_required_dialog, 0, 0, 286, 115, "Appointment required dialog"
  DropListBox 70, 10, 60, 15, "REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  DropListBox 185, 10, 90, 15, "Select one..."+chr(9)+"Current month"+chr(9)+"Current month plus one"+chr(9)+"Current month plus two", footer_selection
  EditBox 70, 30, 205, 15, worker_number
  CheckBox 5, 75, 155, 10, "Select all active workers in the agency", all_workers_check
  CheckBox 5, 90, 155, 10, "Add client phone number(s) to list", add_phone_numbers_check
  ButtonGroup ButtonPressed
    OkButton 170, 90, 50, 15
    CancelButton 225, 90, 50, 15
  Text 5, 50, 265, 15, "Enter 7 digits of each, (ex: x######). If entering multiple workers, separate each with a comma."
  Text 135, 15, 45, 10, "Time period:"
  Text 5, 35, 60, 10, "Worker number(s):"
  Text 5, 15, 55, 10, "Create list from:"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog appointment_required_dialog
		If ButtonPressed = 0 then StopScript
		If worker_number = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Enter a valid worker number."
		If footer_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the time period for your list."
		If REPT_panel = "REPT/REVW" and footer_selection = "Current month plus two" then err_msg = err_msg & VbNewLine & "* This is time period is not an option REPT/REVW. Please select a new time period."
		If (REPT_panel = "REPT/REVS" and footer_selection = "Current month plus two" and datePart("d", date) < 16) then err_msg = err_msg & VbNewLine & "* This is not a valid time period for REPT/REVS until the 16th of the month. Please select a new time period."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false	

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Splitting array for use by the for...next statement
	worker_array = split(worker_array, ", ")
End if
msgbox "array complete"

'creating dates for the footer_selection variable
If footer_selection = "Current month" then
	footer_month = DatePart("M", date)
	IF len(footer_month) = 1 THEN footer_month = "0" & footer_month
	footer_year = DatePart("YYYY", date)
	footer_year = right(footer_year, 2)
ELSEif footer_selection = "Current month plus one" then
	footer_month = dateadd("M", 1, date)
	footer_month = datePart("M", footer_month)
	IF len(footer_month) = 1 THEN footer_month = "0" & footer_month
	footer_year = DatePart("YYYY", date)
	footer_year = right(footer_year, 2)
ELSEIF footer_selection = "Current month plus two" then
	footer_month = dateadd("M", 2, date)
	footer_month = datePart("M", footer_month)
	IF len(footer_month) = 1 THEN footer_month = "0" & footer_month
	footer_year = DatePart("YYYY", date)
	footer_year = right(footer_year, 2)
END IF

'We need to get back to SELF and manually update the footer month
back_to_self
REPT_panel = right(REPT_panel, 4)	're-establishing variable to exclude all but the last 4 characters to the right
EMWriteScreen "________", 18, 43

'writing in REPT panel and footer selection
If footer_selection = "Current month plus two" then
	Call navigate_to_MAXIS_screen("REPT", REPT_panel)
	EMWriteScreen CM_mo, 20, 43
	EMWriteScreen CM_yr, 20, 46
	transmit
	msgbox "current month transmit"
	EMWriteScreen footer_month, 20, 55
	EMWriteScreen footer_year, 20, 58
ELSE
	Call navigate_to_MAXIS_screen("REPT", REPT_panel)
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
END IF
transmit
msgbox footer_month

'Opening the Excel file, (now that the dialog is done)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'formatting excel file with columns for case number and phone numbers
objExcel.cells(1, 1).Value = "CASE NUMBER"
objExcel.cells(1, 2).Value = "Worker Number"
IF add_phone_numbers_check = 1 then 
	objExcel.Cells(1, 3).Value = "Phone Number 1"
	objExcel.Cells(1, 4).Value = "Phone Number 2"
	objExcel.Cells(1, 5).Value = "Phone Number 3"
END IF 
objExcel.cells(1, 6).Value = "Privileged Cases"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Grabbing case numbers from REVS for requested worker
Excel_row = 2	'Declaring variable prior to do...loops

'start of the FOR...next loop
For each worker in worker_array
	If trim(worker) = "" then exit for
	worker_ID = trim(worker)

	If REPT_panel = "REPT/ACTV" then 'THE REPT PANEL HAS THE worker NUMBER IN DIFFERENT COLUMNS. THIS WILL DETERMINE THE CORRECT COLUMN FOR THE worker NUMBER TO GO
		worker_ID_col = 13
	Else
		worker_ID_col = 6
	End if
	'writing in the worker number in the correct col
	EMWriteScreen worker_ID, 21, worker_ID_col
	transmit
	msgbox worker_ID
	'THIS DO...LOOP DUMPS THE CASE NUMBER AND NAME OF EACH CLIENT INTO A SPREADSHEET
	Do
		MAXIS_row = 7	'Setting or resetting this to look at the top of the list
		DO	'All of this loops until MAXIS_row = 19
			'Reading case information (case number, SNAP status, and cash status)
			msgbox MAXIS_row
			EMReadScreen case_number, 8, MAXIS_row, 6
			EMReadScreen SNAP_status, 1, MAXIS_row, 45
			If REPT_panel = "REVS" then
				EMReadScreen cash_status, 1, MAXIS_row, 34
			ELSE
				EMReadScreen cash_status, 1, MAXIS_row, 35		'REPT/ACTV cash status is on col 35, REVS is on col 34 (Thanks MAXIS)
			END IF
			msgbox cash_status
			'Navigates though until it runs out of case numbers to read
			IF case_number = "        " then exit do

			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
			If cash_status = "-" 	then cash_status = ""
			If SNAP_status = "-" 	then SNAP_status = ""
			
			'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
			If ( ( trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" ) or ( trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" ) ) and reviews_total <= max_reviews_per_worker then
				add_case_info_to_Excel = True
				reviews_total = reviews_total + 1
			End if
			
			'Adding the case to Excel
			If add_case_info_to_Excel = True then
				ObjExcel.Cells(excel_row, 2).value = worker_number
				ObjExcel.Cells(excel_row, 1).value = case_number
				excel_row = excel_row + 1
			End if
			
			'On the next loop it must look to the next row
			MAXIS_row = MAXIS_row + 1
			'Clearing variables before next loop
			add_case_info_to_Excel = ""
			case_number = ""
			
		Loop until MAXIS_row = 19		'Last row in REPT/REVS
		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
		PF8
		EMReadScreen last_page_check, 21, 24, 02
	Loop until last_page_check = "THIS IS THE LAST PAGE"
NEXT

'Now the script will go through STAT/REVW for each case and check that the case is at CSR or ER and remove the cases that are at CSR from the list.
excel_row = 2		'Resets the variable to 2, as it needs to look through all of the cases on the Excel sheet!

DO 'Loops until there are no more cases in the Excel list
	'Grabs the case number
	case_number = objExcel.cells(excel_row, 1).Value
	If case_number = "" then exit do
	'Goes to STAT/REVW
	CALL navigate_to_MAXIS_screen("STAT", "REVW")

	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 'Delete priv cases from excel sheet, save to a list for later
		priv_case_list = priv_case_list & "|" & case_number
		SET objRange = objExcel.Cells(excel_row, 6).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
	ELSE		'For all of the cases that aren't privileged...
		'Looks at review details
		EMwritescreen "x", 5, 58
		Transmit
		DO
			EMReadScreen SNAP_popup_check, 7, 5, 43
		LOOP until SNAP_popup_check = "Reports"

		'The script will now read the CSR MO/YR and the Recert MO/YR
		EMReadScreen CSR_mo, 2, 9, 26
		EMReadScreen CSR_yr, 2, 9, 32
		EMReadScreen recert_mo, 2, 9, 64
		EMReadScreen recert_yr, 2, 9, 70
		'It then compares what it read to the previously established current month plus 2 and determines if it is a recert or not. If it is a recert we need an interview
		If recert_mo = "__" and recert_yr = "__" then recert_status = "NO" 	'in case there is no SNAP, will check for MFIP being active, if not active, then will remove from list in next IF...THEN statement.
		IF CSR_mo = left(footer_month, 2) and CSR_yr = right(footer_year, 2) THEN recert_status = "NO"
		IF recert_mo = left(footer_month, 2) and recert_yr = right(footer_year, 2) THEN
			recert_status = "YES"
		ELSE
			recert_status = "NO"
		END IF

		'checking for ACTV MFIP'
		IF recert_status = "NO" THEN
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			EMReadScreen MFIP_prog_check, 2, 6, 67		'checking for an active MFIP case
			EMReadScreen MFIP_status_check, 4, 6, 74
			If MFIP_prog_check = "MF" AND MFIP_status_check = "ACTV" THEN 'if MFIP is active, then case will not be deleted.
				recert_status = "YES"
			ELSE
				recert_status = "NO"
			END If
		END IF

		'If it's not a recert, delete it from the excel list and move on with our lives
		If recert_status = "NO" then
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete				'all other cases that are not due for a recert will be deleted
			excel_row = excel_row - 1
		ELSEIF recert_status = "YES" then	'if yes, then grabs the phone numbers from the ADDR panel'
			EMReadScreen worker_number, 7, 21, 21
			objExcel.cells(excel_row, 2).Value = worker_number
		END IF
		IF add_phone_numbers_check = 1 then 
			Call navigate_to_MAXIS_screen("STAT", "ADDR")
			EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
			If phone_number_one <> "( ___ ) ___ ____" then objExcel.cells(excel_row, 3).Value = phone_number_one
			EMReadScreen phone_number_two, 16, 18, 43
			If phone_number_two <> "( ___ ) ___ ____" then objExcel.cells(excel_row, 4).Value = phone_number_two
			EMReadScreen phone_number_three, 16, 19, 43
			If phone_number_three <> "( ___ ) ___ ____" then objExcel.cells(excel_row, 5).Value = phone_number_three	
		END IF
	END IF
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list of cases to check for recert is complete

'Formatting the columns to autofit after they are all finished being created.
FOR i = 1 to 6		'formatting the cells'
 	objExcel.Cells(1, i).Font.Bold = True		'bold font'
 	objExcel.Columns(i).AutoFit()						'sizing the colums'
 NEXT

'Creating the list of privileged cases and adding to the spreadsheet
prived_case_array = split(priv_case_list, "|")
excel_row = 2

FOR EACH case_number in prived_case_array
	objExcel.cells(excel_row, 6).value = case_number
	excel_row = excel_row + 1
NEXT

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the begining (because counting :p)

script_end_procedure("Success! The Excel file now has all of the cases that require interviews for renewals.  Please manually review the list of privileged cases (if any).")
