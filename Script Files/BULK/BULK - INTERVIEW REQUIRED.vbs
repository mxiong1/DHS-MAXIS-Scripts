'OPTION EXPLICIT
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - INTERVIEW REQUIRED.vbs"
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
	
'Declaring variables
DIM appointment_required_dialog, REPT_panel, MAXIS_footer_month, MAXIS_footer_year, ButtonPressed

'DIALOG'----------------------------------------------------------------------------------------------------
BeginDialog appointment_required_dialog, 0, 0, 256, 80, "Appointment required dialog"
  DropListBox 70, 10, 60, 15, "REPT/ACTV"+chr(9)+"REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  EditBox 210, 10, 20, 15, MAXIS_footer_month
  EditBox 230, 10, 20, 15, MAXIS_footer_year
  EditBox 70, 30, 180, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 145, 50, 50, 15
    CancelButton 200, 50, 50, 15
  Text 5, 15, 55, 10, "Create list from:"
  Text 140, 15, 65, 10, "Footer month/year:"
  Text 5, 35, 60, 10, "Worker number(s):"
  Text 5, 50, 130, 25, "Enter 7 digits of each, (ex: x######). If entering multiple workers, separate each with a comma."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs current footer month/year
EMConnect ""
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'DISPLAYS DIALOG
DO                              
	err_msg = ""	
	Dialog appointment_required_dialog	
	If ButtonPressed = 0 then StopScript	
	If worker_number = "" or len(worker_number) > 7 or len(worker_number) < 7 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."	
    If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."	
    If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."	
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine	
LOOP until err_msg = ""	

'creating month plus 1 and plus 2
cm_plus_1 = dateadd("M", 1, date)
cm_plus_2 = dateadd("M", 2, date)
'creating a last day of recert variable
last_day_of_recert = DatePart("M", cm_plus_2) & "/01/" & DatePart("YYYY", cm_plus_2)
last_day_of_recert = dateadd("D", -1, last_day_of_recert)

'Grabbing the worker's X number.
CALL find_variable("User: ", worker_number, 7)


'Opening the Excel file, (now that the dialog is done)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 2).Value = "Interview Date & Time"
objExcel.cells(1, 2).Font.Bold = TRUE
objExcel.cells(1, 5).Value = "Privileged Cases"
objExcel.cells(1, 5).Font.Bold = TRUE

'If the appointments_per_time_slot variable isn't declared, it defaults to 1
IF appointments_per_time_slot = "" THEN appointments_per_time_slot = 1
IF alt_appointments_per_time_slot = "" THEN alt_appointments_per_time_slot = 1

'Checking for MAXIS
CALL check_for_MAXIS(false)

'We need to get back to SELF and manually update the footer month
back_to_SELF
current_month = DatePart("M", date)
IF len(current_month) = 1 THEN current_month = "0" & current_month
current_year = DatePart("YYYY", date)
current_year = right(current_year, 2)

'Determining the month that the script will access REPT/REVS. It's CM+2 for most, but CM+1 in developer mode
revs_month = DateAdd("M", 2, date)
IF developer_mode = True THEN revs_month = DateAdd("M", -1, revs_month)
revs_year = DatePart("YYYY", revs_month)
revs_year = right(revs_year, 2)
revs_month = DatePart("M", revs_month)
IF len(revs_month) = 1 THEN revs_month = "0" & revs_month

'writing current month and transmitting
EMWriteScreen current_month, 20, 43
EMWriteScreen current_year, 20, 46
transmit

'navigating to REVS and entering REVS Month and year as determined above
CALL navigate_to_MAXIS_screen("REPT", "REVS")
EMWriteScreen revs_month, 20, 55
EMWriteScreen revs_year, 20, 58
transmit

'<<<<<<<<<<<<TEMP
worker_number = worker_number_editbox

'Checking to see if the worker running the script is the the worker selected, if not it will enter the selected worker's number
EMReadScreen current_worker, 7, 21, 6
IF UCASE(current_worker) <> UCASE(worker_number) THEN
	EMWriteScreen UCASE(worker_number), 21, 6
	transmit
END IF

'Grabbing case numbers from REVS for requested worker
Excel_row = 2	'Declaring variable prior to do...loops
DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
	MAXIS_row = 7	'Setting or resetting this to look at the top of the list
	DO		'All of this loops until MAXIS_row = 19
		'Reading case information (case number, SNAP status, and cash status)
		EMReadScreen case_number, 8, MAXIS_row, 6
		EMReadScreen SNAP_status, 1, MAXIS_row, 45
		EMReadScreen cash_status, 1, MAXIS_row, 34

		'Navigates though until it runs out of case numbers to read
		IF case_number = "        " then exit do

		'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
		If cash_status = "-" 	then cash_status = ""
		If SNAP_status = "-" 	then SNAP_status = ""
		If HC_status = "-" 		then HC_status = ""

		'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
		If trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" then add_case_info_to_Excel = True
		If trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" then add_case_info_to_Excel = True

		'Adding the case to Excel
		If add_case_info_to_Excel = True then
			ObjExcel.Cells(excel_row, 1).Value = case_number
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
	EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
Loop until last_page_check = "THIS IS THE LAST PAGE"

'Now the script will go through STAT/REVW for each case and check that the case is at CSR or ER and remove the cases that are at CSR from the list.
excel_row = 2		'Resets the variable to 2, as it needs to look through all of the cases on the Excel sheet!
reviews_total = 0	'Sets this to 0 for the following do...loop. It'll exit once it's hit the reviews cap

DO 'Loops until there are no more cases in the Excel list

	'Grabs the case number
	case_number = objExcel.cells(excel_row, 1).Value

	'Goes to STAT/REVW
	CALL navigate_to_MAXIS_screen("STAT", "REVW")

	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 'Delete priv cases from excel sheet, save to a list for later

		priv_case_list = priv_case_list & "|" & case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
		msgbox priv_case_list

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
		IF CSR_mo = left(cm_plus_2, 2) and CSR_yr = right(cm_plus_2, 2) THEN recert_status = "NO"
		IF recert_mo = left(cm_plus_2, 2) and recert_yr = right(cm_plus_2, 2) THEN recert_status = "YES"

		'If it's not a recert, delete it from the excel list and move on with our lives
		IF recert_status = "NO" THEN
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			EMReadScreen MFIP_prog_check, 2, 6, 67		'checking for an active MFIP case
			EMReadScreen MFIP_status_check, 4, 6, 74
			If MFIP_prog_check <> "MF" AND MFIP_status_check <> "ACTV" THEN 	'if MFIP is active, then case will not be deleted.
				SET objRange = objExcel.Cells(excel_row, 1).EntireRow
				objRange.Delete				'all other cases that are not due for a recert will be deleted
				excel_row = excel_row - 1
			END If
		END IF
	END IF
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter	
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list of cases to check for recert is complete

NEXT

'Formatting the columns to autofit after they are all finished being created.
objExcel.Columns(1).autofit()
objExcel.Columns(2).autofit()
objExcel.Columns(3).autofit()
objExcel.Columns(4).autofit()

'Creating the list of privileged cases and adding to the spreadsheet
prived_case_array = split(priv_case_list, "|")
excel_row = 2

FOR EACH case_number in prived_case_array
	objExcel.cells(excel_row, 5).value = case_number
	excel_row = excel_row + 1
NEXT

script_end_procedure("Success! The Excel file now has all of the cases that require interviews for renewals.  Please manually review the list of privileged cases.")