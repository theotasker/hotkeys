; ================================================================================================================
; Selenium Paths, need updating as RXWizard gets updated
; ================================================================================================================

; ---------------------------------------------------------------------------------------------------------------------
; RXWizard Paths


global reviewPageCSS := {"":""
, "generateSchedule": "#main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button:nth-child(1)"
, "startStop": "#main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button"
, "scanScript": "#react > section > header > form > div > input"
, "panNumber": "#main-content-wrapper > div.page-content > div > div.name > div > span.pan > div > span"
, "scriptNumber": "#main-content-wrapper > div.page-content > div > div.name > div > span.case-name"
, "clinicName": "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(1) > div.ant-col.ant-col-17 > div > div"
, "patientName": "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(8) > div.ant-col.ant-col-17 > div > div"
, "newNote": ""}

global editPageCSS := {"":""
, "scanScript": "#react > section > header > form > div > input"
, "panNumber": "#pan > div > div > div.ant-select-selection-selected-value"
, "scriptNumber": "#main-content-wrapper > div > div > div.name > div > span.case-name"
, "clinicName": "#office > div > div > div.ant-select-selection-selected-value"
, "patientFirstName": "#patient_first_name"
, "patientLastName": "#patient_last_name"
, "newNote": ""}

global Path_GenerateScheduleCSS := "#main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button:nth-child(1)"

global Path_StartStopXPATH := "//*[@id=""""main-content-wrapper""""]/div[1]/main/div[2]/div/div[1]/button/span"

global Path_StartStopCSS := "#main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button"

; If user has manager privileges on the website, the button location changes
global Path_StartStopManagerXPATH := "//*[@id=""""main-content-wrapper""""]/div[1]/main/div[2]/div/div/div[2]/button[1]"

global Path_SaveNoteXPATH := "/html/body/div[4]/div/div[2]/div/div[2]/div[2]/div/button[2]"

global Path_UploadFileXPATH := "//*[@id=""""main-content-wrapper""""]/div[1]/main/div[1]/div[1]/form/div/div/div/div[2]/span/div[1]/span/div/p[2]"

global Path_ScanScriptCSS := "#react > section > header > form > div > input"

global Path_ReviewButtonCSS := "#main-content-wrapper > div > main > form > div.ant-row-flex.ant-row-flex-end > div:nth-child(2) > button"

global Path_ReviewPanNumberCSS := "#main-content-wrapper > div.page-content > div > div.name > div > span.pan > div > span"

global Path_ReviewPanNumberXPATH := "//*[@id=""main-content-wrapper""]/div[1]/div/div[1]/div/span[4]/div/span/text()"

global Path_ScriptNumberCSS := "#main-content-wrapper > div.page-content > div > div.name > div > span.case-name"

global Path_ClinicNameCSS := "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(1) > div.ant-col.ant-col-17 > div > div"

global Path_PatientNameCSS := "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(8) > div.ant-col.ant-col-17 > div > div"

global Path_NewNoteCSS := "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(2) > div:nth-child(4) > div > button"

; ---------------------------------------------------------------------------------------------------------------------
; Mycadent paths

global Path_CadentSearchFieldID := "ctl00_body_OrdersListReport_ctl01_ctl05_ctl05-string-operand"

global Path_CadentOrderNumberID := "ctl00_body_txtOrderHeaderID"

global Path_CadentExport := "ctl00_body_ucOrthoCadLink_OrthoCadLink"
