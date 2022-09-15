; ================================================================================================================
; Selenium Paths, need updating as RXWizard gets updated
; ================================================================================================================

; ---------------------------------------------------------------------------------------------------------------------
; RXWizard Paths

global reviewPageCSS := {"":""
, "generateSchedule": "#main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button:nth-child(1)"
, "startStop": "#main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button"
, "panNumber": "#main-content-wrapper > div.page-content > div > div.name > div > span.pan > div > span"
, "clinicName": "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(1) > div.ant-col.ant-col-17 > div > div"
, "patientName": "#main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(8) > div.ant-col.ant-col-17 > div > div"
,"":""}

global bothPageCSS := {"":""
, "uploadFile": "div[data-cy='Case files'] div[class='ant-col ant-col-xxl-12'] div[class='ant-upload-drag-container']"
, "newNote": "div[data-cy='Notes'] button[class='ant-btn ant-btn-primary']"
, "noteSave": "div[role='document'] button[class='ant-btn ant-btn-primary']"
, "":""}

; ---------------------------------------------------------------------------------------------------------------------
; Mycadent paths

global Path_CadentSearchFieldID := "ctl00_body_OrdersListReport_ctl01_ctl05_ctl05-string-operand"

global Path_CadentOrderNumberID := "ctl00_body_txtOrderHeaderID"

global Path_CadentExport := "ctl00_body_ucOrthoCadLink_OrthoCadLink"
