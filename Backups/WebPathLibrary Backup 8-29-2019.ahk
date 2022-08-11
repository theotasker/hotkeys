; ================================================================================================================
; This Library should contain all paths to objects on websites, editing here will update the main script
; ================================================================================================================


; ================================================================================================================
; RXWizard paths
; ================================================================================================================

Path_GenerateScheduleCSS = #main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button:nth-child(1)

Path_StartStopXPATH = //*[@id="main-content-wrapper"]/div[1]/main/div[2]/div/div[2]/div[2]/button

; If user has manager privileges on the website, the button location changes
Path_StartStopManagerXPATH = //*[@id="main-content-wrapper"]/div[1]/main/div[2]/div/div/div[2]/button[1]

Path_SaveNoteXPATH = /html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/button[2]

Path_UploadFileXPATH = //*[@id="main-content-wrapper"]/div/main/form/div[1]/div[1]/div[6]/div/div/div[2]/span/div[1]/span/div/p[2]

Path_ScanScriptCSS = .ant-input:nth-child(2)

Path_ReviewButtonCSS = #main-content-wrapper > div > main > form > div.ant-row-flex.ant-row-flex-end > div:nth-child(2) > button

Path_ScriptNumberCSS = #main-content-wrapper > div > div.header > div.name > div > span.case-name

Path_ClinicNameCSS = #office > div > div > div.ant-select-selection-selected-value

Path_NewNoteCSS = #main-content-wrapper > div > main > form > div.ant-row > div:nth-child(2) > div:nth-child(3) > div > button


; ================================================================================================================
; Mycadent paths
; ================================================================================================================

Path_CadentSearchFieldID = ctl00_body_OrdersListReport_ctl01_ctl05_ctl05-string-operand

Path_CadentOrderNumberID = ctl00_body_txtOrderHeaderID
