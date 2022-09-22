; ================================================================================================================
; Paths for use by Selenium for RXWizard and MyCadent
; ================================================================================================================

/*
These locators are used by Selenium to find their associated elements.
Most of the RXWizard elements have been updated to CSS IDs instead of paths, as that's a more reliable handle
Tried to use CSS locators instead of absolute paths for the rest, but couldn't get them all

To update these, open the hotkeys browser and use the "inspect" tool in Chrome. 
*/

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
, "noteSave": "div[class='ant-modal modal-form'] button[class='ant-btn ant-btn-primary']"
, "":""}

global casesPageCSS := {"":""
, "assignedCases":"button[data-cy='Assigned to me']"
, "":""}

; Mycadent paths
global cadentCssID := {"":""
, "searchField":"ctl00_body_OrdersListReport_ctl01_ctl05_ctl05-string-operand"
, "orderID":"ctl00_body_txtOrderHeaderID"
, "exportLink":"ctl00_body_ucOrthoCadLink_OrthoCadLink"
, "":""}