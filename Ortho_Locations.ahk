; ===========================================================================================================================
; Constants for positional elements in 3Shape
; ===========================================================================================================================

/*
The 3shapeButtons array is a list of X and Y coordinates for all used clickable buttons in 3shapeButtons
The 3shapeFields array is all of the edit fields and buttons that can be seen by AHK

Use "Window Spy" software (bundled with AutoHotKey) to get these. Coordinates are based on "client"
*/

global 3shapeButtons := {"":""
, "allViewX":"1902"
, "topViewY":"286"
, "bottomViewY":"253"
, "rightViewY":"215"
, "leftViewY":"177"
, "frontViewY":"106"
, "backViewY":"140"
, "transparencyX":"800"
, "transparencyY":"800"
, "":""
, "artifactX":"111"
, "artifacty":"54"
, "artifactApplyX":"100"
, "artifactApplyY":"210"
, "planeCutX":"150"
, "planeCutY":"55"
, "planeCutApplyX":"190"
, "planeCutApplyY":"230"
, "splineCutX":"195"
, "splineCutY":"50"
, "splineSmoothX":"180"
, "splineSmoothY":"180"
, "splineCutApplyX":"187"
, "splineCutApplyY":"260"
, "waxKnifeX":"35"
, "waxKnifeY":"175"
, "nextButtonX":"190"
, "nextButtonY":"30"
, "":""
, "newPatientModelX":"78"
, "newPatientModelY":"46"
, "patientBrowserX":"27"
, "patientBrowserY":"43"
, "":""}

global 3shapeFields := {"":""
, "advSearchFirst":"TEdit10"
, "advSearchLast":"TEdit9"
, "advSearchClinic":"Edit1"
, "advSearchScript":""
, "advSearchScript2019":"TEdit16"
, "advSearchScript2021":"TEdit17"
, "advSearchScriptX":"1135"
, "advSearchScriptY":"119"
, "advSearchGo":"TButton2"
, "newPatientExt":"TEdit4"
, "newPatientFirst":"TEdit10"
, "newPatientLast":"TEdit9"
, "newPatientClinic":"Edit1"
, "newModelScript":"TEdit5"
, "":""}