#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client

currentURL := "https://portal.rxwizard.com/cases/edit/486689"


finalizeSTLs(finishOptions, existingArchFilenames, filenameBase) {
    if (finishOptions["arches"] = "both")
    {
        filenameTag := "[2].stl"
    }
    Else
    {
        filenameTag := "[1].stl"
    }

    if (finishOptions["auto"] = True)
    {
        destinationDir := autoImportDir
    }
    Else
    {
        destinationDir := tempModelsDir
    }

    if (finishOptions["upper"] = True)
    {
        currentFullFilename := tempModelsDir existingArchFilenames["upper"]
        destFullFilename := destinationDir filenameBase "Upr" filenameTag
        FileMove, %currentFullFilename%, %destFullFilename%
    }
    if (finishOptions["lower"] = True)
    {
        currentFullFilename := tempModelsDir existingArchFilenames["lower"]
        destFullFilename := destinationDir filenameBase "Lwr" filenameTag
        FileMove, %currentFullFilename%, %destFullFilename%
    }
    return
}