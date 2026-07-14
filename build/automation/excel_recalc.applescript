on run argv
    if (count of argv) is not 1 and (count of argv) is not 4 then
        error "Usage: excel_recalc.applescript workbook [sheet cell value]"
    end if

    set workbookPath to item 1 of argv
    do shell script "test -f " & quoted form of workbookPath
    set workbookFile to POSIX file workbookPath as text
    set workbookName to do shell script "basename " & quoted form of workbookPath
    set targetSheet to ""
    set targetCell to ""
    set targetValue to ""
    if (count of argv) is 4 then
        set targetSheet to item 2 of argv
        set targetCell to item 3 of argv
        set targetValue to item 4 of argv
    end if

    set openedWorkbook to missing value
    set workbookOpened to false
    set stateCaptured to false
    set calculationCaptured to false
    set cleanupProblems to ""
    set lastExcelError to ""

    try
        set excelReady to false
        repeat 40 times
            try
                tell application "Microsoft Excel"
                    set currentWorkbookNames to name of every workbook
                    set previousAlerts to display alerts as boolean
                end tell
                if currentWorkbookNames is missing value then set currentWorkbookNames to {}
                set excelReady to true
                exit repeat
            on error transientMessage number transientNumber
                set lastExcelError to "Error " & transientNumber & ": " & transientMessage
                delay 0.25
            end try
        end repeat
        if not excelReady then
            error "Excel did not become scriptable before full rebuild. " & lastExcelError
        end if
        if (count of currentWorkbookNames) is not 0 then
            set previousDelimiters to AppleScript's text item delimiters
            set AppleScript's text item delimiters to ", "
            set openWorkbookNames to currentWorkbookNames as text
            set AppleScript's text item delimiters to previousDelimiters
            error "Desktop Excel full rebuild requires no other open workbooks. Close: " & openWorkbookNames
        end if
        tell application "Microsoft Excel"
            set display alerts to false
        end tell
        set stateCaptured to true

        tell application "Microsoft Excel"
            open workbook workbook file name workbookFile update links do not update links read only false
        end tell
        repeat 120 times
            try
                tell application "Microsoft Excel"
                    set currentWorkbookNames to name of every workbook
                    if workbookName is in currentWorkbookNames then
                        set openedWorkbook to workbook workbookName
                        set workbookFullName to full name of openedWorkbook as text
                        set workbookOpened to true
                    end if
                end tell
                if workbookOpened then exit repeat
            on error transientMessage number transientNumber
                set lastExcelError to "Error " & transientNumber & ": " & transientMessage
            end try
            delay 0.5
        end repeat
        if not workbookOpened then
            error "Workbook did not become scriptable in Excel; an open or repair dialog may be blocking it: " & workbookName & ". " & lastExcelError
        end if

        tell application "Microsoft Excel"
            set previousCalculation to calculation
            if previousCalculation is missing value then
                error "Excel did not expose its calculation mode after opening " & workbookName
            end if
            set calculationCaptured to true

            if targetSheet is not "" then
                set value of range targetCell of worksheet targetSheet of openedWorkbook to targetValue
            end if

            set calculation to calculation automatic
            calculate full rebuild
            set sentinel to (value of range "M2" of worksheet "Calc" of openedWorkbook) as text
            if sentinel is not "All" then
                error "Recalculation sentinel Calc!M2 is " & quoted form of sentinel & "; expected All"
            end if

            save openedWorkbook
            set calculation to previousCalculation
            set calculationCaptured to false
            close openedWorkbook saving no
            set openedWorkbook to missing value
            set workbookOpened to false
            set display alerts to previousAlerts
            set stateCaptured to false
        end tell
    on error failureMessage number failureNumber
        if calculationCaptured then
            set calculationRestored to false
            repeat 40 times
                try
                    tell application "Microsoft Excel"
                        set calculation to previousCalculation
                        set currentCalculation to calculation
                    end tell
                    if currentCalculation is previousCalculation then
                        set calculationRestored to true
                        set calculationCaptured to false
                        exit repeat
                    end if
                on error cleanupMessage number cleanupNumber
                    set lastExcelError to "Error " & cleanupNumber & ": " & cleanupMessage
                end try
                delay 0.25
            end repeat
            if not calculationRestored then
                set cleanupProblems to cleanupProblems & " calculation restore failed after retries: " & lastExcelError
            end if
        end if
        if workbookOpened then
            set workbookClosed to false
            repeat 40 times
                try
                    tell application "Microsoft Excel"
                        set currentWorkbookNames to name of every workbook
                        if workbookName is in currentWorkbookNames then
                            close workbook workbookName saving no
                        end if
                    end tell
                    set workbookClosed to true
                    exit repeat
                on error cleanupMessage number cleanupNumber
                    set lastExcelError to "Error " & cleanupNumber & ": " & cleanupMessage
                    delay 0.25
                end try
            end repeat
            if workbookClosed then
                set workbookOpened to false
                set openedWorkbook to missing value
            else
                set cleanupProblems to cleanupProblems & " close failed after retries: " & lastExcelError
            end if
        end if
        if stateCaptured then
            set alertsRestored to false
            repeat 40 times
                try
                    tell application "Microsoft Excel"
                        set display alerts to previousAlerts
                        set currentAlerts to display alerts as boolean
                    end tell
                    if currentAlerts is previousAlerts then
                        set alertsRestored to true
                        exit repeat
                    end if
                on error cleanupMessage number cleanupNumber
                    set lastExcelError to "Error " & cleanupNumber & ": " & cleanupMessage
                end try
                delay 0.25
            end repeat
            if not alertsRestored then
                set cleanupProblems to cleanupProblems & " alert restore failed after retries: " & lastExcelError
            end if
        end if
        if cleanupProblems is not "" then
            error failureMessage & "; cleanup failures:" & cleanupProblems number failureNumber
        end if
        error failureMessage number failureNumber
    end try

    return "RECALCULATED sentinel=All"
end run
