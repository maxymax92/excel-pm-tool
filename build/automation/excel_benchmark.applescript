use framework "Foundation"
use scripting additions

on secondsNow()
    return (current application's NSDate's timeIntervalSinceReferenceDate()) as real
end secondsNow

on millisecondsPerOperation(startedAt, finishedAt, operationCount)
    return ((finishedAt - startedAt) * 1000.0) / operationCount
end millisecondsPerOperation

on run argv
    if (count of argv) is not 1 then
        error "Usage: excel_benchmark.applescript workbook"
    end if

    set workbookPath to item 1 of argv
    do shell script "test -f " & quoted form of workbookPath
    set workbookFile to POSIX file workbookPath as text
    set workbookName to do shell script "basename " & quoted form of workbookPath

    set openedWorkbook to missing value
    set workbookOpened to false
    set stateCaptured to false
    set previousCalculation to missing value
    set cleanupProblems to ""
    set pollErrorCount to 0
    set lastPollError to ""

    try
        tell application "Microsoft Excel"
            if workbookName is in (name of every workbook) then
                error "A workbook named " & quoted form of workbookName & " is already open"
            end if
            set previousAlerts to display alerts
            set previousCalculation to calculation
            set stateCaptured to true
            set display alerts to false
            set calculation to calculation automatic
        end tell

        set openStarted to my secondsNow()
        tell application "Microsoft Excel"
            open workbook workbook file name workbookFile update links do not update links read only false
        end tell
        tell application "Microsoft Excel"
            repeat 120 times
                try
                    if workbookName is in (name of every workbook) then
                        set openedWorkbook to workbook workbookName
                        set workbookOpened to true
                        if ((value of range "M2" of worksheet "Calc" of openedWorkbook) as text) is "All" then
                            exit repeat
                        end if
                    end if
                on error pollMessage number pollNumber
                    set pollErrorCount to pollErrorCount + 1
                    set lastPollError to "Error " & pollNumber & ": " & pollMessage
                end try
                delay 0.05
            end repeat
            if not workbookOpened then
                error "Workbook did not register in Excel: " & workbookName & ¬
                    ". Last polling error: " & lastPollError
            end if
            if ((value of range "M2" of worksheet "Calc" of openedWorkbook) as text) is not "All" then
                error "Workbook did not expose the Calc!M2 readiness sentinel"
            end if
            set openFinished to my secondsNow()

            activate
            activate object worksheet "Overview" of openedWorkbook

            set tabStarted to my secondsNow()
            repeat 25 times
                activate object worksheet "Plan" of openedWorkbook
                activate object worksheet "Items" of openedWorkbook
                activate object worksheet "RAID" of openedWorkbook
                activate object worksheet "Overview" of openedWorkbook
            end repeat
            set tabFinished to my secondsNow()

            activate object worksheet "Items" of openedWorkbook
            set selectionStarted to my secondsNow()
            repeat 50 times
                select range "A3" of worksheet "Items" of openedWorkbook
                select range "C3" of worksheet "Items" of openedWorkbook
                select range "I3" of worksheet "Items" of openedWorkbook
                select range "K3" of worksheet "Items" of openedWorkbook
            end repeat
            set selectionFinished to my secondsNow()

            set originalStatus to value of range "J3" of worksheet "Items" of openedWorkbook
            set editStarted to my secondsNow()
            repeat 10 times
                set value of range "J3" of worksheet "Items" of openedWorkbook to "Performance probe"
                set value of range "J3" of worksheet "Items" of openedWorkbook to originalStatus
            end repeat
            set editFinished to my secondsNow()

            set originalHealth to value of range "I3" of worksheet "Items" of openedWorkbook
            if (originalHealth as text) is "On track" then
                set probeHealth to "At risk"
            else
                set probeHealth to "On track"
            end if
            set calculationEditStarted to my secondsNow()
            repeat 10 times
                set value of range "I3" of worksheet "Items" of openedWorkbook to probeHealth
                set value of range "I3" of worksheet "Items" of openedWorkbook to originalHealth
            end repeat
            set calculationEditFinished to my secondsNow()

            close openedWorkbook saving no
            set openedWorkbook to missing value
            set workbookOpened to false
            set calculation to previousCalculation
            set display alerts to previousAlerts
            if calculation is not previousCalculation then
                error "Excel calculation mode was not restored"
            end if
            set stateCaptured to false
        end tell

        set openMs to (openFinished - openStarted) * 1000.0
        set tabMs to my millisecondsPerOperation(tabStarted, tabFinished, 100)
        set selectionMs to my millisecondsPerOperation(selectionStarted, selectionFinished, 200)
        set editMs to my millisecondsPerOperation(editStarted, editFinished, 20)
        set calculationEditMs to my millisecondsPerOperation(¬
            calculationEditStarted, calculationEditFinished, 20)
        return "OPEN_MS=" & openMs & " OPEN_RETRIES=" & pollErrorCount & ¬
            " TAB_SWITCH_MS=" & tabMs & ¬
            " CELL_SELECT_MS=" & selectionMs & " EDIT_MS=" & editMs & ¬
            " CALC_EDIT_MS=" & calculationEditMs & ¬
            " CALCULATION_RESTORED=1"
    on error failureMessage number failureNumber
        tell application "Microsoft Excel"
            if workbookOpened then
                try
                    close workbook workbookName saving no
                    set workbookOpened to false
                on error cleanupMessage
                    set cleanupProblems to cleanupProblems & " close failed: " & cleanupMessage
                end try
            end if
            if stateCaptured then
                try
                    set calculation to previousCalculation
                on error cleanupMessage
                    set cleanupProblems to cleanupProblems & ¬
                        " calculation restore failed: " & cleanupMessage
                end try
                try
                    set display alerts to previousAlerts
                on error cleanupMessage
                    set cleanupProblems to cleanupProblems & " alert restore failed: " & cleanupMessage
                end try
            end if
        end tell
        if cleanupProblems is not "" then
            error failureMessage & "; cleanup failures:" & cleanupProblems number failureNumber
        end if
        error failureMessage number failureNumber
    end try
end run
