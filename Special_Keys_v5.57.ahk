/*
_______________________________________________________________________________________________________________
Special KB - Allows special characters string (ex-A!!) based inventory items (correspond to special keys) to be
used with the CRE.
How it works - Scripts basically takes a number appended with Special character string as input and sends action
based on the character string accordingly (Ex- negate in case of BXX and otherwise)

Update 1 - In case if a user presses the same key again (ex - A!! after pressing the 100A!! key), Our script
should be able to repeat the previous/* order (This will keep the transaction intact for accidental key presses)

Update 2 - Since, we do not want any side effects from other AHK scripts operating on the same hotkeys therefore
merging them into single script makes sense and this script now includes both `ok` and `@` shortcuts as well

Update 3 - Eili is using a special keyboard for all the hotstrings and one issue that we constantly run into is
the fact that sometimes a user can outperforms the script execution due to the hardware (n-key-rollover) and that
can mess up the whole flow. In order to take care of those special cases -- we are going with a single hotkey
action instead of hotstrings

Why? - Because it's quick and we will never have to do the complex actions before executing the action.
Allowed Hotkeys - !@#$%^&*()_+-=[]\;'<>?,./

Update 4 - Let's use the simple alphabetic characters instead of special characters and make them only to get
into action whenever they are typed inside a specific text field

Update 5 - The script should be able to send keys to the cash sub-window for efficiency purposes. Though it should
be limited to cash, EBT, credit and debit

Update 6 - Script should be able to handle cashback and other operations with a robust backend of MS Sql instead of
second guessing it via our own logic. Basically Integrate the DB data with the script for a better Overall Design

Update 7 - For a very long time this script was keep on changing but, the changes are not being tracked on
any version control system which means -- this is going to be a problem in long run. So here is the repository
https://github.com/okkosh/EiliKB (private repo) to keep track of all the changes for better code quality

Update 8 - The script should be able to create SVC2 based on the amount

Update 9 -Added a new `n` key functionality where all the SVC based logic is bypassed to make sure things
are working as expected

Update 10 - Add reporting functionality inside the script for better
_______________________________________________________________________________________________________________

SPECIAL NOTES:

Q - Why aren't we creating hotkeys using the "Hotkey" command, It's simple, dynamic and easy on the eyes?
A - Well, AHK does not recommend using that method for performance reasons and especially due to the fact
that they are mainly for modifying hotkeys that are not known to the script (Dynamic keys).

"Creating hotkeys via double-colon labels performs better than using the Hotkey command because the hotkeys can
all be enabled as a batch when the script starts (rather than one by one). Therefore, it is best to use this
command to create only those hotkeys whose key names are not known until after the script has started running.
One such case is when a script's hotkeys for various actions are configurable via an INI file.

I know it's sort of an ugly action but, you gotta do what you gotta do for that performance :)
*/

#NoEnv
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetBatchLines, -1
SetKeyDelay -1
#HotString SI
; OPTIMIZATION ENDS

Global ADOSQL_LastError, ADOSQL_LastQuery ; These super-globals are for debugging your SQL queries.
Num =	; Number to form on keypresses (see below for more information)
prev_hotkey = ; Previous hotkey to keep track of
prev_num = ; Previous number to be saved
CreWin := "ahk_exe CRE2004.exe" ; Target CRE Window name
run_as_admin()

; Hide window on touch input

min_debit := 15 ; Minimum debit amount (Can be false to ignore this flag)
max_debit := 5000.0 ; Maximum debit amount that should bypass the SVC
msg_threshold := 75.0 ; Threshold amount that should show the Signature msg Popup (Applies only in credit mode)

; Limit message - Set it to false for no alert and set it to a
; number for setting max amount limit per item
; crr - set it to 10000 - for 100 dollars
limit_item := 100000

;Exempt limit on followings hotkeys
; use a comma separated hotkey (NO SPACE) ex - a,b,c
exempted_keys =

; Disable print on failed transactions
disable_print := true

; SVC2 range from start-end for credit cards
svc2_credit_range_start := 20
svc2_credit_range_end := 2000000000000000000

; SVC2 range from start-end for debit cards
svc2_debit_range_start := 1500000000000
svc2_debit_range_end := 2500000000000000

; SVC2 Values
svc2_credit_percentage := 0.025 ; Credit SVC percentage
svc2_debit_percentage := 0.05 ; Debit SVC percentage

ebt_svc1 := true ; EBT option to add SVC 1 set (true) or for bypass set (false)

is_client := false ; client flag
client_address := "192.168.1.96,32001" ; server address

enable_grand_total_gui := true ; set false for disable

enable_recipt_on_credit_debit := true ; set false to disable

global price_change_gui := true ; set false for disable

auto_print_daily_month_report := true ; set false for disable
daily_report_print_time := "030000" ; HHmmss 24HOURS FORMAT

CARD_ENTRY_METHOD := {1: "Manual", 2: "Swiped", 7: "Contactless EMV", 10: "CHIP", 11: "Fallback Swipe"}

; Printing client app Location (Used for printing reports)
PRINT_EXE := """C:\Program Files (x86)\Auto POS\autopos.exe"" --cli -p "
EXE_PATH := "C:\Program Files (x86)\Auto POS\"

CASH_WIN := "Custom Platform and Payment"
CHECK21_WIN := "Select Birthday"
INVENTORY_WIN := "Optional Info"
WIN_TAG_ALONG := "Tag Along Item"
BTN_CTRL := "WindowsForms10.BUTTON.app."
EDIT_CTRL := "WindowsForms10.EDIT.app."
POP_PAY_CTRL := "WindowsForms10.Window.8.app."
TAG_ALONG_ADD_BTN := ""
TAG_ALONG_REM_BTN := ""

TAG_ALONG_WIN := "Inventory Maintenance"
TAG_ALONG_SEARCH_WIN := "Search Inventory"

ERROR_MESSAGE := false

; Invoices Variables
LAST_INVOICE := 0
CURRENT_INVOICE := 1

; Inventory
CASHBACK_ITEM_NUM := "e2"

; Date editbox id for reporting screen automation
START_DATE_EDITBOX := ""
END_DATE_EDITBOX := ""
REPORT_WIN := "Detailed Daily Report"
ITEM_NF_WIN := "Item Not Found!"
; Date Editbox
start_date_editbox_suffix := "_ad14"
end_date_editbox_suffix := "_ad13"
; Pop pay/Online pay
pop_pay_btn_suffix := "_ad13"
online_pay_btn_suffix := "_ad14"

; Responsive Total GUI
ScreenWidth := A_ScreenWidth
ScreenHeight := A_ScreenHeight

GuiWidth := ScreenWidth * 0.27
GuiHeight := ScreenHeight * 0.05

GuiX := ScreenWidth * 0.57
GuiY := ScreenHeight * 0.12

print_receipt_x := ScreenWidth * 0.004
print_receipt_y := ScreenHeight * 0.55

print_receipt_width := ScreenWidth * 0.55
print_receipt_height := ScreenHeight * 0.15

global blockTrigger := false

; Partial cash flag (wether the cash is paid completely or not)
global PARTIAL_CASH := false
; Create the amounts based on allowed values
Amount_non_tax := [0.05, 0.10, 0.30, 0.60, 1.50, 1.80]
Amount_tax := [0.05, 0.10, 0.30, 0.60]
Amount_tax_beer := [0.05, 0.10, 0.20, 0.30, 0.60, 0.75, 0.90, 1.00, 1.20, 1.60, 1.80]
Amount_tax_wine := [0.05, 0.10, 0.20, 0.25]
Amount_tax_liquor := [0.05, 0.10, 0.50]
All_amounts := [0.05, 0.10, 0.20, 0.25, 0.30, 0.50, 0.60, 0.75, 0.90, 1.00, 1.20, 1.50, 1.60, 1.80] ; Should contain all the numbers possible

global db_ := new DB
; db_.init("username","pass","database","server") leaving empty will use default values
; db_.init("sa", "pcAmer1ca", "cresql", "192.168.010.153.,30021\\PCAMERICA")

if (is_client){
    db_.init("sa", "pcAmer1ca", "cresql", "" . client_address . "\\PCAMERICA")
} else{
    db_.init()
}

; Check if the database is up and running by quering the store info
global store_info := db_.check_db()
if (not store_info)
    splash_box("Mismatched DB. Press any key to continue...")

; Gui for selecting tag along items based on the user preferences
Gui, font, s14
Gui, Add, Text, x0 y19 h30 w325 +center, Select tag along Item
Gui, Add, Radio, x100 y60 w150 h100 vRadio 0x1000 0x1 gRadioCtrl, Non-Taxable ; CRVA2
Gui, Add, Radio, xp+150 y60 w150 h100 0x1000 0x1 gRadioCtrl, Taxable ; CRVA1
Gui, Add, Radio, xp+150 y60 w150 h100 0x1000 0x1 gRadioCtrl, Beer ;CRVA4
Gui, Add, Radio, x100 y160 w150 h100 0x1000 0x1 gRadioCtrl, Wine
Gui, Add, Radio, xp+150 y160 w150 h100 0x1000 0x1 gRadioCtrl, Liquor
Gui, Add, Text, x10 yp+105 w325 h30 +center vSelText, Choose an amount

; Setup a starting point for automatically creating a button UI
x_ := -140
y_ := 300

; Loop through all amount and populate buttons for the selection
Loop % All_amounts.Length()
{
    x_ := x_ + 150
    if (x_ >= 480) {
        x_ := 10
        y_ := y_ + 100
    }
    Gui, Add, Button, x%x_% y%y_% w150 h100 0x7 vBtn_%A_index% gSave, % "$" . All_amounts[A_Index]
    ; Make all the buttons disabled initally
    GuiControl, disable, Btn_%A_index%
}
Gui, Show, , % WIN_TAG_ALONG
Gui +AlwaysOnTop

; No SVC screen GUI
Gui, no_svc:new
Gui, no_svc:font, s14
Gui, no_svc:Add, Button, x21 y16 w192 h86 gBtnCredit, CREDIT
Gui, no_svc:Add, Button, x232 yp w192 h86 gBtnDebit , DEBIT
Gui, no_svc:Add, Button, x444 yp w201 h86 gBtnCancel, CANCEL
Gui, no_svc:+AlwaysOnTop

; Reports Screen GUI
Gui, reports:new
Gui, reports:font, s14
Gui, reports:Add, Button, x22 y30 w220 h120 gBtnMonthToDate, Month-to-date report
Gui, reports:Add, Button, xp+230 yp w220 h120 gBtnTodayReport, Today's report
Gui, reports:Add, Button, x22 yp+130 w220 h120 gBtnLastMonth, Last month report
Gui, reports:Add, Button, xp+230 yp w220 h120 gBtnYestReport, Yesterday's report
Gui, reports:Add, Button, x22 yp+130 w220 h120 gBtnCustomDate, Custom Date Report
Gui, reports:+AlwaysOnTop

; Total Screen GUI
Gui, TotalGUI:New, -Caption
Gui, TotalGUI:Font, s30 c18db28, Arial
Gui, TotalGUI:Add, Text, vTotal x0 y0 w%GuiWidth% h%GuiHeight%, Grand Totals
Gui, TotalGUI:Color, black
Gui, TotalGUI:+AlwaysOnTop

;Debt, Credit custom receipt
Gui, custom_receipt:New, -Caption
Gui, custom_receipt:Font, s20, Arial
Gui, custom_receipt:Add, Progress, x0 y0 w138 h105 Disabled Background4683b4 vCashRegisterReceiptProgress
Gui, custom_receipt:Add, Text, xp yp wp hp BackgroundTrans gCashRegisterReceipt vCashRegisterReceiptText, Cash Register Receipt
Gui, custom_receipt:Add, Progress, x150 y0 w138 h105 Disabled Background4683b4 vCardReceiptProgress
Gui, custom_receipt:Add, Text, xp yp wp hp BackgroundTrans gCardReceipt vCardReceiptText, Card Receipt
Gui, custom_receipt:Add, Progress, xp+200 yp w450 h110 Disabled
Gui, custom_receipt:Add, Text, xp yp wp hp BackgroundTrans c18db28, % Chr(0x2611) " successful Transaction"
Gui, custom_receipt:Add, Text, xp yp+35 wp hp vCashBack BackgroundTrans c18db28 ,
Gui, custom_receipt:Add, Text, xp yp+30 wp hp vSignature BackgroundTrans c18db28 ,

Gui, custom_receipt: +AlwaysOnTop

is_price_change := False
;Item Price update
Gui, update_price:New, +AlwaysOnTop -0x20000
Gui, update_price:Font, s30, Arial
Gui, update_price:Add, GroupBox, x40 y20 w520 h130, Scan an Item
Gui, update_price:Add, Edit, x50 y90 w500 h50 vItemNumber,
Gui, update_price:Add, Button, x0 y0 w10 h10 Default gCheckItemExists Hidden, OK
Gui, price_change:New, +AlwaysOnTop -0x20000
Gui, price_change:font, s14
Gui, price_change:Add, Text, x200 y10 w250 vItemNum, ItemNum
Gui, price_change:Add, Text, x20 y10 , Item Number
Gui, price_change:Add, Text, x200 y36 w250 vItemName, ItemName
Gui, price_change:Add, Text, x20 y36 , Item Name
Gui, price_change:Add, Edit, x200 y86 w250 vCPrice gPriceEdit,
Gui, price_change:Add, Text, x20 y86 w150 , Before Tax Price
Gui, price_change:Add, Edit, x200 y130 w250 vAPrice gTaxPriceEdit,
Gui, price_change:Add, Text, x20 y130 w150 , After Tax Price
Gui, price_change:Add, CheckBox, x20 y180 w80 vTax gTax, Tax
Gui, price_change:Add, CheckBox, x20 y210 w80 vTax2 gTax2, Tax 2
Gui, price_change:Add, ListBox, x200 y210 w250 vTagAlongItem,
Gui, price_change:Add, Text, x200 y180 , Tag Along Items

; Gui, price_change:Add, Button, x120 y310 w70 h40 gPriceAdded, SAVE

Gui, price_change:Add, CheckBox, x20 y260 w100 vFoodstampable, Foodstampable

Gui, price_change:Add, Button, x120 y310 w70 h40 gPriceAdded Default, SAVE
Gui, price_change:Add, Button, x220 y310 w70 h40 gPriceCancel, Cancel


GuiClose:
    WinHide, Tag Along Item

if auto_print_daily_month_report{
    beeped := false ; Initialize a flag to track beeping
    SetTimer, CheckTime, 1000 ; Check every sec initially
}

CheckTime:
    if winexist("Run Time Support") {
        if !beeped{
            WinActivate, "Run Time Support"
            WinGetText, txt, % "Run Time Support"
            MESSAGE := txt
            if InStr(MESSAGE, "You have not entered a valid date. Please enter a valid date format(MM/dd/yyyy or MM-dd-yyyy)") {
                SoundPlay, *-1
                beeped := true
            }
        }

    } else {
        beeped := false
    }
    FormatTime, CurrentTime,, HHmmss ; Get the current time in HHMM format
    FormatTime, CurrentDate,, d
    If (CurrentTime = daily_report_print_time)
    {
        reporting()
        global REPORT_WIN
        win_id := get_window_class_id(Crewin)
        START_DATE_EDITBOX := EDIT_CTRL . "" . win_id . "" . start_date_editbox_suffix
        END_DATE_EDITBOX := EDIT_CTRL . "" . win_id . "" . end_date_editbox_suffix
        hide_win()
        prerequisites()
        L_day_flag := False
        L_day := A_DD -1
        Month := A_MM
        Year := A_YYYY
        if (L_day <= 0){
            Month := Month - 1
            if (Month = 0){
                Month := 12
                Year := Year - 1
            }
            L_day := last_day(Month)
            L_day_flag := True
        }
        WinWait, % REPORT_WIN, , 3

        StartDate := Month . "/" . L_day . "/" . Year
        EndDate := L_day_flag ? A_MM . "/" . A_DD . "/" . A_YYYY : Month . "/" . (L_day + 1) . "/" . Year

        ControlSetText, % START_DATE_EDITBOX, % StartDate, % REPORT_WIN
        ControlSetText, % END_DATE_EDITBOX, % EndDate, % REPORT_WIN
        sendInput, !p
        sleep, 1000
        sendInput, !x
        sleep, 1000
        sendInput, !x
        sleep, 1000
        sendInput, !x
        sleep, 1000
        sendInput, !x
        sleep, 1000
        If (CurrentDate = 1) {
            hide_win()
            prerequisites()
            Month := A_MM
            Year := A_YYYY
            if (A_MM = 01){
                Month := 12
                Year := A_YYYY -1
                StartDate := Month . "/01/" . Year
                EndDate := A_MM . "/01/" . A_YYYY
            } else {
                Month := A_MM -1
                StartDate := Month . "/01/" . Year
                EndDate := (Month+1) . "/01/" . Year
            }
            WinActivate, %REPORT_WIN%
            WinWait, % REPORT_WIN, , 5
            ControlSetText, % START_DATE_EDITBOX, % StartDate, % REPORT_WIN
            ControlSetText, % END_DATE_EDITBOX, % EndDate, % REPORT_WIN
            sendInput, !p
            sleep, 1000
            sendInput, !x
            sleep, 1000
            sendInput, !x
            sleep, 1000
            sendInput, !x
            sleep, 1000
            sendInput, !x
            sleep, 1000
        }
    }
return
    ; Do not exit the app on closing GUI instead hide the entire UI in BG


!h:: ;open label printer
    global is_client, client_address
    LABEL_EXE := """C:\Program Files (x86)\Label Print\LabelPrint.exe"" -link " . client_address
    EXE := "C:\Program Files (x86)\Label Print"
    if(is_client){
        RunWait, % ComSpec . " /c """ . LABEL_EXE . " """, % EXE, hide
        return
    }
    Run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Money Machine Inc\Label Print.lnk"
return

!m:: ;open actual cash gui
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt
    LABEL_EXE := """C:\Program Files (x86)\Auto POS\AutoPOS.exe"" --cli -scash "
    EXE := "C:\Program Files (x86)\Auto POS"
    RunWait, % ComSpec . " /c """ . LABEL_EXE . " """, % EXE, hide
return

!7:: ;open az barcode gui
    LABEL_EXE := """C:\Program Files (x86)\Auto POS\AutoPOS.exe"" --cli -az "
    EXE := "C:\Program Files (x86)\Auto POS"
    RunWait, % ComSpec . " /c """ . LABEL_EXE . " """, % EXE, hide
return
!1::
    msgbox, hello
return

; item price update
!v::
    Gui, update_price:Show, , Update Price
return
; Automatically populate the options based on radio selection
RadioCtrl:
    Gui, submit, Nohide
    Amount_selected := ""
    switch (Radio) {
    case 1:
        Amount_selected := Amount_non_tax
        GuiControl, , SelText, Select Non-Taxable Amount
    case 2:
        Amount_selected := Amount_tax
        GuiControl, , SelText, Select Taxable Amount
    case 3:
        Amount_selected := Amount_tax_beer
        GuiControl, , SelText, Select Beer Amount
    case 4:
        Amount_selected := Amount_tax_wine
        GuiControl, , SelText, Select Wine Amount
    case 5:
        Amount_selected := Amount_tax_liquor
        GuiControl, , SelText, Select Liquor Amount
    }

    ; Currently this is o(n2) efficient. We can make use of built-in
    ; comparison for better performance
    loop % All_amounts.length()
    {
        curr_index := A_Index
        GuiControl, Disable, Btn_%curr_index%
        loop % Amount_selected.length()
        {
            if (Amount_selected[A_index] == All_amounts[curr_index]) {
                GuiControl, Enable, Btn_%curr_index%
            }
        }
    }
    return

    ;On save
    Save:
        if (is_price_change){
            Ctrl_index := Substr(A_GuiControl, 5)
            Choice := All_amounts[Ctrl_index]
            Final_Code := ""

            ;TODO: use a code list from CSV or excel file instead of directly appending the harcoded items
            switch (Radio) {
            case 1:
                Final_Code = $%Choice%  NON-TAX
            case 2:
                Final_Code = $%Choice%  SODA
            case 3:
                Final_Code = $%Choice%  BEER
            case 4:
                Final_Code = $%Choice%  WINE
            case 5:
                Final_Code = $%Choice%  LIQUOR
            }
            sql_query := "Select ItemNum FROM Inventory WHERE ItemNum LIKE '%" . Final_Code . "%'"

            tag := db_.execute(sql_query)

            tag_alongs := ""
            tag_alongs := tag[2,1]

            GuiControl, price_change:, TagAlongItem, |
            GuiControl, price_change:, TagAlongItem, %tag_alongs%
            WinHide, Tag Along Item
            return
        }
        Gui, Submit, Nohide
        ; Remove any previous entry if already exists
        loop, 5 ; TODO: use listview to remove items present in Tag along instead of hardcoding 5 times
            ControlClick, % TAG_ALONG_REM_BTN, % Crewin
        ; A button control contains "Btn_" as a prefix therefore removing it from the
        ; Current control will reveal the index of amount selected
        Ctrl_index := Substr(A_GuiControl, 5)
        Choice := All_amounts[Ctrl_index]
        Final_Code := ""

        ;TODO: use a code list from CSV or excel file instead of directly appending the harcoded items
        switch (Radio) {
        case 1:
            Final_Code = $%Choice% NON-TAX CRV
        case 2:
            Final_Code = $%Choice% SODA CRV
        case 3:
            Final_Code = $%Choice% BEER CRV
        case 4:
            Final_Code = $%Choice% WINE CRV
        case 5:
            Final_Code = $%Choice% LIQUOR CRV
        }
        ControlClick, % TAG_ALONG_ADD_BTN, % TAG_ALONG_WIN
        WinWait , % TAG_ALONG_SEARCH_WIN , , 5	; Wait for the window to appear before sending final code
        send, %Final_code%
        Send, {Enter}
        Send, !l
        Goto, GuiClose
        return

        #IF WinActive(CreWin) or (WinActive("Price Change")) ; Valid only if the windows criteria matches
        !k:: ; Tag along item
        if (is_price_change){
            WinShow, % WIN_TAG_ALONG
            return
        }
        if !is_main_window(){
            if is_window_active(INVENTORY_WIN) {
/**
Tag along buttons ctrl Id changes based on the action type in stock ex - add or update
To check whether our buttons have changed we use the `check_suffix` control id _ad119
if this control contains a text i.e. Print Labels then it implies that suffix 1 & 2 will be used for
adding and removing buttons respectively otherwise 2 & 3 will be used
                */

                tag_along_suffix1 := "_ad115" ; Can be Add
                tag_along_suffix2 := "_ad116" ; Can be Add/Remove
                tag_along_suffix3 := "_ad117" ; Can be Add
                check_suffix := "_ad119"

                win_id := get_window_class_id("Inventory Maintenance")

                ControlGetText, text_ctrl, % BTN_CTRL . "" . win_id . "" . check_suffix, % CreWin

                if (text_ctrl == "Print Labels"){
                    TAG_ALONG_ADD_BTN := BTN_CTRL . "" . win_id . "" . tag_along_suffix1
                    TAG_ALONG_REM_BTN := BTN_CTRL . "" . win_id . "" . tag_along_suffix2
                } else {
                    TAG_ALONG_ADD_BTN := BTN_CTRL . "" . win_id . "" . tag_along_suffix2
                    TAG_ALONG_REM_BTN := BTN_CTRL . "" . win_id . "" . tag_along_suffix3
                }

                WinShow, % WIN_TAG_ALONG
            }
        }
    return
#IfWinActive

#IF WinActive(CreWin) and (not WinActive(CHECK21_WIN)) ; Valid only if the windows criteria matches
    +x::
x::	; Advance Unlimited
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt
    Num := get_number()
    if is_main_window(){
        total_amount := get_price()
        if (!total_amount){
            return
        }
        if not Num {
            Send, !p!c
            sleep, 1000
            if(enable_grand_total_gui){
                grand_total := "Grand Total $" . total_amount
                GuiControl, TotalGUI:, Total, %grand_total%
                Gui, TotalGUI:Show, x%GuiX% y%GuiY% w%GuiWidth% h%GuiHeight%, GrandTotal
                sleep, 300
                WinActivate, %Crewin%
                Suspend, On
                Input, SingleKey,L1V,{LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause},
                Suspend, Off
                WinHide, GrandTotal
                If InStr(ErrorLevel, "EndKey:")
                {
                    return
                }
                if (SingleKey = "x")
                {
                    Send, ^a{BackSpace}
                    Gosub, x
                }
                if (SingleKey = "e")
                {
                    Send, ^a{BackSpace}
                    Gosub, e
                }
                if (SingleKey = "c")
                {
                    Send, ^a{BackSpace}
                    Gosub, c
                }
                if (SingleKey = "d")
                {
                    Send, ^a{BackSpace}
                    Gosub, d
                }
                if (SingleKey = "w")
                {
                    Send, ^a{BackSpace}
                    Gosub, w
                }
                if (SingleKey = "s")
                {
                    Send, ^a{BackSpace}
                    Gosub, s
                }
                if (SingleKey = "n")
                {
                    Send, ^a{BackSpace}
                    Gosub, n
                }
            }
        } else {
            critical, on
            Send,% "{BS " StrLen(Num) "}"
            Send, !p
            Send, % Format("{:.2f}", Num/100)
            Send, !c
            sleep, 500
            if is_window_active(CASH_WIN){
                PARTIAL_CASH := true
            }
            critical, off
        }
    } else {
        if is_window_active(CASH_WIN){
            Send, !c
        } else {
            send, X
        }
        PARTIAL_CASH := false
    }
    Num =
return

!q:: ; reporting
    reporting(){
        global is_main_window, get_window_class_id, EDIT_CTRL, start_date_editbox_suffix, end_date_editbox_suffix
        if is_main_window(){
            win_id := get_window_class_id(Crewin)
            START_DATE_EDITBOX := EDIT_CTRL . "" . win_id . "" . start_date_editbox_suffix
            END_DATE_EDITBOX := EDIT_CTRL . "" . win_id . "" . end_date_editbox_suffix
            Gui, reports:Show, w500 h500, Report
        }
    }
return

;GUI buttons for reporting script
BtnMonthToDate:
    global REPORT_WIN, START_DATE_EDITBOX, END_DATE_EDITBOX
    hide_win()
    prerequisites()
    WinWait, % REPORT_WIN, , 3

    LastDayOfMonth := last_day(A_MM)
    if(A_DD+1 > LastDayOfMonth){
        ; Adjust the date to the first day of the next month
        NextMonth := A_MM + 1
        if (NextMonth > 12){
            NextMonth := 1
            NextYear := A_YYYY + 1
        } else {
            NextYear := A_YYYY
        }
        StartDate := A_MM . "/01/" . A_YYYY
        EndDate := NextMonth . "/01/" . NextYear
    }else {
        StartDate := A_MM . "/01/" . A_YYYY
        EndDate := A_MM . "/" . (A_DD+1) . "/" . A_YYYY
    }

    ControlSetText, % START_DATE_EDITBOX, % StartDate, % REPORT_WIN
    ControlSetText, % END_DATE_EDITBOX, % EndDate, % REPORT_WIN
    sendInput, !p
return

BtnLastMonth:
    global REPORT_WIN, START_DATE_EDITBOX, END_DATE_EDITBOX
    hide_win()
    prerequisites()

    Month := A_MM
    Year := A_YYYY
    if (A_MM = 01){
        Month := 12
        Year := A_YYYY -1
        StartDate := Month . "/01/" . Year
        EndDate := A_MM . "/01/" . A_YYYY
    } else {
        Month := A_MM -1
        StartDate := Month . "/01/" . Year
        EndDate := (Month+1) . "/01/" . Year
    }
    WinActivate, % REPORT_WIN
    WinWait, % REPORT_WIN, , 5
    ControlSetText, % START_DATE_EDITBOX, % StartDate, % REPORT_WIN
    ControlSetText, % END_DATE_EDITBOX, % EndDate, % REPORT_WIN
    sendInput, !p
return

BtnYestReport:
    global REPORT_WIN, START_DATE_EDITBOX, END_DATE_EDITBOX
    hide_win()
    prerequisites()
    L_day_flag := False
    L_day := A_DD -1
    Month := A_MM
    Year := A_YYYY
    if (L_day <= 0){
        Month := Month - 1
        if (Month = 0){
            Month := 12
            Year := Year - 1
        }
        L_day := last_day(Month)
        L_day_flag := True
    }
    WinWait, % REPORT_WIN, , 3

    StartDate := Month . "/" . L_day . "/" . Year
    EndDate := L_day_flag ? A_MM . "/" . A_DD . "/" . A_YYYY : Month . "/" . (L_day + 1) . "/" . Year

    ControlSetText, % START_DATE_EDITBOX, % StartDate, % REPORT_WIN
    ControlSetText, % END_DATE_EDITBOX, % EndDate, % REPORT_WIN
    sendInput, !p
return

BtnTodayReport:
    global REPORT_WIN, START_DATE_EDITBOX, END_DATE_EDITBOX

    hide_win()
    prerequisites()
    WinWait, % REPORT_WIN, , 3
    ; Get the last day of the current month
    LastDayOfMonth := last_day(A_MM)

    ; Check if adding 1 to the current day exceeds the last day of the month
    if (A_DD + 1 > LastDayOfMonth){
        ; Adjust the date to the first day of the next month
        NextMonth := A_MM + 1
        if (NextMonth > 12){
            NextMonth := 1
            NextYear := A_YYYY + 1
        } else {
            NextYear := A_YYYY
        }
        StartDate := A_MM . "/" . A_DD . "/" . A_YYYY
        EndDate := NextMonth . "/1/" . NextYear
    } else {
        ; Use the next day in the same month
        StartDate := A_MM . "/" . A_DD . "/" . A_YYYY
        EndDate := A_MM . "/" . (A_DD + 1) . "/" . A_YYYY
    }

    WinWait, % REPORT_WIN, , 3
    ControlSetText, % START_DATE_EDITBOX, % StartDate, % REPORT_WIN
    ControlSetText, % END_DATE_EDITBOX, % EndDate, % REPORT_WIN

    sendInput, !p
return

BtnCustomDate:
    hide_win()
    prerequisites()
return

; hides the report window
hide_win(){
    WinHide, Report
}

; Some repetitive steps that are necessary to reach the reporting screen
prerequisites(){
    global Crewin
    WinActivate , % Crewin
    sendInput, !o
    sleep, 500
    sendInput, 5
    sleep, 500
    sendInput, l
    sleep, 500
    SendInput, d
    sleep, 500
    SendInput, d
    sleep, 500
    SendInput, !d
}

; Method to get last day of any month
last_day(month){
    switch (month){
    case 1,3,5,7,8,10,12:
    return 31
case 4,6,9,11:
return 30
case 2:
    if ((Mod(A_YYYY,4) == 0) && (Mod(A_YYYY,4) || Mod(A_YYYY,100)!= 0)){
        return 29
    }
return 28
}
}

+c::
c::
credit:
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt
    if is_main_window() {
        PARTIAL_CASH := false
        price := get_price()
        if (!price){
            send, C
            return
        }

        if (price >= svc2_credit_range_start and price < svc2_credit_range_end) {
            handle_svc_keys(Format("{:.2f}", price*svc2_credit_percentage), true, false)
        } else if (price >= svc2_credit_range_end){
            handle_svc_keys("SVC3", true, false)
        } else {
            handle_svc_keys("SVC1", true, false)
        }
    } else {
        if is_window_active(CASH_WIN){
            if (PARTIAL_CASH){
                price := get_price(true)
                Send, !n
                total_price := get_price()
                if (price >= svc2_credit_range_start and price < svc2_credit_range_end) {
                    svc_value := Format("{:.2f}", price*svc2_credit_percentage)
                } else if (price >= svc2_credit_range_end){
                    svc_value := "SVC3"
                } else {
                    svc_value := "SVC1"
                }

                if (svc_value){
                    if InStr(svc_value, "SVC"){
                        ; Send the SVC command
                        send, %svc_value%{enter}
                    } else {
                        send, SVC2{Enter}
                        send, %svc_value%{Enter}
                    }
                    PARTIAL_CASH := false
                }
                sleep, 300
                cash_paid := Format("{:.2f}", (total_price - price))
                send, !p
                sleep, 300
                send, %cash_paid%
                send, !c
            }
            Send, !r
            Send, !r
            ; Starts the transaction
            transaction := _start_transaction()
            ; Perform post transaction/payment action based on the results
            _post_payment_action(transaction, false, true)
        } else {
            send, C
        }
    }
return

; no-svc-credit-debit
+n::
n::
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if is_main_window() {
        PARTIAL_CASH := false
        price := get_price()
        if (!price){
            send, N
            return
        }
        Gui, no_svc:Show, w677 h150, No SVC
    } else {
        send, N
    }
return

; Below are the Button controls for no svc
BtnCredit:
    Gui, no_svc: hide
    WinActivate , % CreWin
    handle_svc_keys(false, true, false)
return

BtnDebit:
    Gui, no_svc: hide
    WinActivate , % CreWin
    handle_svc_keys(false, false, false)
return

no_svcGuiEscape:
no_svcGuiClose:
BtnCancel:
    WinActivate , % CreWin
    Gui, no_svc: hide
return

; debit
+d::
d::
debit:
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt
    Gui, no_svc: hide

    if is_main_window() {
        PARTIAL_CASH := false
        price := get_price()
        if (!price){
            send, D
            return
        }
        if (price>= max_debit){
            handle_svc_keys(false, false, false)
        }else if ((min_debit) and (price <= min_debit)){
            splash_box("Use Credit Instead")
        } else {
            if (price >= svc2_debit_range_start and price < svc2_debit_range_end){
                handle_svc_keys(Format("{:.2f}", price*svc2_debit_percentage), false, false)
            } else if (price >= svc2_debit_range_end){
                handle_svc_keys("SVC3", false, false)
            } else {
                handle_svc_keys("SVC1", false, false)
            }
        }

    } else {
        if is_window_active(CASH_WIN){
            if (PARTIAL_CASH){
                price := get_price(true)
                Send, !n
                total_price := get_price()

                if (price >= svc2_debit_range_start and price < svc2_debit_range_end) {
                    svc_value := Format("{:.2f}", price*svc2_credit_percentage)
                } else if (price >= svc2_debit_range_end){
                    svc_value := "SVC3"
                } else {
                    svc_value := "SVC1"
                }

                if (svc_value){
                    if InStr(svc_value, "SVC"){
                        ; Send the SVC command
                        send, %svc_value%{enter}
                    } else {
                        send, SVC2{Enter}
                        send, %svc_value%{Enter}
                    }
                    PARTIAL_CASH := false
                }
                sleep, 300
                cash_paid := Format("{:.2f}", (total_price - price))
                send, !p
                sleep, 300
                send, %cash_paid%
                send, !c
            }
            Send, !r
            Send, !d
            ; Starts the transaction
            transaction := _start_transaction()
            ; Perform post transaction/payment action based on the results
            _post_payment_action(transaction, false, false)
        } else {
            send, D
        }
    }
return

; testing purpose
; +d::
; d::
; debit:
;     hide_splash()
;     WinHide, GrandTotal
;     WinHide, Custom Receipt
;     ; GuiControl, custom_receipt:, CashBack, Cashback Amount : $ %enable_recipt_on_credit_debit%
;     ; GuiControl, custom_receipt:, Signature, Do not forget to collect Signature
;     _post_payment_action(false, false, true, 0, is_ebt:=false)
; return

; split
+s::
s::
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if is_main_window() {
        PARTIAL_CASH := false
        Num := get_number()
        price_ := get_price()
        if Num {
            price := price_ - Format("{:.2f}", Num/100) ; Set global price to balance after cash
            if (price >= svc2_credit_range_start and price < svc2_credit_range_end) {
                handle_svc_keys(Format("{:.2f}", price*svc2_credit_percentage), true, Num)
            } else if (price >= svc2_credit_range_end){
                handle_svc_keys("SVC3", true, Num)
            } else {
                handle_svc_keys("SVC1", true, Num)
            }
            return
        } else {
            send, S
        }
    } else {
        send, S
    }
return

; POP PAY
+::
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if is_main_window() {
        price:= get_price()
        if (!price){
            return
        }
        win_class := get_window_class_id(Crewin)
        send, !p
        sleep, 500
        BTN_POP_INTF := POP_PAY_CTRL . "" . win_class . "_ad13"
        ControlClick, % BTN_POP_INTF, % Crewin
        sleep, 500
        BTN_POP := POP_PAY_CTRL . "" . win_class . "" . pop_pay_btn_suffix
        BTN_POP2 := POP_PAY_CTRL . "" . win_class . "" . online_pay_btn_suffix
        ControlClick, % BTN_POP, % Crewin
        sleep, 500
        ControlClick, % BTN_POP2, % Crewin
    }
    out_of_paper := check_Printer_Paper_Status()
    if(out_of_paper){
        splash_box("REPLACE THE PRINTER PAPER NOW TO AVOID FREEZING ISSUES.")
    }
return

; ONLINE PAY
-::
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if is_main_window() {
        price:= get_price()
        if (!price){
            return
        }
        win_class := get_window_class_id(Crewin)
        send, !p
        sleep, 500
        BTN_POP_INTF := POP_PAY_CTRL . "" . win_class . "_ad13"
        ControlClick, % BTN_POP_INTF, % Crewin
        sleep, 500
        BTN_ONLINE := POP_PAY_CTRL . "" . win_class . "" . online_pay_btn_suffix
        BTN_ONLINE2 := POP_PAY_CTRL . "" . win_class . "" . pop_pay_btn_suffix
        ControlClick, % BTN_ONLINE, % Crewin
        sleep, 500
        ControlClick, % BTN_ONLINE2, % Crewin
    }
    out_of_paper := check_Printer_Paper_Status()
    if(out_of_paper){
        splash_box("REPLACE THE PRINTER PAPER NOW TO AVOID FREEZING ISSUES.")
    }
return

; EBT
+e::
e::
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if is_main_window() {
        PARTIAL_CASH := false
        Num := get_number()
        price := get_price()
        if (!price){
            send, E
            return
        }
        if (ebt_svc1){
            handle_svc_keys("SVC1", false, false, true)
        } else {
            handle_svc_keys(false, false, false, true)
        }

    } else {
        if is_window_active(CASH_WIN){
            if(PARTIAL_CASH){
                price := get_price(true)
                Send, !n
                sleep, 300
                total_price := get_price()
                if(ebt_svc1){
                    send, SVC1{Enter}
                }
                PARTIAL_CASH := false
                sleep, 300
                cash_paid := Format("{:.2f}", (total_price - price))
                send, !p
                sleep, 300
                send, %cash_paid%
                send, !c
            }
            Send, !e
            Send, !e
            ; Starts the transaction
            transaction := _start_transaction(0,true)
            ; Perform post transaction/payment action based on the results
            _post_payment_action(transaction, false, false, 0, is_ebt:=true)
        } else {
            send, E
        }
    }
return

; cashback
+w::
w::
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if is_main_window() {
        PARTIAL_CASH := false
        Num := get_number()
        if Num {
            send,% "{BS " StrLen(Num) "}"
            send , e2{Enter}
            send , %Num% {enter}
        } else {
            send, W
        }
    } else {
        send, W
    }
return

^l::

    if WinExist("Credit Card Information") {
        return
    }
    WinGet, PID, PID, % Crewin
    Process, close, % PID

    CRE_PATH := "C:\Program Files (x86)\CRE.NET\CRE2004.exe"
    WORK_PATH := "C:\Program Files (x86)\CRE.NET"

    if FileExist(CRE_PATH){
        Run , % CRE_PATH, % WORK_PATH
    }
    Reload

return

Enter::
    send, {Enter}
    sleep, 100
    if WinExist(ITEM_NF_WIN){
        SoundBeep, 1000, 1000
    }
return

; Check special notes above
a::
b::
f::
g::
h::
i::
j::
k::
l::
m::
o::
p::
q::
r::
t::
u::
v::
    Critical
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    Num := get_number()
    PARTIAL_CASH := false
    if Num {
        prev_num := Num
        if A_ThisHotkey not contains %exempted_keys%
        {
            if limit_item {
                if (limit_item <= Num){

                    splash_box("Maximum limit has reached")
                    return

                }
            }
        }
        prev_hotkey := A_ThisHotkey
        execute_action(A_ThisHotkey, Num)
        Num =
    } else {
        if (prev_hotkey = A_ThisHotkey){
            execute_action(A_ThisHotkey, prev_num)
        } else {
            send, % Format("{:U}", A_ThisHotkey)
        }
    }
    Critical, Off
return
y::
z::
    Critical
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    Num := get_number()
    PARTIAL_CASH := false
    if Num {
        prev_num := Num
        prev_hotkey := A_ThisHotkey
        execute_action(A_ThisHotkey, "-" . Num)
        Num =
    } else {
        if (prev_hotkey = A_ThisHotkey){
            execute_action(A_ThisHotkey, "-" . prev_num)
        } else {
            send, % Format("{:U}", A_ThisHotkey)
        }
    }
    Critical, Off
return

; Return to main window (Closes all sub-windows)
Home::
mainScreen:
    hide_splash() ; Make sure there is no splash screen or overlay on screen
    WinHide, GrandTotal
    WinHide, Custom Receipt

    if WinExist("ProgressView") {
        return
    }
    ; loop and recursively close all windows
    ; until we are left with last one (i.e. main window)
    loop {
        if not is_main_window() {
            WinActivate , % CreWin
            WinClose , % CreWin
        } else {
            break
        }
    }
return

; Returns a number pressed on the control
/*
	Checks if a number is pressed inside the focused control (Usually the top most input)
	Params:- None
	Returns - Number or key pressed before action
*/

get_number(){
    global Num
    global CreWin
    ControlGetFocus, focused_ctrl
    ControlGetText , Num, % focused_ctrl, % CreWin
    if Num is number
        return Num
return false
}

; Key agnostic action method
/*
 Executes action based on the hotkey and opened window
 Ex - sends the keys directly to window if the action is executed over a dialog or other
 typebox
 Params :-
	hotkey - Hotkey to send to main screen
	Num - Number to be used
 Returns - Null
*/
execute_action(hotkey, Num){

    CreEditCtrl := "_ad11" ; Target edit control
    ControlGetFocus, focused_ctrl
    ControlGetFocus, focused_ctrl
    pos := Instr(focused_ctrl, CreEditCtrl)
    new_str := SubStr(focused_ctrl, pos)
    is_main_screen := is_main_window()
    ; Bypass the whole process in case of different editbox or
    ; if the command is executed over a/an dialog/overlay in CRE
    if ((StrLen(new_str) != StrLen(CreEditCtrl)) or not is_main_screen) {
        send, %hotkey%
        return
    }

    if not Num ; A fail safe measure to prevent accidents
        return

    hotkey := SubStr(hotkey,0) ; Sanitise hotkey to single key
    Send, % "{BS "StrLen(Num) "}"	; Clears the input
    SendRaw, %hotkey%	; Send the hotkey (Should be same as inventory item)
    Send, {enter}
    Send, % Num	; Send the number
    Send, !o	; Send OK

}

; Checks whether the keys we are trying to send are being sent to main window or not
/**
	This method checks whether there are multiple opened dialogs or windows in CRE
	via a straight forward Winget method
	Returns - True if there is only one, false otherwise
*/
is_main_window(){
    global Crewin
    WinGet , win_count, count, % Crewin
    if win_count = 1
        return true
return false
}

; Returns the class number/ID of given CRE window
/**
	This method parses the CRE window class and returns the ID
	Returns - ID string of the class
*/
get_window_class_id(win_title){
    WinGetClass, win_class, % win_title
    Class_Num := SubStr(win_class, 29, -4)
return Class_Num
}

; Checks whether the current window is specified sub-window or not
/**
	Returns - True if the current window is a sub-window, false otherwise
*/
is_window_active(win_name){
    global Crewin
    WinGetText, win_txt, % Crewin
    ; Check whether the first text in this window is "&cash" or not
    loop, parse, % win_txt,`r`n
    {
        if ((A_Index = 1) and (A_LoopField = win_name)){
            return True
        } else {
            break
        }
    }
return False
}

/**
	Retrieves total price of all current items present in the list
	params:-
			skip_key - Whether to skip pressing the cash dialog key (Useful when the cash window is already opened)
	returns - price in decimal form or false if there's an error
			Usually an error means that there are no items present
*/
get_price(skip_key:= false){
    runtime_err_win := "Run Time Support"
    if (not skip_key){
        send, !P
        Sleep, 300	; Give drawer some time to appear onto the screen
    }
    ; Check if it's a runtime error
    if WinExist(runtime_err_win){
        WinActivate , % runtime_err_win
        Send {enter}
        return false
    }
    ; Get focused control and retrieve price information
    ControlGetFocus, price_editbox, ahk_exe CRE2004.exe
    ControlGetText, price, %price_editbox%, ahk_exe CRE2004.exe

    ;Create a fail safe to prevent empty reading
    if price =
    {
        Clipboard =
        send, ^c
        ClipWait, 2
        price := Clipboard
    }
    ; Exit the screen
    send, !n
    ; remove any commas
    price := StrReplace(price, ",")
    ; Return price without currency sign
return LTrim(price, "$")
}

/**
	Send SVC keys to actual main interface based on the hotkeys pressed
	params:-
		svc_value - the command or SVC value to send. False to ignore
		is_credit - whether this is a credit or debit transaction
		split - Amount to split. False otherwise
			    This will also enable a split command i.e. partial cash and credit
	returns - null
*/
handle_svc_keys(svc_value, is_credit:= true, split := false, is_ebt := false){

    global LAST_INVOICE

    ; Send the actual keys
    if (svc_value){

        if InStr(svc_value, "SVC"){
            ; Send the SVC command
            send, %svc_value%{enter}
        } else {
            send, SVC2{Enter}
            send, %svc_value%{Enter}
        }
    }

    sleep, 500
    original_price := get_price()

    if (split){
        sleep, 500
        ; get the total_price
        total_price := get_price()
        ; open the main payment screen
        send, !p
        send, % split
    } else {
        send, !p
    }

    ; Press the credit/debit button
    if(not is_ebt){
        send, !r
    }

    ; Record last invoice number for cross-checking whether the current transaction went successful or not
    data := db_.execute("SELECT TOP(1) Invoice_Number FROM Invoice_Totals ORDER BY DateTime DESC")
    LAST_INVOICE := data[2,1]

    ; Check if this is a manual split
    if (split){
        sleep, 500
    } else {
        ; Select credit or debit accordingly
        if (is_credit) {
            send, !r
        } else if(is_ebt){
            send, !e
            send, !e
        } else {
            send, !d
        }
    }

    ; Starts the transaction
    transaction := _start_transaction(original_price, is_ebt)
    ; Perform post transaction/payment action based on the results
    _post_payment_action(transaction, svc_value, is_credit, original_price)
    SplashTextOff

}

; Private method (use internally only)
/**
	Performs the actual transaction and waits for the response from CC machine
	params - original_price - original price of current whole transaction
	returns - true if a transaction is successfull else false
*/
_start_transaction(original_price:=0, is_ebt := false){

    global interrupt, LAST_INVOICE, CURRENT_INVOICE, ERROR_MESSAGE
    waiting_win := "Credit Card Information"
    if(not is_ebt){
        WinWait , % waiting_win
    }
    SplashTextOn, , , Processing please wait...
    is_errored := false
    ERROR_MESSAGE := false
    ; loop and check whether the transaction is processed
    ; successfully via the machine or not
    loop {
        ; If we are at a main window. It means that either the transaction errored or went successful
        if is_main_window() {

            Sleep, 2000
            data := db_.execute("SELECT TOP(1) Invoice_Number FROM Invoice_Totals ORDER BY DateTime DESC")
            CURRENT_INVOICE := data[2,1]

            if ( CURRENT_INVOICE == LAST_INVOICE){
                is_errored := true
                ERROR_MESSAGE := "Failed Transaction/Timeout or Cancelled By User"
            }
            break
        }
        if winexist(waiting_win) {
            ; Usually an error leads to Payment processor error window
            if winexist("Payment Processor Error"){
                WinActivate, "Payment Processor Error"
                is_errored := true
                Clipboard =
                send, !c ; Copy the error message
                ClipWait, 2
                ERROR_MESSAGE := Clipboard
                SplashTextOff
                break
            }
        } else {
            ; if the window does not exist then it implies
            ; that either the whole process went smooth or there was a cancellation via user
            ; i.e. Transaction went successful or user pressed the cancel button
            Sleep, 2000
            ; Get the latest data from database
            data := db_.execute("SELECT TOP(1) Invoice_Number FROM Invoice_Totals ORDER BY DateTime DESC")
            CURRENT_INVOICE := data[2,1]

            ; Check if the current invoice is same as last invoice
            if ( CURRENT_INVOICE == LAST_INVOICE) {
                is_errored := true
                ERROR_MESSAGE := "Failed Transaction/Timeout or Cancelled By User"
            }
            break
        }
    }
    SplashTextOff

return is_errored
}

; Private method (use internally only)
/**
	Performs actions based on the transaction i.e. either rollback or continue
	params:-
		status - true/false. Where true means - An error is encountered
		is_credit - whether this is a credit or debit transaction
		original_price - Transaction Amount
	returns - null
*/
_post_payment_action(error_status, svc_value, is_credit:=true, original_price:=0, is_ebt:=false){

    global Crewin, msg_threshold, disable_print, CURRENT_INVOICE, LAST_INVOICE, CASHBACK_ITEM_NUM, ERROR_MESSAGE, print_receipt_x, print_receipt_y, print_receipt_width, print_receipt_height, enable_recipt_on_credit_debit
    if (error_status) {
        ; Print the receipt via thermal printer
        if (!disable_print){
            print_receipt(original_price, ERROR_MESSAGE)
        }

        ; Wait for the user to arrive at the main screen (Item list)
        loop {
            sleep, 500
            if is_main_window() {
                break
            }
        }

        if (svc_value){
            ; Once arrived. Check if the price is still valid (i.e. user didn't do full payment)
            price_check := get_price()
            if price_check
                send {del}	; if the price is still valid (i.e. >0) delete the automated SVC entry
        }

    } else {
        ; Verify if the whole process is successful (Optional)
        ; Since there should not be a price to check

        ; Check whether the current Invoice has Cashback in it
        data := db_.execute("SELECT ItemNum, origPricePer FROM Invoice_Itemized WHERE Invoice_Number='" . CURRENT_INVOICE . "' ")
        loop % data.MaxIndex()
        {
            ; Check whether there is a cashback item inside the last invoice
            if ((data[A_index,1] == "e2") or (data[A_index,1] == "E2")){
                ; Cashback item found. Open the drawer and show the cashback amount
                send, !o
                Sleep, 500
                send, 1b
                Sleep, 500
                send, !x
                amount := data[A_Index, 2]
                GuiControl, custom_receipt:, CashBack, Cashback Amount : $%amount%
                break
            }
            else{
                GuiControl, custom_receipt:, CashBack,
            }
        }

        ; Display signature if the original price exceeds the message threshold (in case of credit only)
        if (is_credit) {
            data := db_.execute("SELECT Grand_Total FROM Invoice_Totals WHERE Invoice_Number=" . CURRENT_INVOICE)
            d := data[2,1]
            GuiControl, custom_receipt:, Signature,
            if (data[2,1] > msg_threshold) {
                GuiControl, custom_receipt:, Signature, Do not forget to collect Signature
            }
            else{
                GuiControl, custom_receipt:, Signature,
            }
        }
        if (!is_ebt)
        {
            if(enable_recipt_on_credit_debit){
                Gui, custom_receipt:Show, x%print_receipt_x% y%print_receipt_y% w%print_receipt_width% h%print_receipt_height%, Custom Receipt
                sleep, 300
                WinActivate, %Crewin%
                Suspend, On
                Input, SingleKey,L1V, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause}
                Suspend, Off
                WinActivate, %Crewin%
                WinHide, Custom Receipt
                If InStr(ErrorLevel, "EndKey:")
                {
                    return
                }
                if (SingleKey = "x")
                {
                    Send, ^a{BackSpace}
                    Gosub, x
                }
                if (SingleKey = "e")
                {
                    Send, ^a{BackSpace}
                    Gosub, e
                }
                if (SingleKey = "c")
                {
                    Send, ^a{BackSpace}
                    Gosub, c
                }
                if (SingleKey = "d")
                {
                    Send, ^a{BackSpace}
                    Gosub, d
                }
                if (SingleKey = "w")
                {
                    Send, ^a{BackSpace}
                    Gosub, w
                }
                if (SingleKey = "s")
                {
                    Send, ^a{BackSpace}
                    Gosub, s
                }
                if (SingleKey = "n")
                {
                    Send, ^a{BackSpace}
                    Gosub, n
                }
            }
        }

    }
return
}

/**
    price update
*/
global first_entry := True
CheckItemExists:
    Gui, update_price:Submit
    if(!ItemNumber){
        return
    }
    query := "SELECT ItemNum, ItemName, Price, Tax_1, Tax_2, Tax_3, Tax_4, Tax_5, Tax_6, FoodStampable FROM Inventory WHERE ItemNum='" . ItemNumber . "'"
    item_data := db_.execute(query)
    item_num := item_data[2,1]
    item_name := item_data[2,2]
    price := item_data[2,3]
    item_tax1 := (item_data[2,4] != "" ? item_data[2,4] : 0)
    item_tax2 := (item_data[2,5] != "" ? item_data[2,5] : 0)
    item_tax3 := (item_data[2,6] != "" ? item_data[2,6] : 0)
    item_tax4 := (item_data[2,7] != "" ? item_data[2,7] : 0)
    item_tax5 := (item_data[2,8] != "" ? item_data[2,8] : 0)
    item_tax6 := (item_data[2,9] != "" ? item_data[2,9] : 0)
    food_stampable := (item_data[2,10] != "" ? item_data[2,10] : 0)

    price := format("{:0.2f}", price)

    tax_query := "SELECT Tax1_Rate, Tax2_Rate, Tax3_Rate, Tax4_Rate, Tax5_Rate, Tax6_Rate FROM Tax_Rate"
    tax_rates := db_.execute(tax_query)

    tax1_rate := tax_rates[2,1]
    tax2_rate := tax_rates[2,2]
    tax3_rate := tax_rates[2,3]
    tax4_rate := tax_rates[2,4]
    tax5_rate := tax_rates[2,5]
    tax6_rate := tax_rates[2,6]

    total_tax_rate := 0
    if (item_tax1)
        total_tax_rate += tax1_rate
    if (item_tax2)
        total_tax_rate += tax2_rate
    if (item_tax3)
        total_tax_rate += tax3_rate
    if (item_tax4)
        total_tax_rate += tax4_rate
    if (item_tax5)
        total_tax_rate += tax5_rate
    if (item_tax6)
        total_tax_rate += tax6_rate

    ; Calculate the price with tax
    price_with_tax := price * (1 + total_tax_rate)
    price_with_tax := format("{:0.2f}", price_with_tax)

    tag_along_query := "SELECT TagAlong_ItemNum FROM Inventory_TagAlongs WHERE ItemNum='" . ItemNumber . "'"
    inventory_tag_alongs := db_.execute(tag_along_query)

    tag_alongs := ""

    ; Iterate through the results and concatenate the TagAlong_ItemNum values
    Loop, % inventory_tag_alongs.MaxIndex() {
        if (A_Index = 1){
            Continue
        }
        if (tag_alongs != "") {
            tag_alongs .= "| " ; Add a separator if it's not the first item
        }
        tag_alongs .= inventory_tag_alongs[A_Index, 1]
    }
    if(item_num){
        GuiControl, price_change:, ItemNum, %item_num%
        GuiControl, price_change:, ItemName, %item_name%
        GuiControl, price_change:Focus, CPrice,
        GuiControl, price_change:, CPrice, $%price%
        GuiControl, price_change:, APrice, $%price_with_tax%
        GuiControl, price_change:, Tax, %item_tax1%
        GuiControl, price_change:, Tax2, %item_tax2%
        GuiControl, price_change:, Foodstampable, %food_stampable%
        GuiControl, price_change:, TagAlongItem, %tag_alongs%

        sleep, 300
        Send, ^a
        is_price_change := True
        Gui, price_change:Show, , Price Change
    } else{
        alert_message("Item Not Exist")
        GuiControl, update_price:, ItemNumber,
    }
return

PriceAdded:
    Gui, price_change:Submit,
    GuiControlGet, ItemNumber, ,ItemNum
    item_num := ItemNumber

    missingFields := ""
    if (CPrice = "")
        missingFields .= "Item Number"
    if (missingFields != "")
    {
        alert_message("Required Field: `n" . missingFields)
        return
    }
    CPrice := StrReplace(CPrice, "$", "")
    if (!InStr(CPrice, ".")) ; Check if CPrice does not contain a decimal point
    {
        CPrice := CPrice / 100 ; Divide by 100 only if there's no decimal point
    }
    sql_query := "UPDATE Inventory SET Price=" . CPrice . ", Tax_1=" . Tax . ", Tax_2=" . Tax2 . ", FoodStampable=" . Foodstampable . " WHERE ItemNum='" . item_num . "'"

    db_.execute(sql_query)

    delete_tag_along_sql_query := "DELETE FROM Inventory_TagAlongs WHERE ItemNum='" . item_num . "'"
    db_.execute(delete_tag_along_sql_query)

    Loop, Parse, tag_alongs, |
    {
        tag_along := Trim(A_LoopField) ; Trim any leading/trailing spaces
        insert_tag_along_sql_query := "INSERT INTO Inventory_TagAlongs (ItemNum, Store_ID, TagAlong_ItemNum, Quantity) VALUES ('" . item_num . "', '" . store_info[2,5] . "', '" . tag_along . "', 1 )"
        db_.execute(insert_tag_along_sql_query)
    }

    if(ErrorLevel)
    {
        GuiControl, update_price:, ItemNumber,
        is_price_change := False
        alert_message("Price not updated, Something went wrong")
    } else {
        tag_alongs := ""
        is_price_change := False
        GuiControl, price_change:, TagAlongItem, |
        GuiControl, update_price:, ItemNumber,
        send, {F5}
        send, %item_num%
        send, {Enter}
        send, !x
    }

    first_entry := True
return

Tax:
    GuiControlGet, Tax1, ,Tax
    item_tax1 := Tax1
    GuiControl, Focus, CPrice
    Gosub, PriceEdit
return
Tax2:
    GuiControlGet, Tax_2, ,Tax2
    item_tax2 := Tax_2
    GuiControl, Focus, CPrice
    Gosub, PriceEdit
return

PriceEdit:
    ControlGetFocus, focused_control, A
    if (focused_control != "Edit1")
    {
        return
    }
    GuiControlGet, retrive_price, ,CPrice
    retrive_price := StrReplace(retrive_price, "$", "")
    total_tax_rate := 0
    if (item_tax1)
        total_tax_rate += tax1_rate
    if (item_tax2)
        total_tax_rate += tax2_rate

    if(retrive_price = 0){
        price_tax = 0
    }
    else{
        price_tax := retrive_price * (1 + total_tax_rate)
        if(first_entry == false){
            if (!InStr(retrive_price,".")){
                price_tax := price_tax / 100
            }
        }
    }
    first_entry := False
    if(price_tax = ""){
        return
    }
    price_tax := Format("{:.2f}", price_tax)
    GuiControl, price_change:, APrice, $%price_tax%
return

TaxPriceEdit:
    ControlGetFocus, focused_control, A
    if (focused_control = "Edit1")
    {
        return
    }
    GuiControlGet, retrive_tax_price, , APrice
    retrive_tax_price := StrReplace(retrive_tax_price, "$", "")
    total_tax_rate := 0
    if (item_tax1)
        total_tax_rate += tax1_rate
    if (item_tax2)
        total_tax_rate += tax2_rate
    c_price := retrive_tax_price / (1 + total_tax_rate)
    if(first_entry == False){
        if (!InStr(retrive_tax_price,".")){
            c_price := c_price / 100
        }
    }
    first_entry := False
    c_price := Format("{:.2f}", c_price)
    GuiControl, price_change:, CPrice, $%c_price%
return

PriceCancel:
    GuiControl, update_price:, ItemNumber,
    GuiControl, price_change:, TagAlongItem, |
    first_entry := True
    is_price_change := False
    tag_alongs := ""
    WinHide, Price Change
return

price_changeGuiEscape:
price_changeGuiClose:
    GuiControl, update_price:, ItemNumber,
    GuiControl, price_change:, TagAlongItem, |
    first_entry := True
    is_price_change := False
    tag_alongs := ""
    winhide, Price Change
return

update_priceGuiEscape:
update_priceGuiClose:
    GuiControl, update_price:, ItemNumber,
    GuiControl, price_change:, TagAlongItem, |
    first_entry := True
    is_price_change := False
    tag_alongs := ""
    winhide, Price Change
    winhide, Update Price
return

alert_message(msg){
    MsgBox,262144,,%msg%, 1
}

/**
	Shows a splash screen with custom message on top which disappears on key press
	params:-
		message - Message to show on screen
		capture_key - True to capture input and false to ignore
	returns - null
*/
splash_box(message, capture_key:= true, is_post_payment:= false){
    global Crewin
    ; Use progress as splashtext (Check AHK documentation)
    Progress, zh0 fs18 w300, % message
    sleep, 300
    Input, Key, L1V, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause} ; wait for any keypress
    Progress, off ; Exit splash text
return
}

/*
	Hide a splash window
*/
hide_splash(){
    Progress, off
}

print_receipt(price, message){
    global PRINT_EXE, EXE_PATH
    ; Generate Error Message for receipt
    receipt_file := A_Temp . "\receipt.txt"
    FormatTime, TimeString,, Time
    FormatTime, DateString,, ShortDate
    FileDelete, % receipt_file
    text = Time: %TimeString%`r`nDate: %DateString%`r`nError Details: %message%
    final_text := create_printer_text(text, "$" . price)
    FileAppend, % final_text, % receipt_file
    ; Let's print it
    RunWait, % ComSpec . " /c """ . PRINT_EXE . " """ . receipt_file . """", % EXE_PATH, hide
}

/*
	Creates a printer friendly text via html with text as the body and amount as total
*/

create_printer_text(text, amount){
    global store_info, PRINT_WIN
return "RECEIPT`r`n************`r`n" . store_info[2,1] . "`r`n**************`r`n" . text . "`r`n**************`r`n" . amount . "`r`n**************`r`n"

}

check_Printer_Paper_Status(){
    global EXE_PATH, PRINT_EXE
    paper_arg := "countpaper"
    PAPER_STAT_EXE := PRINT_EXE . paper_arg
    OutputFile := A_Temp "\AutoPOS_Output.txt"
    RunWait, %ComSpec% /c %PAPER_STAT_EXE% > %OutputFile%, % EXE_PATH , hide, UseErrorLevel
    Sleep, 100
    FileRead, OutputVar, %OutputFile%

    if (OutputVar ~= "True") {
        return 1
    } else {
        return 0
    }

    ; Delete the temporary file
    FileDelete, %OutputFile%
}

/*
    custom receipt
*/

CashRegisterReceipt:
    WinActivate, %Crewin%

    GuiControl, +Background6A96C8, CashRegisterReceiptProgress
    Sleep, 100
    GuiControl, +Background4683b4, CashRegisterReceiptProgress
    Gui, Font, c000000
    GuiControl, Font, CashRegisterReceiptText

    custom_receipt("cash_register_receipt")
return

CardReceipt:
    WinActivate, %Crewin%

    GuiControl, +Background6A96C8, CardReceiptProgress
    Sleep, 100
    GuiControl, +Background4683b4, CardReceiptProgress
    Gui, Font, c000000
    GuiControl, Font, CardReceiptText

    custom_receipt("card_receipt")
return

custom_receipt(receipt_type){
    if is_main_window() {
        global PRINT_EXE, EXE_PATH, CARD_ENTRY_METHOD
        receipt_data := {}
        data := db_.execute("SELECT TOP(1) Invoice_Number, Grand_Total, Cashier_ID, Total_Tax1, Total_Tax2, Total_Tax3, DateTime, Total_Price, Station_ID, Payment_Method FROM Invoice_Totals WHERE Payment_Method IN('cc','dc') ORDER BY DateTime DESC")
        receipt_data["invoice_number"] := data[2,1]
        receipt_data["invoice_grand_total"] := data[2,2]
        receipt_data["invoice_cashier_id"] := data[2,3]
        tax_1 := data[2,4]
        tax_2 := data[2,5]
        tax_3 := data[2,6]
        receipt_data["date_time"] := data[2,7]
        receipt_data["total_price"] := data[2,8]
        receipt_data["tax"] := tax_1 + tax_2 + tax_3
        receipt_data["tax"] := Format("{:0.2f}", receipt_data["tax"])
        receipt_data["station_id"] := data[2,9]
        payment_method := data[2,10]
        if (payment_method == "cc"){
            receipt_data["payment_method"] := "CREDIT CARD PURCHASE"
        } else{
            receipt_data["payment_method"] := "DEBIT CARD PURCHASE"
        }
        employee_data := db_.execute("SELECT Cashier_ID, EmpName FROM Employee where Cashier_ID='" . receipt_data["invoice_cashier_id"] . "' ")
        receipt_data["employee_name"] := employee_data[2,2]

        transaction_data := db_.execute("SELECT TOP(1) Amount, appLabel, TruncatedCardNumber, TransType, Reference, CardEntrySource, emv_aid, tsi_Indicator, tc_acc FROM CC_Trans ORDER BY DateTime DESC")
        receipt_data["trans_amount"] := transaction_data[2,1]
        receipt_data["trans_app_label"] := transaction_data[2,2]
        receipt_data["trans_truncated_card_number"] := transaction_data[2,3]
        receipt_data["trans_trans_type"] := transaction_data[2,4]
        if(transaction_data[2,4] == "C1"){
            receipt_data["trans_trans_type"] := "PURCHASE"
        } else {
            receipt_data["trans_trans_type"] := "****"
        }
        receipt_data["trans_reference"] := transaction_data[2,5]
        receipt_data["trans_card_entry_source"] := CARD_ENTRY_METHOD[transaction_data[2,6]]
        receipt_data["trans_emv_aid"] := transaction_data[2,7]
        receipt_data["trans_tsi_indicator"] := transaction_data[2,8]
        receipt_data["trans_tc_acc"] := transaction_data[2,9]

        receipt_file := A_Temp . "\receipt.txt"
        FileDelete, % receipt_file

        item_data := db_.execute("SELECT Invoice_Number, ItemNum, Quantity, PricePer FROM Invoice_Itemized where Invoice_Number='" . receipt_data["invoice_number"] . "' ")
        receipt_data["item_list"] :=
        item_count := 0
        for index, value in item_data
        {
            if (index = 1) ; Skip the first row
                continue
            item_num := value[2]
            inventory := db_.execute("SELECT ItemName FROM Inventory where ItemNum='" . item_num . "' ")
            item_name := inventory[2,1]
            quantity := value[3]
            price_per := value[4]
            price_per := Format("{:.2f}", price_per)

            if(quantity > 1){
                item_description := quantity . "x " . item_name
            }
            else {
                item_description := quantity . " " . item_name
            }

            receipt_data["item_list"] := receipt_data["item_list"] . item_description . "~$" . price_per . "`r`n"
            item_count := item_count + 1
        }
        receipt_data["item_count"] := item_count
        receipt_data["dash"] := "========================================"
        if(receipt_type == "cash_register_receipt"){
            final_text := create_cash_register_receipt_text(receipt_data, receipt_type)
        } else {
            final_text := create_cash_register_receipt_text(receipt_data, receipt_type)
        }
        FileAppend, % final_text, % receipt_file
        RunWait, % ComSpec . " /c """ . PRINT_EXE . " """ . receipt_file . """", % EXE_PATH, hide

    }
return
}

/*
	Creates a printer friendly text via html with text as the body and amount as total
*/

create_cash_register_receipt_text(receipt_data, receipt_type){
    global store_info, PRINT_WIN
    ; use @ to center align and ~ to right align
    if(receipt_type == "cash_register_receipt"){
        return "@" . store_info[2,1] . "`r`n`r`n@" . store_info[2,2] . "`r`n@" . store_info[2,3] . "`r`n`@" . store_info[2,4] . "`r`n`r`nORDER# " . receipt_data["invoice_number"] . "`r`nINVOICE# " . receipt_data["invoice_number"] . "`r`nDATE/TIME: " . receipt_data["date_time"] . "`r`nCASHIER: " . receipt_data["employee_name"] . "`r`nSTATION: " . receipt_data["station_id"] . "`r`n`r`nItem Count: " . receipt_data["item_count"] . "`r`n" . receipt_data["dash"] . "`r`n" . receipt_data["item_list"] . "`r`n" . receipt_data["dash"] . "`r`nSubTotal:~$" . receipt_data["total_price"] . "`r`nTax:~$" . receipt_data["tax"] . "`r`nGrand Total:~$" . receipt_data["invoice_grand_total"] . "`r`n`r`nCredit:~$" . receipt_data["trans_amount"] . "`r`n`r`n" . receipt_data["payment_method"] . ": $" . receipt_data["trans_amount"] . "`r`nCard Type: " . receipt_data["trans_app_label"] . "`r`n" . receipt_data["trans_truncated_card_number"] . "`r`nTransaction Type: " . receipt_data["trans_trans_type"] . "`r`n`r`n~invoice:" . receipt_data["invoice_number"] . "@`r`n@`r`n@`r`n@`r`n@`r`n@.`r`n@.`r`n@.`r`n@.`r`n"
    } else {
        return "@" . store_info[2,1] . "`r`n`r`n@" . store_info[2,2] . "`r`n@" . store_info[2,3] . "`r`n@" . store_info[2,4] . "`r`n`r`nORDER# " . receipt_data["invoice_number"] . "`r`nINVOICE# " . receipt_data["invoice_number"] . "`r`nDATE/TIME: " . receipt_data["date_time"] . "`r`nCASHIER: " . receipt_data["employee_name"] . "`r`nSTATION: " . receipt_data["station_id"] . " `r`n`r`nItem Count: " . receipt_data["item_count"] . "`r`n" . receipt_data["payment_method"] . ": $" . receipt_data["trans_amount"] . "`r`nCard Type: " . receipt_data["trans_app_label"] . "`r`n" . receipt_data["trans_truncated_card_number"] . "`r`nTransaction Type: " . receipt_data["trans_trans_type"] . "`r`nRef Num: " . receipt_data["trans_reference"] . "`r`nAuth Code: " . receipt_data["trans_reference"] . "`r`nApp Label: " . receipt_data["trans_app_label"] . "`r`nAID: " . receipt_data["trans_emv_aid"] . "`r`nTSI: " . receipt_data["trans_tsi_indicator"] . "`r`nTC ACC: " . receipt_data["trans_tc_acc"] . "`r`n`r`n~invoice:" . receipt_data["invoice_number"] . "@`r`n@`r`n@`r`n@`r`n@`r`n@.`r`n@.`r`n@.`r`n@.`r`n"
    }
}

; Run as admin
run_as_admin(){
    ; Snippet taken from AHK documentation
    full_command_line := DllCall("GetCommandLine", "str")

    if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
    {
        try
        {
            if A_IsCompiled
                Run *RunAs "%A_ScriptFullPath%" /restart
            else
                Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
        }
        ExitApp
    }

}

Class DB {

    init(user:="sa", pass:="pcAmer1ca", db:="cresql", server:= "PCAMERICA")
    {
        if not InStr(server, "\"){
            server := A_ComputerName . "\" . server
        }
        this.conn := "Provider=SQLOLEDB.1;Password=" . pass . ";Persist Security Info=True;User ID= " . user . ";Initial Catalog=" . db . ";Data Source=" . server . ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    }

    execute(query) {
        return ADOSQL(this.conn, query)
    }

    check_db(){
        data:= ADOSQL(this.conn, "select Company_Info_1, Company_Info_2, Company_Info_3, Company_Info_4, Store_ID from Setup")
        if (ADOSQL_LastError)
            return false
        return 	data
    }

}

/*
###############################################################################################################
######                                      ADOSQL v5.04L - By [VxE]                                     ######
###############################################################################################################

	Wraps the utility of ADODB to connect to a database, submit a query, and read the resulting recordset.
	Returns the result as a new object (or array of objects, if the query has multiple statements).
	To instead have this function return a string, include a delimiter option in the connection string.

	For AHK-L (v1.1 or later).
	Freely available @ http://www.autohotkey.com/community/viewtopic.php?p=558323#p558323

	IMPORTANT! Before you can use this library, you must have access to a database AND know the connection
	string to connect to your database.

	Varieties of databases will have different connection string formats, and different drivers (providers).
	Use the mighty internet to discover the connection string format and driver for your type of database.

	crr connection string for SQLServer (2005) listening on port 1234 and with a static IP:
	DRIVER={SQL SERVER};SERVER=192.168.0.12,1234\SQLEXPRESS;DATABASE=mydb;UID=admin;PWD=12345;APP=AHK
*/
ADOSQL( Connection_String, Query_Statement ) {
    ; Uses an ADODB object to connect to a database, submit a query and read the resulting recordset.
    ; By default, this function returns an object. If the query generates exactly one result set, the object is
    ; a 2-dimensional array containing that result (the first row contains the column names). Otherwise, the
    ; returned object is an array of all the results. To instead have this function return a string, append either
    ; ";RowDelim=`n" or ";ColDelim=`t" to the connection string (substitute your preferences for "`n" and "`t").
    ; If there is more than one table in the output string, they are separated by 3 consecutive row-delimiters.
    ; ErrorLevel is set to "Error" if ADODB is not available, or the COM error code if a COM error is encountered.
    ; Otherwise ErrorLevel is set to zero.

    coer := "", txtout := 0, rd := "`n", cd := "CSV", str := Connection_String ; 'str' is shorter.

    ; Examine the connection string for output formatting options.
    If ( 9 < oTbl := 9 + InStr( ";" str, ";RowDelim=" ) )
    {
        rd := SubStr( str, oTbl, 0 - oTbl + oRow := InStr( str ";", ";", 0, oTbl ) )
        str := SubStr( str, 1, oTbl - 11 ) SubStr( str, oRow )
        txtout := 1
    }
    If ( 9 < oTbl := 9 + InStr( ";" str, ";ColDelim=" ) )
    {
        cd := SubStr( str, oTbl, 0 - oTbl + oRow := InStr( str ";", ";", 0, oTbl ) )
        str := SubStr( str, 1, oTbl - 11 ) SubStr( str, oRow )
        txtout := 1
    }

    ComObjError( 0 ) ; We'll manage COM errors manually.

    ; Create a connection object. > http://www.w3schools.com/ado/ado_ref_connection.asp
    ; If something goes wrong here, return blank and set the error message.
    If !( oCon := ComObjCreate( "ADODB.Connection" ) )
        Return "", ComObjError( 1 ), ErrorLevel := "Error"
    , ADOSQL_LastError := "Fatal Error: ADODB is not available."

    oCon.ConnectionTimeout := 3 ; Allow 3 seconds to connect to the server.
    oCon.CursorLocation := 3 ; Use a client-side cursor server.
    oCon.CommandTimeout := 900 ; A generous 15 minute timeout on the actual SQL statement.
    oCon.Open( str ) ; open the connection.

    ; Execute the query statement and get the recordset. > http://www.w3schools.com/ado/ado_ref_recordset.asp
    If !( coer := A_LastError )
        oRec := oCon.execute( ADOSQL_LastQuery := Query_Statement )

    If !( coer := A_LastError ) ; The query executed OK, so examine the recordsets.
    {
        o3DA := [] ; This is a 3-dimensional array.
        While IsObject( oRec )
            If !oRec.State ; Recordset.State is zero if the recordset is closed, so we skip it.
            oRec := oRec.NextRecordset()
        Else ; A row-returning operation returns an open recordset
        {
            oFld := oRec.Fields
            o3DA.Insert( oTbl := [] )
            oTbl.Insert( oRow := [] )

            Loop % cols := oFld.Count ; Put the column names in the first row.
                oRow[ A_Index ] := oFld.Item( A_Index - 1 ).Name

            While !oRec.EOF ; While the record pointer is not at the end of the recordset...
            {
                oTbl.Insert( oRow := [] )
                oRow.SetCapacity( cols ) ; Might improve performance on huge tables??
                Loop % cols
                    oRow[ A_Index ] := oFld.Item( A_Index - 1 ).Value
                oRec.MoveNext() ; move the record pointer to the next row of values
            }

            oRec := oRec.NextRecordset() ; Get the next recordset.
        }

        If (txtout) ; If the user wants plaintext output, copy the results into a string
        {
            Query_Statement := "x"
            Loop % o3DA.MaxIndex()
            {
                Query_Statement .= rd rd
                oTbl := o3DA[ A_Index ]
                Loop % oTbl.MaxIndex()
                {
                    oRow := oTbl[ A_Index ]
                    Loop % oRow.MaxIndex()
                        If ( cd = "CSV" )
                    {
                        str := oRow[ A_Index ]
                        StringReplace, str, str, ", "", A
                        If !ErrorLevel || InStr( str, "," ) || InStr( str, rd )
                            str := """" str """"
                        Query_Statement .= ( A_Index = 1 ? rd : "," ) str
                    }
                    Else
                        Query_Statement .= ( A_Index = 1 ? rd : cd ) oRow[ A_Index ]
                }
            }
            Query_Statement := SubStr( Query_Statement, 2 + 3 * StrLen( rd ) )
        }
    }
    Else ; Oh NOES!! Put a description of each error in 'ADOSQL_LastError'.
    {
        oErr := oCon.Errors ;  http://www.w3schools.com/ado/ado_ref_error.asp
        Query_Statement := "x"
        Loop % oErr.Count
        {
            oFld := oErr.Item( A_Index - 1 )
            str := oFld.Description
            Query_Statement .= "`n`n" SubStr( str, 1 + InStr( str, "]", 0, 2 + InStr( str, "][", 0, 0 ) ) )
            . "`n Number: " oFld.Number
            . ", NativeError: " oFld.NativeError
            . ", Source: " oFld.Source
            . ", SQLState: " oFld.SQLState
        }
        ADOSQL_LastError := SubStr( Query_Statement, 4 )
        Query_Statement := ""
        txtout := 1
    }

    ; Close the connection and return the result. Local objects are cleaned up as the function returns.
    oCon.Close()
    ComObjError( 1 )
    ErrorLevel := coer
Return txtout ? Query_Statement : o3DA.MaxIndex() = 1 ? o3DA[1] : o3DA
} ; END - ADOSQL( Connection_String, Query_Statement )
