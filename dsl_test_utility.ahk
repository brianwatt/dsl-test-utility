/*
DSL Test Utility - tool to automate account information retrieval and DSL modem testing. Written in AutoHotkey(AHK)

This script has been scrubbed to remove the actual URLs and tool names used at the company that I worked for.

Author:  Brian Watt - brianwatt at gmail dot com
Many thanks to Joe Glines for providing resources - https://the-automator.com/

Requriements:
Minimum AutoHotkey version: 1.1.30.01 
   https://www.autohotkey.com/download/
Selenium Basic version 2.09.0 
   https://github.com/florentbr/SeleniumBasic/releases/download/v2.0.9.0/SeleniumBasic-2.0.9.0.exe
Internet Explorer (due to the ease of using COM)

Help:
To understand a large part of what this script does and how to set it up see:
   https://the-automator.com/cross-browser-web-scraping-with-autohotkey-and-selenium/
For general AutoHotkey help:
   https://www.autohotkey.com/docs/AutoHotkey.htm

Reommended:
Editor for AHK scripts
   https://github.com/fincs/SciTE4AutoHotkey
Using this style/theme:
   https://github.com/ahkon/Darchon

Any part of this script containing the WinMove command needs to be modified to meet your display setup arrangement
Currently, this is only in the WinMoveMultipleAccountsMsgBox label
To easily find the coordinates you need use the built in Window Spy tool included with AHK

Feel free to contact me for help, contributions, bug reporting. 

This is my first crack at any kind of programming past 'hello world' - I realize the code is probably a mess.
I just first needed it to work. 

The DSL support contract that I was on ended before I could get my team to use this software, but it served me well in the few weeks I was able to use it. 

I left some code sections in that are commented out because I left the contract so quickly that I couldn't 
remember why they were commented out.  If this script is still usable/useful - those sections can 
likely be removed.  

*/

#SingleInstance, Force
; ################### GUI ###################################

; BTN
gui, add, edit, x10 w85 vBTN, 
; Click on the text 'BTN' to remove dashes in the phone number
gui, add, text, x+10 gdashremove, BTN
; Search DSL Provisioning System with the number in the BTN field
gui, add, button, x+10 w70 gdsl_provisioning_search, DSL DSL Provisioning System
; Search Fiber Provisioning System with the number in the BTN field
gui, add, button, x+10 gfiber_provisioning, Fiber Provisioning System

; WTN
gui, add, edit, x10 w85 vWTN,
; Auto-populated with WTN from DSL DSL Provisioning System
gui, add, text, x+10, WTN
; Run DSL Test If bonded line - BONDED_TEST will run upon DSL_TEST fail
gui, add, button, x+10 gdsltest, DSL_TEST ; h25 w70
; Run DSLTest BONDED_TEST test for bonded DSL
gui, add, button, x+10 gdsltestbonded, BONDED_TEST ; h25 w70
gui, add, text, x+10, DSLTest Tests

; Account Owner
gui, add, edit, x10 w200 vAcctOwner,
; Auto-populated by DSL Provisioning System data
gui, add, text, x+10, Account Owner

; Talked To
gui, add, edit, x10 w200 vTalkedTo, 
; The caller's name
gui, add, text, x+10, Talked To
; Click button if also speaking with account owner
gui, add, button, x+10 w70 gcopyacctowner, =Acct Owner

; CBN (call back number or can be reached number)
gui, add, edit, x10 w85 vCBN,
; Enter the CBN without symbols/spaces - i.e. 5558675309
gui, add, text, x+10, CBN
; Button to copy BTN to CBN if same
gui, add, button, x+10 gcbnequalsbtn, =BTN
; Launch dialer and automatically call customer back 
gui, add, button, x+10 gcallback, Call Back

; Verification
gui, add, edit, x10 w200 vVerification,
; Enter verification method used in Billing System such as 'Verified by PIN' etc.
gui, add, text, x+10, Verification Method

; Service Address auto-populated from DSL Provisioning System
gui, add, edit, x10 h40 w200 vServiceAddress
gui, add, text, x+10, Service Address

; Two letter state name
gui, add, edit, x+10 w30 vState,
; Clicking the text 'State' copies the 2 letter state ID to the clipboard
gui, add, text,x+10 gcopystate, State

; Account Status auto-populated from DSL Provisioning System 
gui, add, edit, cRed x10 w85 vAcctStatus,
gui, add, text, x+10, Acct Status   

; Service Type auto-populated from DSL Provisioning System 
gui, add, edit, x10 w300 vServiceType,
gui, add, text, x+10, Service Type

; Free form field for notes during the call
gui, add, edit, x10 w400 r7 vIssue, 

; Trouble Ticket info
; Manually enter the trouble ticket number
gui, add, edit, x10 w65 vTicketNumber
; Click text to remove the first three characters from ticket if copied from Ticket System - i.e. CX2007295109 resulting in 007295109
gui, add, text,x+10 gcleanticket, Trouble Ticket Number
; Lookup ticket by number in the WTN field
gui, add, button, x+10 gtsphone, Ticket System WTN
; Lookup ticket by number in the 'Troube Ticket Number' field
gui, add, button, x+10 gtsticket, Ticket System Ticket
; Go to Ticket System page
gui, add, button, x+10 gtsfield, Ticket System
; Ticket System Call ID - can be manually entered or auto-populated if 'Ticket System Ticket' button pressed 
gui, add, edit, x10 w65 vTicketSystemCall
gui, add, text,x+10, Ticket System Call Number

; CLLI NAS Information
; CLLI auto-popultated from DSL Provisioning System data
gui, add, edit, x10 w100 vCLLI,
gui, add, text,x+10, CLLI
; NAS ID from DSL Provisioning System (have to add manually for now - i.e. 'kank')
gui, add, edit, x+10 w30 vNAS,
gui, add, text,x+10, NAS
; USI auto-popultated from DSL Provisioning System 
gui, add, edit, x+10 w90 vUSI,
gui, add, text, x+10, USI

; DSLTest VER Code auto-populated 
gui, add, edit, x10 w50 vDSLVER,
; Click the text to copy the VER Code
gui, add, text,x+10 gcopyvercode, VER Code
; DSLTest VER Error auto-populated
gui, add, edit, x+10 h40 w200 vDSLVERError,
; Click the text to copy the concatenated VER Code and VER Error
gui, add, text,x+10 gConcatCodeError, VER Error

; Click to generate notes to send to an L2 or external team
gui, add, button, x10 h25 w70 gassistnotes, Assist Notes
; Click to generate notes from the free-form field.  Eliminates redundant information for customer notes. 
gui, add, button, x+10 h25 w70 gbilling_system_notes, Billing System Notes

; Bring individual IE windows forward to loosley emulate tabbed browsing
gui, add, button, x10 gactivatedsl_provisioning, DSL Provisioning System
gui, add, button, x+10 gactivatedsltest, DSLTest
gui, add, button, x+10 gactivatetsfield, Ticket System
gui, add, button, x+10 gactivatefiber_provisioning, Fiber Provisioning System
gui, add, button, x+10 gactivatelegacy_billing_system, Billing System
; Refresh (Send F5) DSL Provisioning System, DSLTest, Ticket System, Fiber Provisioning System, Dialer, and Billing System
gui, add, button, x+55 h30 w70 grefreshtools, Refresh Tools
; Print the script version at the bottom of the window
gui, add, text, x10, %A_ScriptName%


;################################ GUI Options ###########################
; Display the GUI, autosized with 'DSL Test Utility' as the app name
gui, Show, AutoSize,DSL Test Utility
; Set default field as BTN
GuiControl, Focus, BTN
; Keep app on top of other windows
gui, +AlwaysOnTop
; Sets edit and app background color
gui, Color,d1d1d1, F2F2F2
; Refresh DSL Provisioning System, DSLTest, Ticket System, Fiber Provisioning System, Dialer, and Billing System when app starts
; gosub, refreshtools
return

;################################ Labels ###########################

; Behavior for closing the app with windows close control
GuiClose:
ExitApp
return

; Remove dashes from BTN copied from Billing System
dashremove:
gui, submit, NoHide
BTN := StrReplace(BTN, "-", "")
GuiControl,, BTN,%BTN%
return

; Copy account owner to 'talked to' field
copyacctowner:
gui, submit, NoHide
GuiControl,,talkedto,%acctowner%
return

; Copy BTN to CBN field
cbnequalsbtn:
gui, submit, NoHide
GuiControl,,CBN,%BTN%
return

; Remove first three characters from full trouble ticket number (i.e. CX2)
cleanticket:
gui, submit, NoHide
if RegExMatch(TicketNumber, "\A\SO\/TT")
{
	legacy_billing_system_ticket := Clipboard
	RegExMatch(legacy_billing_system_ticket, "0\d{8}", TicketNumber)
	GuiControl,,TicketNumber, %TicketNumber%
	return
}
else
{
	StringTrimLeft, TicketNumber, TicketNumber, 3
	GuiControl,, TicketNumber,%TicketNumber%
}

; Copy VER Code to clipboard
copyvercode:
gui, submit, NoHide
Clipboard =
Clipboard := DSLVER
return

; Copy concatenated VER Code and VER Error to clipboard
ConcatCodeError:
gui, submit, NoHide
Clipboard =
Clipboard := DSLVER " - " DSLVERError
return

; Get just the two letter state abbreviation from the Service Address
getstate:
    ; Split PsServiceAddress into an array
PsState := StrSplit(PsServiceAddress, ",")
    ; Get the second element in the array
PsState := PsState[2]
    ; Remove trailing spaces
StringTrimRight, PsState, PsState, 7
    ; Remove leading spaces
StringTrimLeft, PsState, PsState, 1
GuiControl,, State, %PsState%
return

; Copy the 2 letter state code to the clipboard
copystate:
gui, submit, NoHide
Clipboard =
Clipboard := State
return

; Activate each IE window and press F5
refreshtools:
gui, submit, NoHide
WinActivate, Billing System
Send {F5}
sleep, 500

WinActivate, Dialer ahk_class IEFrame
Send {F5}
sleep, 500

WinActivate, Provisioning System ahk_class IEFrame
Send {F5}
sleep, 500

WinActivate, Fiber Provisioning System ahk_class IEFrame
Send {F5}
sleep, 500

WinActivate, DSLTest ahk_class IEFrame
Send {F5}
gosub, Wait

sleep, 5000

gosub, ts_window_titles
Send {F5}
return

; Pop up a message box that shows the copied contens and copies the note text to the clipboard formatted for L2 or other team
assistnotes:
gui, submit, NoHide

assistnotes := "BTN - " . BTN . "`n" . "WTN - " . WTN . "`n" . "Account Owner - " . AcctOwner . "`n" . "TT - " . TalkedTo . "`n" . "CBN - " . CBN . "`n" . verification . "`n" . "`n" . "State - " . State . "`n" . "Service - " .  ServiceType . "`n" . "`n" . "Issue - " . issue

Clipboard =
sleep, 500
Clipboard:=assistnotes

; pop up the notes that will be sent to the L2
MsgBox,4096, Notes, % assistnotes
return

; Copy free-form field for Billing System notes - removes redundant customer information
billing_system_notes:
gui, submit, NoHide

billing_system_notes := issue

Clipboard =
sleep, 500
Clipboard:=billing_system_notes

;pop-up the notes that will appear in Billing System
MsgBox,4096, Notes, % billing_system_notes
return

; Get information from DSL Provisioning System
dsl_provisioning_search:
gui, submit, NoHide

WinActivate, ProvisioningSystem
pwb := WBGet()

pwb.Navigate("http://dslprovisioning.company.com/dslprovisioning/showform")

Gosub, Wait

   ; Search for BTN
pwb.document.GetElementsByName("btnWtnCtn")[0].Value :=BTN
pwb.document.GetElementsByName("btnWtnCtnSubmit")[0].click()

Gosub, Wait

   ; Check for BTN tied to multiple accounts
PsMultipleAccts:=pwb.document.GetElementsByTagName("H").item[0].innertext
if (PsMultipleAccts == "Account Search Results" ){
      SetTimer, WinMoveMultipleAccountsMsgBox, 50
      MsgBox,4096,Multiple Accounts, Multiple Accounts
   }
   
   ; Grab Account Status
   ps_status:=pwb.document.GetElementsByTagName("INPUT").item[16].value
   ; MsgBox,, Acct Status, % ps_status
   
   ; Color text for Active accounts Green
   GuiControl,,acctstatus,%ps_status%
   if (InStr(ps_status, "Active"))
      {
         guicontrol, +cGreen, acctstatus
         ;GuiControl, MoveDraw, acctstatus
      }
   
   ; Grab Service Type (no innerhtml)
   PsServiceTypeOuter:=pwb.document.GetElementsByName("svcType").item[0].OuterHTML
   GrabService := StrSplit(PsServiceTypeOuter, """")
   PsServiceType := Grabservice[6]
   GuiControl,,ServiceType, %PsServiceType%
   
   ; Color gui text red if account status is "Frontier Mail E-Mail User"
      if (InStr(PsServiceType, "Frontier Mail E-Mail User"))
      {
         guicontrol, +cRed, ServiceType
         ;GuiControl, MoveDraw, ServiceType
      }
   
   ; Grab the WTN USI
   PsUSI := pwb.document.GetElementsByName("wtnUsi")[0].Value
   GuiControl,,USI,%PsUSI%
   
   ; Grab BTN/WTN
   PsBTN:=pwb.document.GetElementsByName("btn").item[0].Value 
   PsWTN:=pwb.document.GetElementsByName("wtn").item[0].Value 

   ; Grab First name + Last name
   PsFirstName:=pwb.document.GetElementsByName("firstName").item[0].Value
   PsLastName:=pwb.document.GetElementsByName("lastName").item[0].Value
   
   ; Update WTN in GUI
   GuiControl,,WTN,%PsWTN%

   ; Update Account Owner in GUI
   GuiControl,,AcctOwner,%PsFirstName% %PsLastName%
   
   ; Click on Plant tab and get CLLI
   pwb.document.GetElementsByTagName("A").item[6].click()
   Gosub, Wait
   dslamclli:=pwb.document.GetElementsByTagName("TD").item[46].innertext
   
   ; Update GUI with CLLI from Plant tab
   GuiControl,,CLLI,%dslamclli%
   
   ; Copy Service Address Table from Plant tab
   PsServiceInner := pwb.document.getElementsByTagName("TABLE")[6].innertext
   ; Turn Table innertext into an array
   PsServiceInner := StrSplit(PsServiceInner, ":")
      
   ; Street picks up the text "City, State Zip" - needs to be removed
   PsStreet := PsServiceInner[4]
   ; Remove "City, State Zip" 
   PsStreet := StrReplace(PsStreet, "City, State Zip", "")
   ; Remove trailing spaces and line break
   StringTrimRight, PsStreet, PsStreet, 3
   ; Copy City, State Zip info
   PsCityStateZip := PsServiceInner[5]
   
   ; Concatenate street and city, state, zip
   PsServiceAddress := PsStreet "`n" PsCityStateZip
   
   ; Update Service Address in GUI
   GuiControl,,ServiceAddress,%PsServiceAddress%
   
   ; Update State field in GUI
   gosub, getstate
   
   ; Navigate to the 'Radius Usage' page
   ; Click on Tools Tab
   pwb.document.GetElementsByTagName("A").item[7].click()
   
   Gosub, Wait

   ; Select 'Radius Usage' from dropdown
   pwb.document.GetElementsByTagName("SELECT").item[2].selectedIndex :=3 
   ; Click the 'Go' Button next to dropdown list
   pwb.document.GetElementsByName("goBtn")[0].click()
   

return

; Code from here: https://autohotkey.com/board/topic/9950-how-do-i-define-the-position-of-a-msgbox-on-screen/

; Code used to move the app Message Boxes to user defined area of the screen
WinMoveMultipleAccountsMsgBox:
   SetTimer, WinMoveMultipleAccountsMsgBox, OFF
   ID:=WinExist("Multiple Accounts")
   ; Customize the x and y coordinates to your needs
   WinMove, ahk_id %ID%, , -334 , 730 
return

; Get information from DSLTest
dsltest:
gui, submit, NoHide

    ; Clear GUI VER Code and Error
    GuiControl,,DSLVER, 
    GuiControl,,DSLVERError,

	; Activate DSLTest window
	WinActivate, DSLTest
	
	; Call the WBGet function and go to dsltest page
	pwb := WBGet()
	pwb.Navigate("http://dsltest.company.com/dsltest/CircuitTest.jsp")

	Gosub, Wait

	; Put WTN in Telephone field
	pwb.document.GetElementsByName("tn")[0].Value :=WTN
    ; Click Submit
	pwb.document.GetElementsByTagName("IMG").item[36].click()
    
    Gosub, Wait
    ;sleep, 10000
    
    gosub, dsltestwait

	; Get VER Code and VER Error in subroutine
	gosub, dsltestresults
	;Gosub, Wait

	; Catch DSL_TEST test run on bonded line and automatically run BONDED_TEST
	if (InStr(vererror, "USE BONDED_TEST"))
	{    
		; Present message box informing of bonded line
		SetTimer, WinMoveDSL_TESTMsgBox, 50
		MsgBox,4096,Run DSL_TEST, Bonded Line - Click OK to run BONDED_TEST

		; Clear GUI VER Code and Error
		; GuiControl,,DSLVER, 
		; GuiControl,,DSLVERError, 
        
        gosub, dsltestbonded
     }
/*
		; Select 'BONDED_TEST' from the 'Test Request' menu
		pwb.document.GetElementsByName("testrequest")[0].selectedIndex :=6

		pwb.document.GetElementsByName("tn")[0].Value :=WTN
		pwb.document.GetElementsByTagName("IMG").item[42].click()
        
       

		Gosub, Wait
        gosub, dsltestwait
         
         
         
		gosub, dsltestresults
}      

pwb.document.GetElementsByTagName("INPUT").item[3].click()

	; td 1 td 4 in print view
*/
return

dsltestbonded:
gui, submit, NoHide

    ; Clear GUI VER Code and Error
    GuiControl,,DSLVER, 
    GuiControl,,DSLVERError,

	; Activate DSLTest window
	WinActivate, DSLTest

	; Call the WBGet function and go to dsltest page
	pwb := WBGet()
	pwb.Navigate("http://lpcare.corp.pvt/dsltest/CircuitTest.jsp")

	Gosub, Wait

	; Put WTN in Telephone field
	pwb.document.GetElementsByName("tn")[0].Value :=WTN
	
	; Select 'BONDED_TEST' from the 'Test Request' menu
	pwb.document.GetElementsByName("testrequest")[0].selectedIndex :=6
	
	; Click 'Submit'
	pwb.document.GetElementsByTagName("IMG").item[36].click()
    
    Gosub, Wait
    gosub, dsltestwait

	; Get VER Code and VER Error in subroutine
	gosub, dsltestresults
	Gosub, Wait

	;pwb.document.GetElementsByName("tn")[0].Value :=WTN
	;pwb.document.GetElementsByTagName("IMG").item[42].click()

	Gosub, Wait

	;gosub, dsltestresults   

	pwb.document.GetElementsByTagName("INPUT").item[3].click()

		; td 1 td 4 in print view

return


dsltestwait:
   ; Wait until the test stops
   while (pwb.document.getElementsByTagName("TD").item[16].innertext != " Complete") 
   {
      sleep, 50
   }
return

dsltestresults:

	; Grab VER Code and Error
	vercode := pwb.document.getElementsByTagName("TD")[18].innertext
    vererror := pwb.document.getElementsByTagName("TD")[50].innertext

    ; Update GUI with VER Code and Error
	GuiControl,,DSLVER, %vercode%
	GuiControl,,DSLVERError, %vererror%

	; Green text in VER Code if Online, red for Offline, otherwise orange. 
	if (InStr(vercode, "Online"))
		{    
		 guicontrol, +cGreen, DSLVER
		 }
	else if (InStr(vercode, "Offline"))
         {
         ; 
         guicontrol, +cRed, DSLVER
         }
    else
		{
         ; Orange for all other VER codes
		 guicontrol, +cFF8000, DSLVER
		}  
return

WinMoveDSL_TESTMsgBox:
   SetTimer, WinMoveDSL_TESTMsgBox, OFF
   ID:=WinExist("Run DSL_TEST")
   ; Customize the x and y coordinates to your needs
   WinMove, ahk_id %ID%, , -334 , 730 
return

; Search Fiber Provisioning System with BTN
fiber_provisioning:
   gui, submit, NoHide
   WinActivate, Fiber Provisioning System ahk_class IEFrame
   pwb := WBGet()
   pwb.Navigate("http://fiberprovisioning.corp.company.com/search/subscriber")
   
   Gosub, Wait
   
   pwb.document.GetElementsByName("phoneNumber")[0].Select()
   Clipboard =
   Clipboard := BTN
   ; This needs to be fixed - don't want to use keystrokes if possible
   Send ^v
   sleep, 5000
   pwb.document.GetElementsByTagName("BUTTON").item[1].click()
   
return

; Go to Ticket System page
ticketsystem:
   gui, submit, NoHide
   
   gosub, ts_window_titles
   
   pwb := WBGet()
   pwb.Navigate("https://ticketsystem.company.com)
 

return

; Search Ticket System with WTN
tsphone:
   gui, submit, NoHide
   
   gosub, ts_window_titles
   
   ; Phone search link - 
   pwb := WBGet()
   pwb.Navigate("https://ticketsystem.company.com
   
   Gosub, Wait

   ; Search for WTN
   
   pwb.document.GetElementsByName("filtValTxt_0_admin6541fser65")[0].focus()
   Clipboard =
   Clipboard := WTN
   
   ; This needs to be fixed - don't want to use keystrokes
   Send ^v
   
   sleep, 1000
   pwb.document.getElementByID("footerButton_Search").click()

   Gosub, Wait
return

; Search Ticket System with ticket number
tsticket:
   gui, submit, NoHide
   
   gosub, ts_window_titles
   
   pwb := WBGet()
   pwb.Navigate("https://ticketsystem.company.com")
   
   Gosub, Wait
   
   tscallnumber:=pwb.document.GetElementsByTagName("U").item[1].innertext
   GuiControl,,TicketSystemCall, %tscallnumber%

return

; Call customer CBN via Dialer
callback:
   gui, submit, NoHide
   WinActivate,  Dialer - Internet Explorer
   
   ; Go To Dialer page
   pwb := WBGet()
   pwb.Navigate("https://dslcompany.com/phone/dialer/")
   Gosub, Wait
   
   ; Select DSL
   pwb.document.getElementById("progId").focus()
   sleep, 1000
   Send {Down}
   Send {Down}
   Send {Down}
   
   sleep, 500
   
   ; Select Free Form
   pwb.document.getElementById("dial").focus()
   sleep, 1000
   Send {Down}
   
   sleep, 500
   
   ; Put CBN in the clipboard
   clipboard =
   Clipboard := CBN
   
   ; Paste CBN
   Send {Tab}
   sleep, 500
   Send 1
   Send ^v
   
   sleep, 500
   
   ; Click 'Dial'
   pwb.document.GetElementsByTagName("INPUT").item[1].click()
   
return

; Open/Focus outage tool in Google Chrome with CLLI copied to the clipboard
outage:
   gui, submit, NoHide
   Clipboard =
   Clipboard := CLLI
   Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe "http://outage.company.com"
return

; ########## This section brings tool window focus to the foreground ##########
activatefiber_provisioning:
   gui, submit, NoHide
   WinActivate, Fiber Provisioning System ahk_exe iexplore.exe
return

activatedsl_provisioning:
   gui, submit, NoHide
   WinActivate, DSL Provisioning System ahk_exe iexplore.exe
   WinActivate, Login to DSL Provisioning System ahk_exe iexplore.exe
return

activatedsltest:
   gui, submit, NoHide
   WinActivate, DSLTest ahk_exe iexplore.exe
   WinActivate, autologoutLogin ahk_exe iexplore.exe
   WinActivate, Login ahk_class IEFrame ahk_exe iexplore.exe
return

activateticketsystem:
   gui, submit, NoHide
   gosub, ts_window_titles
return

activatelegacy_billing_system:
   gui, submit, NoHide
   WinActivate, ahk_exe legacy_billing_system.exe
return
; Subroutine to wait for page load before processing
Wait:
   while pwb.busy or pwb.ReadyState != 4
   Sleep, 100
return

; Various Ticket System window titles for focus/activation
ts_window_titles:
   WinActivate, Session Expired ahk_class IEFrame
   WinActivate, Service Portal ahk_class IEFrame
   WinActivate, Search For View Calls ahk_class IEFrame
   WinActivate, View Calls Search Results ahk_class IEFrame
   WinActivate, Call Summary ahk_class IEFrame
   WinActivate, https://ticketsystem.company.com ahk_class IEFrame
   WinActivate, Lookup Call Information ahk_class IEFrame
return

; To view the array while working on a particular section
printarray:
For i, value in GrabService
      List.= i ": " value "`r"
      
      MsgBox,,Array, % List
   List:=""
return

; Needed to use existing IE windows borrowed from Joe Glines https://the-automator.com/joe-glines-bio/ who borrowed it from someone else.  
WBGet(WinTitle="ahk_class IEFrame", Svr#=1) {               ;// based on ComObjQuery docs
   static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
        , IID := "{0002DF05-0000-0000-C000-000000000046}"   ;// IID_IWebBrowserApp
;//     , IID := "{332C4427-26CB-11D0-B483-00C04FD90119}"   ;// IID_IHTMLWindow2
   SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%

   if (ErrorLevel != "FAIL") {
      lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
      if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
         DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
         return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
      }
   }
}
; return

; Press ctrl+d to exit the app
^d::ExitApp

/*
TODO
Clean up with best practices - https://github.com/aviaryan/Ahk-Best-Practices
*/