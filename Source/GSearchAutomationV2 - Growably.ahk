#Requires AutoHotkey v2.0
#Include .\UIA.ahk

#SingleInstance

ESC:: ExitApp()
^p:: Pause(-1)
!r:: Reload

SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.

; Function to safely get UIA root element with retries
GetRootElementSafely(hwnd) {
    global retryCount := 5        ; Number of retries to resolve the UIA error
    global retryDelay := 500      ; Delay between retries (in milliseconds)
    Loop retryCount {
        try {
            ; Verify window exists and activate it
            if !WinExist("ahk_id " hwnd) {
                ToolTip("Window not found. Retrying...")
                Sleep(retryDelay)
                continue
            }
            WinActivate("ahk_id " hwnd)
            Sleep(100) ; Allow the window to activate
            
            ; Attempt to get the root element
            rootElement := UIA.ElementFromHandle(hwnd)
            if IsObject(rootElement) {
                return rootElement ; Success
            }
        } catch as e {
            ToolTip("UIA Error: " e.message ". Retrying...")
            Sleep(retryDelay)
        }

        if (retryCount == 3) {
            Click()
        }
    }
    MsgBox("Failed to retrieve UIA root element after retries.")
    ExitApp()
}

; Function to count elements in the array
GetArrayLength(arr) {

    count := 0
    for _, _ in arr {
        count++
    }
    return count
}

; Function to handle key input loop
GetSubKeys(inputValues) {
    ; Display Tooltip with sub-hotkey options and the option to select multiple
    ToolTip("Select multiple sub-hotkeys:`nW: Check for updates`nL: Log All Growably #'s into Notes`n1-3: Call Contact 1-3 & 4-6: Company 1-3`nT: Title search`nC: Contact info search`nN: News search`nE: Email address search`nP: Phone # search`nZ: time Zone search`nPress keys one after another.")

    ; Initialize an empty string for the subkeys
    SubKeys := ""

    ; Wait for the initial sub-hotkey press within 7 seconds
    ih := InputHook("L1 T7")
    ih.Start()
    ih.Wait()

    if (ih.EndReason == "Timeout") {
        ToolTip()  ; Turn off the tooltip
        return ""
    }

    SubKey := StrLower(ih.Input)

    ; Validate the first key press
    if InStr(inputValues, SubKey) {
        SubKeys .= SubKey
    } else {
        return GetSubKeys(inputValues)  ; Re-loop the input request
    }

    ; Loop to collect additional valid subkey inputs, unless 'a' is pressed
    Loop {
        ; Wait for the next key press within 1 second
        ih := InputHook("L1 T0.5")
        ih.Start()
        ih.Wait()

        if (ih.EndReason == "Timeout") {
            ToolTip()  ; Turn off the tooltip
            break
        }

        SubKey := StrLower(ih.Input)

        ; Validate the key press
        if InStr(inputValues, Subkey) {
            if (SubKey == "a") {
                SubKeys := "a"
                break
            }
            SubKeys .= SubKey
        } else {
            return GetSubKeys(inputValues)  ; Re-loop the input request
        }
    }

    ; Return the collected subkeys
    return SubKeys
}

; Function to combine keywords into a single search query
CombineKeywords(base, keywords) {
    combined := '"' base '"'  ; Ensure base text is treated as exact match
    combined .= " ("
    for index, keyword in keywords {
        if (index > 1)
            combined .= " OR "
        combined .= '"' keyword '"'
    }
    combined .= ")"
    return combined
}

; Function to perform Google search
PerformGoogleSearch(query) {
    query := URIEncode(query)  ; URL-encode the search query
    ; Open the search in a new browser tab
    Run("https://www.google.com/search?q=" query)
}

URIEncode(Url, Flags := 0x000C3000) {
	Local CC := 4096, Esc := "", Result := ""
    Loop {
		VarSetStrCapacity(&Esc, CC), Result := DllCall("Shlwapi.dll\UrlEscapeW", "Str", Url, "Str", &Esc, "UIntP", &CC, "UInt", Flags, "UInt")
	    if Result != 0x80004003 ; E_POINTER
        {
            break
        }
    }
    Return Esc
}

UrlUnescape(Url, Flags := 0x00140000) {
   Return !DllCall("Shlwapi.dll\UrlUnescape", "Ptr", StrPtr(Url), "Ptr", 0, "UInt", 0, "UInt", Flags, "UInt") ? Url : ""
}

; Function to process clipboard content into an object of search terms
ProcessClipboardContent(content) {
    ; Split content into separate terms by newlines or "; "
    terms := StrSplit(content, ["`n", "; "])
    cleanTerms := []
    for term in terms {
        term := Trim(term)  ; Remove leading and trailing spaces
        if (term != "" && !(InStr(term, "internal"))) {
            cleanTerms.Push(term)  ; Add non-empty terms to the array
        }
    }
    return cleanTerms
}

; Function to activate the Chrome window
ActivateChromeWindow() {
    if hwnd := WinExist("ahk_exe") {
        WinActivate()
        WinWaitActive("ahk_exe")
        return hwnd
    } else {
        MsgBox("Window not found.")
    }
}

ResetGlobals(names) {
    for name in StrSplit(names, ", ") {
        global name := ""
    }
}

GetGrowablyNames(SubKeys, rootElement) {
    ToolTip("Getting Growably names using " SubKeys ".")
    SetTimer () => ToolTip(""), -2000
    failure := false
    locationFailure := 0
    ResetGlobals("contactName, businessName, companyLocation")
    ; Get HWND and rootElement of the active Chrome window
    hwnd := ActivateChromeWindow()
    rootElement := GetRootElementSafely(hwnd)
    names := ""

    ; Get the contact's name from Growably
    if InStr(SubKeys, "p") || InStr(SubKeys, "e") || InStr(SubKeys, "c") || InStr(SubKeys, "a") || InStr(SubKeys, "n") || InStr(SubKeys, "@") {
        try {
            contactName := rootElement.WaitElementFromPath("VKr").Name
        } catch {
            contactName := ""
            ToolTip("Failed to get contact's name from Growably.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            failure := true
        }
    }

    ; Get the business name from Growably
    if InStr(SubKeys, "p") || InStr(SubKeys, "e") || InStr(SubKeys, "c") || InStr(SubKeys, "a") || InStr(SubKeys, "n") || InStr(SubKeys, "t") {
        try {
            businessName := rootElement.FindFirst({LocalizedType: "edit", Name: "Business Name"}).Value
        } catch {
            contactName := rootElement.FindFirst({LocalizedType: "edit", Name: "First Name"}).Click()
            Sleep(150)
            Send("{PgDn}")
            Sleep(150)
            Send("{PgDn}")
            Send("{PgDn}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Send("{PgUp}")
            Sleep(150)
            Sleep(500)
            businessName := rootElement.FindFirst({LocalizedType: "edit", Name: "Business Name"}).Value
            ToolTip("Failed to get business name from Growably (A).")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            failure := true
        }
    }

    try {
        if contactName {
            ; Do nothing
        }
    } catch {
        contactName := ""
    }

    try {
        if businessName {
            ; Do nothing
        }
    } catch {
        businessName := ""
    }

    if failure = true {
        MsgBox("Failed to get contact name or business name from Growably.`nContact name: " contactName "`nBusiness name: " businessName)
        failure := false
    }

    if InStr(SubKeys, "p") || InStr(SubKeys, "e") || InStr(SubKeys, "c") || InStr(SubKeys, "a") || InStr(SubKeys, "n") {
        ; Get the business name from Growably
        names := contactName "`" AND `"" businessName
    } else if InStr(SubKeys, "z") {
        GetCompanyLocation1:
        try {
            companyLocation := rootElement.FindFirst({LocalizedType: "edit", Name: "Company Location"})
            companyLocation := companyLocation.Value
            failure := false
            ToolTip("Got company location from Growably.`nCompany location: " companyLocation)
            SetTimer () => ToolTip(""), -2000
        } catch {
            locationFailure++
            Sleep(1000)
            if locationFailure = 1 {
                RetryContact1:
                try {
                    contactName := rootElement.FindFirst({LocalizedType: "edit", Name: "Last Name"})
                } catch {
                    contactName := ""
                }
                if contactName = "" {
                    try {
                        contactName := rootElement.FindFirst({LocalizedType: "edit", Name: "First Name"})
                    } catch {
                        contactName := ""
                    }
                }
                if contactName = "" {
                    contactName := rootElement.FindAll({LocalizedType: "text", Name: "Contact"})
                    for i, contact in contactName {
                        if i = 2 {
                            contact.Click()
                        }
                    }
                    goto RetryContact1
                }
                
                if StrLen(contactName.Name) > 1 {
                    contactName.Click()
                } else {
                    try {
                        contactName := rootElement.FindFirst({LocalizedType: "edit", Name: "First Name"})
                        contactName.Click()
                    } catch {
                        contactName := ""
                    }
                    
                }
                Sleep(150)
                Send("{PgDn}")
                Sleep(150)
                Send("{PgDn}")
                Sleep(150)
                Send("{PgDn}")
                Sleep(150)
                Send("{PgDn}")
                Sleep(600)
                goto GetCompanyLocation1
            }
            companyLocation := ""
            ToolTip("Failed to get company location from Growably.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            failure := true
        }

        if !failure {
            if companyLocation = "United States" {
                companyLocation := rootElement.FindFirst({LocalizedType: "edit", Name: "Contact Location"}).Value
            }
            companyLocation := RegExReplace(companyLocation, "\b\d{5}(-\d{4})?\b")
            companyLocation := RegExReplace(companyLocation, "\d+")
            parts := StrSplit(companyLocation, ", ")
            if GetArrayLength(parts) > 3 {
                companyLocation := ""
                partsCount := GetArrayLength(parts)
                ; Loop from the second part onward
                Loop partsCount
                {
                    if (A_Index > 1) ; Skip the first part
                        companyLocation .= (A_Index > 2 ? ", " : "") . parts[A_Index]
                }

                names := companyLocation
            } else {
                names := companyLocation
            }
        } else {
            ToolTip("Failed to get company location from Growably.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            names := ""
        }
        
    } else if InStr(SubKeys, "t") {
        names := businessName
    } else if InStr(SubKeys, "@") {
        names := contactName
    }

    if failure && (InStr(SubKeys, "p") || InStr(SubKeys, "e") && InStr(SubKeys, "c") && InStr(SubKeys, "a") && InStr(SubKeys, "n")) {
        MsgBox("Failed to get task-related information from Growably.`nContact name: " contactName "`nCompany location: " companyLocation)
    } else if failure && InStr(SubKeys, "z") {
        MsgBox("Failed to get task-related information from Growably.`nCompany location: " companyLocation)
    } else if failure {
        MsgBox("Failed to get task-related information from Growably.`nBusiness name: " businessName "`nCompany location: " companyLocation)
    }

    return names
}

updateGrowablyTimeZone(timezone, savedNames, rootElement) {

    updateGrowablyTimeZone1:
    currentName := ""
    successTries := 0

    if timezone = "" || timezone = "null" || timezone = "false" || timezone = "0" {
        names := GetGrowablyNames("c", rootElement)
        if names = "" || names = 0 {
            names := GetGrowablyNames("@", rootElement)
        }
        if names = "" || names = 0 {
            FileAppend("Failed to get & set time zone for: " names "`n", "FailedTasks.txt")
            return false
        }
    }

    ;MsgBox("Setting time zone to: " timezone)

    if timezone == "PST" {
        timezone := "GMT-08:00 America/Los_Angeles (PST)"
    } else if timezone == "MST" {
        timezone := "GMT-07:00 US/Mountain (MST)"
    } else if timezone == "CST" {
        timezone := "GMT-06:00 US/Central (CST)"
    } else if timezone == "EST" {
        timezone := "GMT-05:00 US/Eastern (EST)"
    } else if timezone == "AKST" {
        timezone := "GMT-09:00 US/Alaska (AKST)"
    } else if timezone == "HST" {
        timezone := "GMT-10:00 Pacific/Honolulu (HST)"
    }

    global timeZoneComboBox := ""
    success := false
    ; Get HWND and rootElement of the active Chrome window
    Sleep(600)

    ; Find all combo box elements
    comboBoxes := rootElement.FindAll({LocalizedType: "combo box"})

    ; Filter elements that start with "GMT"
    for comboBox in comboBoxes {
        if InStr(comboBox.Value, "GMT") {
            timeZoneComboBox := comboBox

            if comboBox.Value == timezone {
                return true
            }

            ; Click the combo box
            timeZoneComboBox.Click()
            Sleep(1000)
            ; Wait for the list items to appear and then click the correct one
            Loop {
                success := false
                listItems := rootElement.FindAll({LocalizedType: "list item"})

                for i, item in listItems {
                    if InStr(item.Name, timezone) {
                        item.Click()
                        success := true
                        break
                    }
                }
                if success {
                    break
                } else {
                    ToolTip("Failed to find time zone combo box. (match)")
                    SetTimer () => ToolTip(""), -2000
                    goto UpdateGrowablyTimeZone1
                }
            }
            
            break
        }
    }

    if !success && successTries < 4 {
        ToolTip("Failed to find time zone combo box. (3 or under - no match)")
        successTries++
        goto UpdateGrowablyTimeZone1
    }


    if !success && successTries > 3 {
        ToolTip("Failed to find time zone combo box. (no match)")
        SetTimer () => ToolTip(""), -2000
        Sleep(2000)
        goto UpdateGrowablyTimeZone1
    } else {
        Sleep(500)
        Loop {
            Sleep(500)

            currentName := GetGrowablyNames("@", rootElement)
            if currentName = "" || currentName = 0 {
                currentName := GetGrowablyNames("c", rootElement)
            }
            if savedNames == currentName {
                ; Do nothing
            } else {
                Sleep(1500)
                break
            }
        }
        
        savedNames := currentName
        return true
    }

    MsgBox("Failed to handle setting time zone correctly.")
}

UpdateLastNameGrowably(lastName) {
    Sleep(300)
    hwnd := ActivateChromeWindow()
    rootElement := GetRootElementSafely(hwnd)

    if lastName == "" {
        ToolTip("Google search came up empty.")
        return false
    }

    try {
        lastNameElement := rootElement.FindFirst({LocalizedType: "edit", Name: "Last Name"})
        if StrLen(lastNameElement.Value) < StrLen(lastName) {
            lastNameElement.Value := lastName 
            return true
        } else {
            return false
        }
        
    } catch {
        ToolTip("Failed to update last name within Growably.")
        SetTimer () => ToolTip(""), -2000
        Sleep(2000)
        return false
    }
    
}

GetInfoFromSearch(searchType, name := "") {
    Sleep(500)
    ;global start := A_TickCount
    ;static maxTime := 5000
    global tries := 0
    ;Loop 50 {
    ;    elapsedTime := (A_TickCount - start) / 1000
    ;    if elapsedTime > maxTime {
    ;        ToolTip("")
    ;        break
    ;    }

    ;    if InStr(WinGetTitle("A"), "Google Search") {
    ;        currentTitle := WinGetTitle("A")
    ;        Sleep(500)
    ;        break
    ;    }

    ;    hwnd := ActivateChromeWindow()
    ;    rootElement := GetRootElementSafely(hwnd)
        ;ToolTip("Elapsed Time: " Round(elapsedTime, 3) " seconds")
        ;SetTimer () => ToolTip(""), -1000
        ;Sleep(1000)
    ;    try {
    ;        captchaElement := rootElement.FindFirst({AutomationID: "recaptcha-anchor"})
    ;        ClickElementByPath(captchaElement)
    ;        Sleep(500)
    ;    } catch {
    ;        captchaElement := ""
    ;    }

    ;    try {
    ;        captchaElement := rootElement.FindFirst({LocalizedType: "check box", Name: "I'm not a robot"})
    ;        ClickElementByPath(captchaElement)
    ;        Sleep(500)
    ;    } catch {
    ;        captchaElement := ""
    ;    }

    ;    try {
    ;        captchaElement := rootElement.FindFirst({LocalizedType: "button", Name: "I'm not a robot"})
    ;        ClickElementByPath(captchaElement)
    ;        Sleep(500)
    ;    } catch {
    ;        captchaElement := ""
    ;    }

    ;    try {
    ;        ClickElementByPath("VRb")
    ;    } catch {
    ;        captchaElement := ""
    ;    }

    ;    Sleep(100)
    ;}
    start := 0
    Sleep(500)
    ToolTip("Getting info from Google Search.")
    SetTimer () => ToolTip(""), -2000
    ; Get HWND and rootElement of the active Chrome window
    hwnd := ActivateChromeWindow()
    rootElement := GetRootElementSafely(hwnd)

    if searchType == "timezone" {
        tries := 0

        try {
            timeZoneText := rootElement.FindAll({LocalizedType: "text"})

            TimeZoneCheck1:
            
            for timeZone in timeZoneText {
                if InStr(timeZone.Name, "did not match") || InStr(timeZone.Value, "did not match") || InStr(timeZone.Name, "Try using words that might") || InStr(timeZone.Value, "Try using words that might") {
                    Sleep(100)
                    Send("^w")
                    Sleep(100)
                    return ""
                }
                if InStr(timeZone.Name, "EST") || InStr(timeZone.Name, "EDT") || InStr(timeZone.Name, "GMT-5") || InStr(timeZone.Name, "eastern") {
                    Send("^w")
                    Sleep(100)
                    return "GMT-05:00 US/Eastern (EST)"
                }
                if InStr(timeZone.Name, "CST") || InStr(timeZone.Name, "CDT") || InStr(timeZone.Name, "GMT-6") || InStr(timeZone.Name, "central") {
                    Send("^w")
                    Sleep(100)
                    return "GMT-06:00 US/Central (CST)"
                }
                if InStr(timeZone.Name, "MST") || InStr(timeZone.Name, "MDT") || InStr(timeZone.Name, "GMT-7") || InStr(timeZone.Name, "mountain") {
                    Send("^w")
                    Sleep(100)
                    return "GMT-07:00 US/Mountain (MST)"
                }
                if InStr(timeZone.Name, "PST") || InStr(timeZone.Name, "PDT") || InStr(timeZone.Name, "GMT-8") || InStr(timeZone.Name, "pacific") {
                    Send("^w")
                    Sleep(100)
                    return "GMT-08:00 America/Los_Angeles (PST)"
                }
                if InStr(timeZone.Name, "AKST") || InStr(timeZone.Name, "AKDT") || InStr(timeZone.Name, "GMT-9") || InStr(timeZone.Name, "alaska") {
                    Send("^w")
                    Sleep(100)
                    return "GMT-09:00 US/Alaska (AKST)"
                }
                if InStr(timeZone.Name, "HST") || InStr(timeZone.Name, "HDT") || InStr(timeZone.Name, "GMT-10") || InStr(timeZone.Name, "hawaii") {
                    Send("^w")
                    Sleep(100)
                    return "GMT-10:00 Pacific/Honolulu (HST)"
                }
            }

            if tries = 0 {
                timeZoneText := rootElement.FindAll({LocalizedType: "header"})
                tries++
                goto TimeZoneCheck1
            }
            if tries = 1 {
                timeZoneText := rootElement.FindAll({LocalizedType: "item"})
                tries++
                goto TimeZoneCheck1
            }
            if tries = 2 {
                for timeZone in timeZoneText {
                    global strText := timeZone.Name "`n"
                }
                Send("^w")
                Sleep(100)
                names := GetGrowablyNames("c", rootElement)
                if names = "" || names = 0 {
                    names := GetGrowablyNames("@", rootElement)
                }
                FileAppend("Failed to get & set time zone for: " names "`n", "FailedTasks.txt")
                ;MsgBox("strText for time zone check failure: `n" strText)
                ToolTip("Failed to get time zone from Google (match).")
                SetTimer () => ToolTip(""), -2000
                Sleep(2000)
                return false
            }
        } catch {
            ToolTip("Failed to get time zone from Google (catch).")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            Send("^w")
            Sleep(100)
            return false
        }
    }

    if searchType == "lastname" {
        try {
            ; Find all text elements
            textElements := rootElement.FindAll({LocalizedType: "text"})
            for textElement in textElements {
                if InStr(textElement.Name, name) {
                    fullName := textElement.Name
                    lastName := StrSplit(fullName, " ")
                    lastName := lastName[2]
                    return lastName
                }
            }

            ToolTip("Failed to get last name from Google (match 2).")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            Send("^w")
            Sleep(100)
            return false
        } catch {
            ToolTip("Failed to get last name from Google (catch 2).")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            Send("^w")
            Sleep(100)
            return false
        }
    }
    /*
    if searchType == "phone number" {
        tries := 0

        try {
            phoneNumberText := rootElement.FindAll({LocalizedType: "text"})

            PhoneNumberCheck1:
            
            for phoneNumber in phoneNumberText {
                if InStr(phoneNumber.Name, "did not match") || InStr(phoneNumber.Value, "did not match") {
                    Send("^w")
                    Sleep(100)
                    return ""
                }
                if InStr(phoneNumber.Name, name) {
                    fullName := textElement.Name
                    lastName := StrSplit(fullName, " ")
                    lastName := lastName[2]
                    return lastName
                }
            }
            if tries = 0 {
                phoneNumberText := rootElement.FindAll({LocalizedType: "header"})
                tries := 1
                goto PhoneNumberCheck1
            }
            if tries = 1 {
                phoneNumberText := rootElement.FindAll({LocalizedType: "item"})
                tries := 2
                goto PhoneNumberCheck1
            }

            if tries = 2 {
                for phoneNumber in phoneNumberText {
                    global strText := phoneNumber.Name "`n"
                }
                names := GetGrowablyNames("c")
                FileAppend("Failed to get & set phone number for: " names, "FailedTasks.txt")
                ;MsgBox("strText for time zone check failure: `n" strText)
                ToolTip("Failed to phone number from Google (match).")
                SetTimer () => ToolTip(""), -2000
                Sleep(2000)
                Send("^w")
                Sleep(100)
                return false
            }
        } catch {
            ToolTip("Failed to get time zone from Google (catch).")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            Send("^w")
            Sleep(100)
            return false
        }
    }
    */
}

CallWithRingCentral(finalNumber) {
    WinActivate("RingCentral")
    Sleep(60)
    Send(finalNumber)
}

CallContact(callNumber, rootElement) {

    if (callNumber >= 1 && callNumber <= 3) {
        try {
            callNumberName := "Contact Phone " . callNumber
            callNumberElement := rootElement.FindFirst({LocalizedType: "edit", Name: callNumberName})
            ToolTip("Calling " callNumberName " `n" callNumberElement.Value)
            SetTimer () => ToolTip(""), -2000
            if callNumberElement.Value == "" {
                ToolTip("No " callNumberName " to grab.")
                SetTimer () => ToolTip(""), -2000
                return false
            }
            A_Clipboard := callNumberElement.Value
            Sleep(40)
        } catch {
            ToolTip("Failed to get " callNumberName " from Growably. Did you mean to select a different number?")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        editNote := rootElement.WaitElement({LocalizedType: "edit", Name: "New note"}, 1000)
        if editNote == 0 {
            ; Do nothing
        } else {
            editNote.Click()
            Send("{Enter}")
            Sleep(60)
            Send("^v")
            Sleep(60)
            Send(" - ")
            goto CallWithRingCentral1to3
        }
        Sleep(500)
        notesTab := rootElement.FindFirst({AutomationID: "notes-tab"})
        notesTab.Click()
        Sleep(500)
        addNoteButton := rootElement.WaitElement({LocalizedType: "link" , Name: "+ Add"}, 1000)
        if addNoteButton == 0 {
            ToolTip("Failed to find add note button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        addNoteButton.Click()
        editNote := rootElement.WaitElement({LocalizedType: "edit", Name: "New note"}, 1000)
        if editNote == 0 {
            ToolTip("Failed to find edit note button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        editNote.Click()
        Send(callNumberElement.Value " - ")
        
        CallWithRingCentral1to3:
        finalNumber := callNumberElement.Value
        Sleep(300)
        try {
            CallWithRingCentral(finalNumber)
        } catch {
            ToolTip("Failed to find and call " callNumberName " using RingCentral. Are you sure RingCentral is running?")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
    }
        
    if (callNumber >= 4 && callNumber <= 6) {
        try {
            callnumberName := "Company Phone " . callNumber - 3
            callNumberElement := rootElement.FindFirst({LocalizedType: "edit", Name: callNumberName})
            ToolTip("Calling " callNumberName " `n" callNumberElement.Value)
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            if callNumberElement.Value == "" {
                ToolTip("No " callNumberName " to grab.")
                SetTimer () => ToolTip(""), -2000
                return false
            }
            A_Clipboard := callNumberElement.Value
            Sleep(40)
        } catch {
            ToolTip("Failed to get " callNumberName " from Growably.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        editNote := rootElement.WaitElement({LocalizedType: "edit", Name: "New note"}, 1000)
        if editNote == 0 {
            ; Do nothing
        } else {
            editNote.Click()
            Send("{Enter}")
            Sleep(60)
            Send("^v")
            Sleep(60)
            Send(" - ")
            goto CallWithRingCentral4to6
        }
        Sleep(500)
        notesTab := rootElement.FindFirst({AutomationID: "notes-tab"})
        notesTab.Click()
        Sleep(500)
        addNoteButton := rootElement.WaitElement({LocalizedType: "link" , Name: "+ Add"}, 1000)
        if addNoteButton == 0 {
            ToolTip("Failed to find add note button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        addNoteButton.Click()
        editNote := rootElement.WaitElement({LocalizedType: "edit", Name: "New note"}, 1000)
        if editNote == 0 {
            ToolTip("Failed to find edit note button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        editNote.Click()
        Send(callNumberElement.Value " - ")

        CallWithRingCentral4to6:
        finalNumber := callNumberElement.Value
        Sleep(300)
        try {
            CallWithRingCentral(finalNumber)
        } catch {
            ToolTip("Failed to find and call " callNumberName " using RingCentral. Are you sure RingCentral is running?")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
    }

    return true
}

LoadNumbersIntoNotes(rootElement) {
    global allNumbers := ""

    Loop 7 {
        global allNumbers
        if (A_Index >= 1 && A_Index <= 3) {
            try {
                callNumberName := "Contact Phone " . A_Index
                callNumberElement := rootElement.FindFirst({LocalizedType: "edit", Name: callNumberName})
                if callNumberElement.Value == "" {
                    ToolTip("No " callNumberName " to grab.")
                    SetTimer () => ToolTip(""), -2000
                    continue
                }
                if InStr(allNumbers, callNumberElement.Value) {
                    continue
                }
                allNumbers := allNumbers callNumberElement.Value " - `n"
                Sleep(60)
            } catch {
                callNumberElement := ""
                ToolTip("Failed to get " callNumberName " from Growably. Did you mean to select a different number?")
                SetTimer () => ToolTip(""), -2000
            }
        }

        if (A_Index >= 4 && A_Index <= 6) {
            try {
                callNumberName := "Company Phone " . A_Index - 3
                callNumberElement := rootElement.FindFirst({LocalizedType: "edit", Name: callNumberName})
                if callNumberElement.Value == "" {
                    ToolTip("No " callNumberName " to grab.")
                    SetTimer () => ToolTip(""), -2000
                    continue
                }
                if InStr(allNumbers, callNumberElement.Value) {
                    continue
                }
                allNumbers := allNumbers callNumberElement.Value " - `n"
                Sleep(60)
            } catch {
                callNumberElement := ""
                ToolTip("Failed to get " callNumberName " from Growably. Did you mean to select a different number?")
                SetTimer () => ToolTip(""), -2000
            }
        }

        if (A_Index == 7) {
            try {
                callNumberName := "Contact Mobile Phone"
                callNumberElement := rootElement.FindFirst({LocalizedType: "edit", Name: callNumberName})
                if callNumberElement.Value == "" {
                    ToolTip("No " callNumberName " to grab.")
                    SetTimer () => ToolTip(""), -2000
                    continue
                }
                if InStr(allNumbers, callNumberElement.Value) {
                    continue
                }
                allNumbers := allNumbers callNumberElement.Value " - `n"
                Sleep(60)
            } catch {
                callNumberElement := ""
            }
            
        }
    }

    if allNumbers == "" {
        ToolTip("No numbers to load into notes.")
        SetTimer () => ToolTip(""), -2000
        return false
    }

    try {
        notesTab := rootElement.FindFirst({AutomationID: "notes-tab"})
        notesTab.Click()
        Sleep(500)
        addNoteButton := rootElement.WaitElement({LocalizedType: "link" , Name: "+ Add"}, 1000)
        if addNoteButton == 0 {
            ToolTip("Failed to find add note button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        addNoteButton.Click()
        editNote := rootElement.WaitElement({LocalizedType: "edit", Name: "New note"}, 1000)
        if editNote == 0 {
            ToolTip("Failed to find edit note button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        editNote.Click()
        Sleep(40)
        Send(allNumbers)
        Sleep(1000)
        saveBtn := rootElement.WaitElement({LocalizedType: "button", Name: "Save"}, 1000)
        if saveBtn == 0 {
            ToolTip("Failed to find save button.")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            return false
        }
        saveBtn.Click()
        return true
    } catch {
        ToolTip("Failed to load numbers into notes.")
        SetTimer () => ToolTip(""), -2000
        Sleep(2000)
        return false
    }
}

CheckForUpdates(type, SubKeys, rootElement) {

    global strSubKeys := ""
    global foundPhone := false
    global phoneName := ""
    global phoneValue := ""
    ;ToolTip("Checking for updates.")
    ;SetTimer () => ToolTip(""), -2000

    strSubKeys .= "z" ; Always add "z" for checking time zone
    if type == "last name" || type == "all" {
        try {
            updateButtonA := rootElement.FindFirst({LocalizedType: "edit", Name: "Last Name"})
            if StrLen(updateButtonA.Value) < 2 {
                strSubKeys .= "t"
            }
        } catch {
            ToolTip("Didn't find last name in Growably, did the field name change?")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            strSubKeys .= "t"
        }
    }

    if type == "title" || type == "all" {
        try {
            updateButtonB := rootElement.FindFirst({LocalizedType: "edit", Name: "Title"})
            if StrLen(updateButtonB.Value) < 2 {
                strSubKeys .= "c"
            }
        } catch {
            ToolTip("Didn't find title in Growably, did the field name change?")
            SetTimer () => ToolTip(""), -2000
            Sleep(2000)
            strSubKeys .= "c"
        }   
    }
    
    if type == "phone" || type == "all" {
        allElements := rootElement.FindAll({LocalizedType: "edit"})
        for i, element in allElements {
            if InStr(element.Name, "Contact Phone 1") {
                phoneName := element.Name
                phoneValue := element.Value
                foundPhone := true
                ;MsgBox("Found phone name and value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
            } else if InStr(element.Name, "Contact Phone 2") {
                phoneName := element.Name
                phoneValue := element.Value
                foundPhone := true
                ;MsgBox("Found phone name and value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
            } else if InStr(element.Name, "Contact Phone 3") {
                phoneName := element.Name
                phoneValue := element.Value
                foundPhone := true
                ;MsgBox("Found phone name and value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
            } else if InStr(element.Name, "Company Phone 1") {
                phoneName := element.Name
                phoneValue := element.Value
                foundPhone := true
                ;MsgBox("Found phone name and value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
            } else if InStr(element.Name, "Company Phone 2") {
                phoneName := element.Name
                phoneValue := element.Value
                foundPhone := true
                ;MsgBox("Found phone name and value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
            } else if InStr(element.Name, "Company Phone 3") {
                phoneName := element.Name
                phoneValue := element.Value
                foundPhone := true
                ;MsgBox("Found phone name and value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
            }
            if foundPhone && (0 < StrLen(phoneValue)) && (StrLen(phoneName) < 2) {
                MsgBox("Found phone name but no value.`nPhone name: " phoneName "`nPhone value: " phoneValue)
                names := GetGrowablyNames("p", rootElement)
                FileAppend("No phone number listed for : " names "`n", "FailedTasks.txt")
                break
            }
        }

        if !foundPhone {
            MsgBox("Failed to find phone in Growably.")
            names := GetGrowablyNames("p", rootElement)
            FileAppend("No phone number listed for : " names "`n", "FailedTasks.txt")
        } else {
            foundPhone := false
        }
    }

    ;ToolTip("Updating with: " strSubKeys)
    ;SetTimer () => ToolTip(""), -2000
    SubKeys := Trim(SubKeys, "w")
    ToolTip("Updating with: " strSubKeys)
    SetTimer () => ToolTip(""), -2000
    Sleep(500)
    SubKeys := strSubKeys . SubKeys
    return strSubKeys
}

; Main function to handle Google Search Automation
F1:: {

    ; Define acceptable subkeys for inputValues
    global inputValues := "tcnpeazwqlu123456"
    global contactName
    global businessName
    global companyLocation
    global boolGetTimeZone := false boolGetName := false
    global timeZone
    global SubKeys := ""
    global updatesChecked := false
    global failedNextBtn := false
    global noSearch := false
    global pacificStates := [" CA ", " CA,", "California", " WA ", " WA,", " NV ", " NV,", "Nevada", " OR ", " OR,", "Oregon"]
    global mountainStates := [" ID ", " ID,", "Indiana", " MT ", " MT,", "Montana", " WY ", " WY,", "Wyoming", " UT ", " UT,", "Utah", " CO ", " CO,", "Colorado", " AZ ", " AZ,", "Arizona", " NM ", " NM,", "New Mexico"]
    global centralStates := [" ND ", " ND,", "North Dakota", " SD ", " SD,", "South Dakota", " NE ", " NE,", "Nebraska", " MN ", " MN,", "Minnesota", " IA ", " IA,", "Iowa", " WI ", " WI,", "Wisconsin", " IL ", " IL,", "Illinois", " MO ", " MO,", "Missouri", " KS ", " KS,", "Kansas", " OK ", " OK,", "Oklahoma", " AR ", " AR,", "Arizona", " TX ", " TX,", "Texas", " LA ", " LA,", "Louisiana", " MS ", " MS,", "Missouri", " AL ", " AL,", "Alabama"]
    global easternStates := [" FL ", " FL,", "Florida", " GA ", " GA,", "Georgia", " SC ", " SC,", "Sourth Carolina", " NC ", " NC,", "North Carolina", " VA ", " VA,", "Virginia", " WV ", " WV,", "West Virginia", " OH ", " OH,", "Ohio", " MI ", " MI,", "Mississippi", " PA ", " PA,", "Pennsylvania", " DC ", " DC,", "District of Columbia", "D.C.", " MD ", " MD,", "Maryland", " DE ", " DE,", "Delaware", " NJ ", " NJ,", "New Jersey", " NY ", " NY,", "New York", "York New", " CT ", " CT,", "Connecticut", " RI ", " RI,", "Rhode Island", " MA ", " MA,", "Massachusetts", " NH ", " NH,", "New Hampshire", " ME ", " ME,", "Maine", " VT ", " VT,", "Vermont"]
    global alaskaState := [" AK ", " AK,", "Alaska"]
    global hawaiiState := [" HI ", " HI,", "Hawaii"]

    ;global boolGetPhone := false

    ; Display Tooltip with sub-hotkey options and the option to select multiple
    SubKeys := GetSubKeys(inputValues)

    ; If no valid keys were entered, exit the script
    if (SubKeys = "") {
        Exit
    }

    hwnd := ActivateChromeWindow()
    Sleep(40)
    rootElement := GetRootElementSafely(hwnd)

    ; Hide the Tooltip after keys are pressed
    ToolTip()

    ; Save the current clipboard content
    ClipSaved := ClipboardAll()
    Sleep(40)
    ; Clear the clipboard to avoid old data interference
    A_Clipboard := ""
    Sleep(40)
    ; Copy the selected text
    Send("^c")
    gotClip := ClipWait(1)
    Sleep(40)

    global savedNames := ""
    RunScript:
    global numGoogleSearches := StrLen(SubKeys)

    ; We will process each subkey individually.
    ; If we have multiple subkeys, we do separate searches for each.
    for i, singleKey in StrSplit(SubKeys) {

        if singleKey = "u" {
            UIA.Viewer()
        }

        if singleKey = "l" {
            LoadNumbersIntoNotes(rootElement)
            continue
        }

        if singleKey = "w" {
            SubKeys := CheckForUpdates("all", SubKeys, rootElement)
            updatesChecked := true
            goto RunScript
        }

        if singleKey = "q" && !updatesChecked {
            SubKeys := CheckForUpdates("", SubKeys, rootElement) . SubKeys
            updatesChecked := true
            goto RunScript
        }

        if singleKey = "q" {
            Sleep(300)
            permNextButton := rootElement.WaitElementFromPath({T:30}, {T:20, i:11})
            if permNextButton.Name == "First Name" || permNextButton.Name == "Followers" {
                permNextButton := rootElement.WaitElementFromPath({T:30}, {T:20, i:7}).Click()
            } else if InStr(permNextButton.Name, "Auto Save") {
                permNextButton := rootElement.WaitElementFromPath({T:30}, {T:20, i:6}).Click()
            } else {
                permNextButton.Click()
            }

            Sleep(2000)
            SubKeys := "q"
            updatesChecked := false
            goto RunScript
        }

        if singleKey = "1" || singleKey = "2" || singleKey = "3" || singleKey = "4" || singleKey = "5" || singleKey = "6" {
            if CallContact(singleKey, rootElement) {
                ToolTip("Successfully Grabbed Number")
                SetTimer () => ToolTip(""), -2000
                continue
            } else {
                continue
            }
        }

        ; If we got clipboard text successfully, use that as the search base.
        ; Otherwise, try to get names from Growably for this singleKey.
        if (gotClip) && !(InStr(gotClip, "internal")) {
            clipboardContent := A_Clipboard
            if (clipboardContent = "") {
                ; If somehow clipboard is empty, fallback to growably
                names := GetGrowablyNames(singleKey, rootElement)
                if (names = "false" || names = "") {
                    ToolTip("No text was selected and unable to get data from Growably for subkey: " singleKey)
                    SetTimer () => ToolTip(""), -2000
                    continue
                }
                savedNames := names
            } else {
                names := clipboardContent
                savedNames := names
            }
        } else {
            ; No clipboard text found, get from growably
            names := GetGrowablyNames(singleKey, rootElement)
            if (names = "false" || names = "" || InStr(names, "internal")) {
                ToolTip("No text was selected (but you knew that) and unable to get data from Growably for subkey: " singleKey)
                SetTimer () => ToolTip(""), -2000
                continue
            }
            savedNames := names
        }

        ; Determine keywords based on singleKey
        Keywords := []
        if InStr(singleKey, "z") || singleKey = "a" {
            for state in pacificStates {
                if InStr(names, state) {
                    updateGrowablyTimeZone("PST", savedNames, rootElement)
                    noSearch := true
                    break
                }
            }

            for state in mountainStates {
                if InStr(names, state) {
                    updateGrowablyTimeZone("MST", savedNames, rootElement)
                    noSearch := true
                    break
                }
            }

            for state in centralStates {
                if InStr(names, state) {
                    updateGrowablyTimeZone("CST", savedNames, rootElement)
                    noSearch := true
                    break
                }
            }

            for state in easternStates {
                if InStr(names, state) {
                    updateGrowablyTimeZone("EST", savedNames, rootElement)
                    noSearch := true
                    break
                }
            }

            for state in alaskaState {
                if InStr(names, state) {
                    updateGrowablyTimeZone("AKST", savedNames, rootElement)
                    noSearch := true
                    break
                }
            }

            for state in hawaiiState {
                if InStr(names, state) {
                    updateGrowablyTimeZone("HST", savedNames, rootElement)
                    noSearch := true
                    break
                }
            }

            if noSearch {
                noSearch := false
                continue
            }

            Keywords.Push("time zone")
            boolGetTimeZone := true
        }

        if InStr(singleKey, "t") {
            Keywords.Push("COO", "CEO", "current COO", "current CEO", "current CFO")
            boolGetName := true
        }
        if singleKey = "c" || singleKey = "a" {
            Keywords.Push("LinkedIn", "website", "email address", "phone number")
        }
        if InStr(singleKey, "n") || singleKey = "a" {
            Keywords.Push("company news", "financial reports", "company profile", "BBB", "Facebook", "recent events")
        }
        if InStr(singleKey, "p") || singleKey = "a" {
            Keywords.Push("phone number", "direct line", "mobile number", "contact")
            ;boolGetPhone := true
        }
        if InStr(singleKey, "e") || singleKey = "a" {
            Keywords.Push("email address", "contact email", "contact")
        }

        if boolGetName {
            if names = "false" || names = "" || names = 0 || names = "false" || names = "null" {
                names := GetGrowablyNames("t", rootElement)
            }

            ; Get the name of the contact
            contactFullName := GetGrowablyNames("@", rootElement)
        }

        searchTerms := ProcessClipboardContent(names)

        for searchTerm in searchTerms {
            combinedQuery := CombineKeywords(searchTerm, Keywords)
            PerformGoogleSearch(combinedQuery)
        }

        /*
        if boolGetPhone {
            phoneNumber := GetInfoFromSearch("phone number")
            boolGetPhone := false
            Sleep(1000)
            if phoneNumber != "" || phoneNumber != "null" || phoneNumber != "false" || phoneNumber != "0" {
                UpdateGrowablyPhoneNumber(phoneNumber)
            } else {
                ToolTip("Failed to update phone number within Growably.")
                SetTimer () => ToolTip(""), -2000
                Sleep(2000)
                continue
            }
        }
        */

        if boolGetTimeZone {
            timeZone := GetInfoFromSearch("timezone")
            boolGetTimeZone := false
            Sleep(1000)
            if timeZone != "" || timeZone != "null" || timeZone != "false" || timeZone != "0" {
                UpdateGrowablyTimeZone(timezone, savedNames, rootElement)
            } else {
                ToolTip("Failed to update time zone within Growably.")
                SetTimer () => ToolTip(""), -2000
                Sleep(2000)
                continue
            }
        }

        if boolGetName {
            boolGetName := false
            ToolTip("Sleep for 1 second(A).")
            SetTimer () => ToolTip(""), -1000
            Sleep(1000)
            lastName := GetInfoFromSearch("lastname", contactFullName)
            Send("^w")
            Sleep(100)
            success := UpdateLastNameGrowably(lastName)

            if !success {
                FileAppend("Failed to update last name within Growably.`nLast Name: " lastName "`n", "GSearchAutomationV2_debug_log.txt")
                ToolTip("Failed to update last name within Growably.")
                SetTimer () => ToolTip(""), -2000
                Sleep(2000)
                continue
            } else {
                FileAppend("Successfully updated last name within Growably.`nLast Name: " lastName "`n", "GSearchAutomationV2_debug_log.txt")
                ToolTip("Successfully updated last name within Growably.")
                SetTimer () => ToolTip(""), -2000
                Sleep(2000)
            }
        }

        if i >= numGoogleSearches {
            break
        }
    }

    contactName := ""
    businessName := ""
    companyLocation := ""
    savedNames := ""

    A_Clipboard := ClipSaved
    return
}

CoordMode("Mouse", "Screen")

; Function to introduce a random delay
RandomDelay(min := 200, max := 800) {
    delay := Random(min, max)
    Sleep delay
}

; Function to move the mouse with random variability
RandomMouseMove(x, y, elementRight, elementBottom, rangeX := elementRight - x / 2, rangeY := elementBottom - y / 2, speed := 10, padding := 10) {
    ; Ensure rangeX and rangeY are positive and stay within the element's dimensions
    rangeX := rangeX - padding
    rangeY := rangeY - padding

    rangeX := Max(rangeX, 0)
    rangeY := Max(rangeY, 0)
    
    if elementRight < x || elementBottom < y {
        elementRight := elementRight + x
        elementBottom := elementBottom + y
        rangeX := elementRight / 2
        rangeY := elementBottom / 2
    }
    averageX := (elementRight - x) / 2 + x
    averageY := (elementBottom - y) / 2 + y
    
    ; Ensure the random movement stays within the element's boundaries
    ; Calculate minimum and maximum X/Y values based on range and element boundaries
    minX := Max(averageX - rangeX, x)  ; Ensure minimum X stays within the element's left boundary
    maxX := Min(averageX + rangeX, elementRight - padding)  ; Ensure maximum X stays within the element's right boundary
    minY := Max(averageY - rangeY, y)  ; Ensure minimum Y stays within the element's top boundary
    maxY := Min(averageY + rangeY, elementBottom - padding)  ; Ensure maximum Y stays within the element's bottom boundary
    ;MsgBox("Moving mouse non-randomly to X=" minX " Y=" minY " with rangeX=" rangeX " and rangeY= " rangeY)
    
    ; Randomize movement within the calculated boundary range
    moveX := Round(Random(minX, maxX))
    moveY := Round(Random(minY, maxY))
    
    ; Debugging: Log the calculated random movement positions
    ;MsgBox("Moving mouse randomly to X=" moveX " Y=" moveY " with rangeX=" rangeX " and rangeY= " rangeY)

    ; Get current mouse position
    MouseGetPos(&currentX, &currentY)
    
    ; Move the mouse to the new position with a specific speed
    SendMode("Event")
    MouseMove(moveX, moveY, speed)
    ; Debugging: Log the final mouse position after movement
    ;MouseGetPos(&currentX, &currentY)
    ;MsgBox("After Moving: Final Mouse Position X=" currentX ", Y=" currentY)
}

; Function to perform a random UIA click with coordinates
RandomUIAClick(x, y, minDelay := 300, maxDelay := 500) {
    ; Introduce a random delay before interacting
    RandomDelay(minDelay, maxDelay)
    
    ; Debugging: Log the coordinates received for the click
    ;MsgBox("RandomUIAClick using X=" x ", Y=" y)
    
    ; Perform the click at the exact coordinates without further movement
    Click(x, y)
    
    ; Introduce another random delay after the click
    RandomDelay(minDelay, maxDelay)
}

; Global variable to hold the end time for countdown
global endTime, tooltipMessage

; Function to show ToolTip with a visual timer
ShowToolTipWithTimer(message := "", duration := 2000, sleepDuration := "", visualTimer := false) {
    global endTime, tooltipMessage
    tooltipMessage := message
    ToolTip(message)  ; Show the initial ToolTip
    
    ; Initialize the end time based on duration
    endTime := A_TickCount + duration
    
    if visualTimer {
        ; Set a timer to update the tooltip every 1000 ms for smooth countdown
        SetTimer(UpdateToolTipTimer, 1000)
    }

    ; Set the tooltip duration timer to clear it when duration ends
    SetTimer(() => ToolTip(""), -duration)

    ; Handle sleepDuration within the same function to keep it synchronous
    if (sleepDuration != "") {
        Sleep(sleepDuration)
    }
}

; Standalone function to update the tooltip countdown
UpdateToolTipTimer() {
    global endTime, tooltipMessage
    remainingTime := endTime - A_TickCount
    seconds := remainingTime // 1000  ; Convert ms to seconds
    if (remainingTime <= 0) {
        ToolTip("")  ; Clear the tooltip when time is up
        SetTimer(UpdateToolTipTimer, "0")  ; Turn off the update timer
    } else {
        ToolTip(tooltipMessage . "`nTime remaining: " . seconds . " seconds")  ; Update ToolTip with remaining time
    }
}

global secondErrorCaught := false

if FileExist("GetElementCoordinates_Debug_Log.txt") {
    FileDelete("GetElementCoordinates_Debug_Log.txt")
}

; Shared function to retrieve and adjust coordinates for an element or path
GetElementCoordinates(elementOrPath, rootElement := "", alignToElement := false) {
    errorCaught := 0
    try {
        ; Reset variables before each calculation
        x := 0
        y := 0

        ; Check if the parameter is a string, indicating it's a path
        if Type(elementOrPath) = "String" {
            if !IsObject(rootElement) {
                MsgBox("GetElementCoordinates Error: rootElement is required when passing path: " elementOrPath)
                return false
            }
            ; Find the element using the path
            element := rootElement.WaitElementFromPath(elementOrPath, 5000)
            if !element {
                MsgBox("Failed to find element with path: " elementOrPath)
                return false
            }
        } else {
            ; Check if rootElement was mistakenly passed as the element
            if elementOrPath = rootElement {
                MsgBox("GetElementCoordinates Error: rootElement was passed instead of an element or path.")
                return false
            }
            element := elementOrPath  ; Directly use the passed element if it's not a string
        }

        element.ScrollIntoView()
        ;SendInput("{WheelDown}")
        Sleep(250)
        Scroll:
        {
            global xElementCoordinates_GEC := element.Location.x  ; Extract x coordinate
            x := xElementCoordinates_GEC
            ;MsgBox("x=" x)
            global yElementCoordinates_GEC := element.Location.y  ; Extract y coordinate
            y := yElementCoordinates_GEC
            ;MsgBox("y=" y)

            ; Calculate the bottom-right corner of the element by adding width and height
            elementRight := x + element.Location.w  ; Right boundary (x + width)
            ;MsgBox("elementRight=" elementRight)
            elementBottom := y + element.Location.h  ; Bottom boundary (y + height)
            ;MsgBox("elementBottom=" elementBottom)

            ; Get the total number of monitors
            MonitorCount := MonitorGetCount()

            ; Identify which monitor the coordinates belong to
            monitorNumber := 0
            Loop MonitorCount {
                MonitorGet(A_Index, &Left, &Top, &Right, &Bottom)
                if (xElementCoordinates_GEC >= Left && xElementCoordinates_GEC <= Right && yElementCoordinates_GEC >= Top && yElementCoordinates_GEC <= Bottom) {
                    monitorNumber := A_Index
                    ;MsgBox("Coordinates found on Monitor " monitorNumber)
                    break
                }
            }
            xElementCoordinates_GEC := 0
            yElementCoordinates_GEC := 0

            ; If no monitor is found, return an error
            if (monitorNumber = 0) {
                MsgBox("Failed to find a monitor containing the coordinates.")
                return false
            }

            if alignToElement {
                ; Get the current mouse position
                MouseGetPos(&currentX, &currentY)

                averageX := (elementRight - x) / 2 + x
                
                MouseMove(averageX, currentY, speed := 10)
            }

            ; Get the monitor dimensions where the element is located
            MonitorGetWorkArea(monitorNumber, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom) ; 1 refers to the primary monitor

            ; Compare element's position with monitor/screen boundaries
            if (x >= monitorLeft && elementRight <= monitorRight && y >= monitorTop && elementBottom <= monitorBottom) {
                ;MsgBox("Element is fully visible in the viewport.")
            } else {
                ;MsgBox("Element is not fully visible in the viewport.")
                Click("WheelDown")
                Sleep(250)
                goto Scroll
            }
        }

        ; Optionally, log the element's boundaries and monitor boundaries
        ;MsgBox("Element boundaries: Left=" x ", Top=" y ", Right=" elementRight ", Bottom=" elementBottom)
        ;MsgBox("Monitor boundaries: Left=" monitorLeft ", Top=" monitorTop ", Right=" monitorRight ", Bottom=" monitorBottom)

        ; No need to adjust coordinates, just return them as-is for the correct monitor
        return {x: x, y: y, elementRight: elementRight, elementBottom: elementBottom, monitorNumber: monitorNumber}
    } catch as e {
        errorCaught++
        errorMessage := "Error in GetElementCoordinates: " e.Message
        errorLine := "Line: " e.Line
        errorExtra := "Extra Info: " e.Extra
        errorFile := "File: " e.File
        errorWhat := "Error Context: " e.What
        
        ; Display detailed error information
        ShowToolTipWithTimer(errorMessage "`n" errorLine "`n" errorExtra "`n" errorFile "`n" errorWhat)
        
        ; Log the detailed error information
        FileAppend(errorMessage "`n" errorLine "`n" errorExtra "`n" errorFile "`n" errorWhat "`n", "GetElementCoordinates_Debug_Log.txt")
    }

    switch [errorCaught] {
        case 1:
        Sleep(5000) GetElementCoordinates(elementOrPath, rootElement)
        case 2:
        return false
            
    }
}

if FileExist("ClickElementByPath_debug_log.txt") {
    FileDelete("ClickElementByPath_debug_log.txt")
}

; Function to click an element by path or directly by element
ClickElementByPath(pathOrElement, rootElement := "", value := "", timeout := 2000, alignToElement := false) {

    ; Determine if pathOrElement is already an element or needs to be resolved from a path
    if IsObject(pathOrElement) {
        element := pathOrElement  ; It's already an element
    } else {
        ; If pathOrElement is not an object, assume it's a path and resolve it
        if !IsObject(rootElement) {
            ShowToolTipWithTimer("ClickElementByPath Error: rootElement is required when passing path: " element)
            FileAppend("ClickElementByPath Error: rootElement is required when passing path: " element, "ClickElementByPath_debug_log.txt")
            return false
        }
        element := rootElement.WaitElementFromPath(pathOrElement, timeout)
        if !element {
            ShowToolTipWithTimer("Element with path " pathOrElement " not found within timeout.")
            FileAppend("Element with path " pathOrElement " not found within timeout.", "ClickElementByPath_debug_log.txt")
            return false
        }
    }

    ; Activate the Chrome window before any further action
    ActivateChromeWindow()
    
    ; Retrieve and adjust the element's coordinates (adjust with ScrollIntoView and Send("{WheelDown}"))
    coords := GetElementCoordinates(element, rootElement , alignToElement)
    if !coords {
        ShowToolTipWithTimer("ClickElementByPath failed to get element coordinates for element: " element.Name)
        FileAppend("ClickElementByPath failed to get element coordinates for element: " element.Name "`n`n", "ClickElementByPath_debug_log.txt")
        return false
    }
   
    ; Perform random mouse movement before interacting
    RandomMouseMove(coords.x, coords.y, coords.elementRight, coords.elementBottom,,, speed := 10)
    
    ; Adding a MsgBox to see where the mouse is after moving
    MouseGetPos(&currentX, &currentY)

    ; If a value is provided, set it
    if (value != "") {
        element.Value := value
    } else {
        ; Pass the exact coordinates to RandomUIAClick
        RandomUIAClick(currentX, currentY)
    }
    return true
}