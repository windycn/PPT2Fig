on runPPT2Fig(parameterText)
    set oldDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to tab
    set parts to text items of parameterText
    set AppleScript's text item delimiters to oldDelimiters
    if (count of parts) < 2 then return "ERROR: missing arguments"
    set modeName to item 1 of parts
    set firstPath to item 2 of parts
    set helperPath to (POSIX path of (path to home folder)) & "Library/Application Support/PPT2Fig/ppt2fig"
    set commandLine to "PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin " & quoted form of helperPath & " " & quoted form of modeName & " " & quoted form of firstPath
    if (count of parts) > 2 then
        set secondPath to item 3 of parts
        if secondPath is not "" then set commandLine to commandLine & " " & quoted form of secondPath
    end if
    try
        do shell script commandLine
        return "OK"
    on error errorText
        return "ERROR: " & errorText
    end try
end runPPT2Fig
