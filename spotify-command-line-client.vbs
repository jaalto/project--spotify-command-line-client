' -*- comment-start: "'" -*-
'
'   spotify-command-line-client.vbs -- Send URI to a desktop spotify client
'
'   Copyright
'
'       Copyright (C) 2017-2019 Jari Aalto <jari.aalto@cante.net>
'
'   License
'
'       This program is free software; you can redistribute it and/or
'       modify it under the terms of the GNU General Public License as
'       published by the Free Software Foundation; either version 2 of
'       the License, or (at your option) any later version.
'
'       This program is distributed in the hope that it will be useful, but
'       WITHOUT ANY WARRANTY; without even the implied warranty of
'       MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'       General Public License for more details.
'
'       Visit <http://www.gnu.org/copyleft/gpl.html>
'
'   USAGE
'
'	cscript /nologo spotify-command-line-client.vbs --help
'
'   NOTES
'
'       Expected Spotify client location:
'       %APPDATA%\Roaming\Spotify\spotify.exe
'
'   CODING STYLE
'
'	if, while 	Standard indentifiers in lowercase
'	varName		Variables in camelCase
'	FunctionName	Functions/Subs in capital first letter
'
'       NOT USED: "Microsoft All Caps Style" as it violates all
'       the known typographical legibility rules. Also impossible
'	hard to type while coding.
'
'   VB SCRIPT DOCUMENTATION FOR DEVELOPERS
'
'       https://docs.microsoft.com/en-us/dotnet/visual-basic/index
'       https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference
'       https://msdn.microsoft.com/library/ms950396.aspx
'       https://technet.microsoft.com/fi-fi/scriptcenter/dd772284.aspx
'       https://www.regular-expressions.info/vbscript.html
'       site:ss64.com/vb/
'
'       wscript.Quit


' See --help and command line options to alter values

VERBOSE_MODE = false
DEBUG_SECONDS = 0
SEARCH_STRING = ""

INTERPRETER = lcase(mid(wscript.FullName, instrrev(wscript.FullName,"\")+1))
set SHELL = wscript.CreateObject("wscript.Shell")

sub Version
    ' Display version etc. infromation

    wscript.echo _
      "Copyright (C) 2017-2019 Jari Aalto <jari.aalto@cante.net>" & vbnewline _
    & "Homepage: http://githum.com/jaalto/project--spotify-vb-client" & vbnewline _
    & "License: GPL-2+ See <http://www.gnu.org/copyleft/gpl.html>"

    wscript.Quit
end sub

sub Help
    ' Display help

    wscript.echo _
      "SYNOPSIS" & vbnewline _
    & "    cscript /nologo <file>.vbs [options] <spotify URI>" & vbnewline _
    & vbnewline _
    & "OPTIONS" & vbnewline _
    & "    -a, --args     Convert args into search. See DESCRIPTION" & vbnewline _
    & "    -d, --debug    Turn on debug delay. Useful when called as proxy:" & vbnewline _
    & "                   DESKTOP APP > THIS PROGRAM (proxy) > SPOTIFY." & vbnewline _
    & "                   Turning on debug hold the (proxy) shell window open" & vbnewline _
    & "                   long anough to see possible error message." & vbnewline _
    & "    -h, --help     Display brief help and quit (this screen)" & vbnewline _
    & "    -v, --verbose  Turn on verbose mode." & vbnewline _
    & "    -V, --version  Display Copyright, License etc. and quit" & vbnewline _
    & vbnewline _
    & "DESCRIPTION" & vbnewline _
    & "    Send URI to the desktop spotify client. URI forms:" & vbnewline _
    & vbnewline _
    & "        spotify:{album|playlist|search|track|user:UID:playlist:}:<data>" & vbnewline _
    & vbnewline _
    & "    In case of ""search"", multiple words must be separated by plus signs:" & vbnewline _
    & vbnewline _
    & "        spotify:search:word+word..." & vbnewline _
    & vbnewline _
    & "    In auto mode, non-option arguments are converted into search:" & vbnewline _
    & vbnewline _
    & "        //  cscript PROGRAM.vbs -a ARG1 ARG2 ARG3 ..." & vbnewline _
    & "        spotify:search:arg1+arg2+arg3..." & vbnewline _
    & vbnewline _
    & "        //  Same, even if using quotes: ccript PROGRAM.vbs -a ""ARG1 ARG2"" ARG3 ..." & vbnewline _
    & "        spotify:search:arg1+arg2+arg3..."

    wscript.Quit
end sub

sub Warn(str)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stderr = fso.GetStandardStream(2)
    stderr.WriteLine(str)
end sub

function isDebug
    ' Return true if debug is active

    if DEBUG_SECONDS > 0 then
        isDebug = true
    else
        isDebug = false
    end if
end function

sub Sleep
    ' Sleep if debug is active

    if isDebug() then
       wscript.sleep DEBUG_SECONDS * 1000
     end if
end sub

sub Quit(msg)
    ' Display message to stdout and exit

    wscript.Echo msg
    Sleep
    wscript.Quit
end sub

sub Die(str)
    ' Display message to stderr and exit

    Warn str
    Sleep
    wscript.Quit
end sub

sub Verbose(msg)
    if VERBOSE_MODE then
        wscript.echo msg
    end if
end sub

function isProcess(exe)
    ' Return true if program EXE is running

    Set procshell = wscript.CreateObject("WScript.Shell")
    Set proclist = GetObject("Winmgmts:").ExecQuery ("Select * from Win32_Process")

    retval = false

    for each proc in proclist
        If lcase(proc.name) = lcase(exe) then
            retval = true
            exit for
        end if
    next

    isProcess = retval
end function

sub Spotify(path, uri)
    ' Call PATH (spotify.exe) and send URI

    sleepTime = 1500
    running = True

    if isProcess("spotify.exe") then
	Verbose "Spotify already running"
    else
        running = False
        Verbose "NOTE: Spotify not running. Sending an URI may take a while..."
    end if

    Verbose "Sending URI to Spotify: " & uri

    CreateObject("wscript.Shell").Run(uri)

    if not running then
       ' conservative time to wait for desktop app to launch properly
        sleepTime = 5000
    end if

    ' Don't play track
    ' Send SPC to pause the track. Value is in milliseconds
    ' https://ss64.com/vb/sendkeys.html

    if instr(uri, "spotify:track:") then
	' Wait frame to raise to receive keys
        wscript.sleep sleepTime
        SHELL.SendKeys " "

    elseif instr(uri, "spotify:search:") then
	'  Put search string to the standard serch field in Spotify
	'  so that it can be modified

	uri = replace(uri, "spotify:search:", "")
	uri = strConvertUnplus(uri)

	Verbose "Sending Control-L to set focus on search field for: " & uri

        wscript.sleep sleepTime
        SHELL.SendKeys "^L"
        SHELL.SendKeys uri
    end if
end sub

function isUri(str)
    ' Check valid Spotify URI

    set re = new RegExp
    re.IgnoreCase = true
    re.Pattern = "spotify:(track|album|user|search):"

    ' Set matches = re.Execute(str)
    ' count = matches.Count

    isUri = re.Test(str)
end function

function strConvertSpace(str, replacement)
    ' Return string with spaces replaced with REPLACEMENT

    str = trim(str)

    with new regexp
        .Pattern = "\s+"
	.Global = true
        strConvertSpace = .Replace(str, replacement)
    end with
end function

function strConvertUnplus(str)
    ' Return string with plus(+) converted to spaces

    str = trim(str)

    with new regexp
        .Pattern = "\+"
	.Global = true
        strConvertUnplus = .Replace(str, " ")
    end with
end function


function strStripPunctuation(str)
    ' Return string where punctuations are replaced with spaces

    with new regexp
        .Pattern = "[][<>|;:_+=*,.?!'""`#%&/()~^-]"
        .Global = true
        strStripPunctuation = .Replace(str, " ")
    end with
end function

function uriCanonicalize(str)
    ' Return string suitable for "spotify:search:<data>" protocol

    original = str
    str = strConvertSpace(str, "+")

    if str <> original then
        Verbose "Canonicalizing URI without whitespace"
    end if

    uriCanonicalize = str
end function

function inPath(file)
    ' Return absolute path if FILE was found in PATH

    paths = SHELL.ExpandEnvironmentStrings("%PATH%")

    for each item in split(paths, ";")
	path = item & "\" & file

	Set f = CreateObject("Scripting.FileSystemObject")

	if (f.FileExists(path)) then
	    Verbose "Found in PATH: " & path
	    inPath = path
	    exit for
	end if
    next

end function

function SpotifyAppdataPath
    ' Return path for Spotify desktop client under APPDATA

    appdata = SHELL.ExpandEnvironmentStrings("%APPDATA%")
    path = appdata & "\Spotify\spotify.exe"

    Set f = CreateObject("Scripting.FileSystemObject")

    if not f.FileExists(path) then
        Die "ERROR: no such path " & path
    else
	Verbose "Using spotify: " & path
    end if

    SpotifyAppdataPath = path
end function

sub Main
    ' Main program to handle command line argments

    argc = wscript.Arguments.Count

    if argc = 0 then
        Die "ERROR: missing spotify URI. See --h, --help"
    end if

    set reOption = new RegExp
    reOption.IgnoreCase = true
    reOption.Pattern = "^-"

    autoMode = false
    pos = 0
    uriPos = 0
    uri = ""

    for each arg in wscript.Arguments
        pos = pos + 1

        if arg = "-a" or arg = "--args" then
            autoMode = true
        elseif arg = "-d" or arg = "--debug" then
           DEBUG_SECONDS = 5
	   VERBOSE_MODE = true
        elseif arg = "-h" or arg = "--help" then
           Help
        elseif arg = "-V" or arg = "--version" then
           Help
        elseif arg = "-v" or arg = "--verbose" then
           VERBOSE_MODE = true
           Verbose "Verbose enabled"
        elseif reOption.Test(arg) then
           Warn "WARN: unknown option " & arg
        else
           if autoMode then
                uriPos = pos

                if uri = "" then
                    uri = arg
                else
                    uri = uri & " " & arg
                end if
            else
                uriPos = pos
                uri = arg
                exit for
            end if
        end if
    next

    Verbose "Command line URI data: [" & uri & "]"

    if uri = "" then
       Die "ERROR: missing spotify URI. See -h, --help"
    end if

    if uriPos < argc then
        Warn "WARN: Multiple arguments ignored without --args. Using arg: " & uri
    end if

    if autoMode then
	original = uri
	uri = strStripPunctuation(uri)

	if str <> original then
	    Verbose "Canonicalizing URI without punctuation"
	end if

	uri = "spotify:search:" & uri
    end if

    uri = uriCanonicalize(uri)

    path = inPath("spotify.exe")

    if path = "" then
        path = SpotifyAppdataPath()

        if path = "" then
	    Die "ERROR: Can't locate spotify.exe"
	end if
    end if

    Verbose "Using URI: " & uri
    Sleep
    Spotify path, uri
end sub

Main

' End of file
