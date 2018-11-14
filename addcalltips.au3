; AddCallTips.au3

;automates adding User Call Tips to SciTE. See SciTE help for User Call Tips.
  Const $version = "V1.5"
;version changes
;1.5 Allowed for having functions in UDF which must not be given Call Tips. Useful for internal functions.
;1.4 rewote functions UpdateCallTips and replaceUCT
;1.31 Removed \V at start of search patter for StringREgExpr/Replace
;1.3 Improved matching function names so if parameters have been changed update still works etc.
;1.2 Added warnings for invalid files, too few comment chars and no Key Word for comments.
;   Rearranged things a bit to improve readabilty. Simplified some things.
;1.1 change default folder to last folder used for fileopendialog {Bowmore}
;   made folder for backups and added numbered backups

;next line only used when developing the script
;#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6

  #include <EditConstants.au3>
  #include <GUIConstantsEx.au3>
  #include <WindowsConstants.au3>

  #Region ### START Koda GUI section ###
  $Form2 = GUICreate("Add User Call Tips", 380, 443, 303, 219)
  $Label1 = GUICtrlCreateLabel("UDF whose functions are to be  added", 12, 8, 187, 17)
  $IpUDFpath = GUICtrlCreateInput("", 12, 27, 351, 21)
  $BtnBrowse = GUICtrlCreateButton("&Browse", 206, 5, 50, 20, 0)
  $BtnApply = GUICtrlCreateButton("&Apply the above changes to Scite UserCallTips", 12, 301, 353, 25, 0)
  $Group1 = GUICtrlCreateGroup("Options", 8, 175, 362, 113)
  $ChkCommentLimit = GUICtrlCreateCheckbox("&Limit maximum comment length", 200, 194, 163, 17)
  GUICtrlSetState(-1, $GUI_CHECKED)
  $RadioReplaceAll = GUICtrlCreateRadio("&Replace existing CallTips", 12, 194, 137, 17)
  $RadioReplaceWarn = GUICtrlCreateRadio("&Warn before replacing Call Tips", 12, 217, 169, 17)
  $RadioReplaceNone = GUICtrlCreateRadio("Do &not replace existing Call Tips", 11, 240, 170, 17)
  GUICtrlSetState(-1, $GUI_CHECKED)
  $InputMaxChar = GUICtrlCreateInput("80", 238, 215, 32, 21)
  $Label4 = GUICtrlCreateLabel("characters", 276, 218, 54, 17)
  $ChkShowSummary = GUICtrlCreateCheckbox("&Show result summary", 200, 265, 145, 17)
  GUICtrlSetState(-1, $GUI_CHECKED)
  GUICtrlSetState(-1, $GUI_HIDE)
  $Label5 = GUICtrlCreateLabel("", 189, 184, 1, 99)
  GUICtrlSetBkColor(-1, 0xD4D0C8)
  $Label7 = GUICtrlCreateLabel("to", 223, 217, 13, 17)
  $RadioRemoveAll = GUICtrlCreateRadio("Remo&ve all Call Tips for this UDF", 12, 265, 174, 17)
  GUICtrlCreateGroup("", -99, -99, 1, 1)
  $EdResults = GUICtrlCreateEdit("", 15, 352, 347, 81, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_HSCROLL, $WS_VSCROLL))
  $LabResSummary = GUICtrlCreateLabel("Result summary", 16, 335, 78, 17)
  $Group2 = GUICtrlCreateGroup("Key words in UDF (must be in a comment before the function.)", 8, 64, 362, 97)
  $Key = GUICtrlCreateLabel("the text after this Keyword", 18, 84, 126, 17)
  $Label6 = GUICtrlCreateLabel("to exclude the Function", 25, 117, 115, 17)
  $Label8 = GUICtrlCreateLabel(" from UserCallTips", 28, 132, 89, 17)
  $IpKeyDescription = GUICtrlCreateInput("Description:", 155, 86, 201, 21)
  $IpKeyExclude = GUICtrlCreateInput("Internal Use", 155, 121, 201, 21)
  $Label2 = GUICtrlCreateLabel("is the comment", 28, 97, 75, 17)
  GUICtrlCreateGroup("", -99, -99, 1, 1)
  GUISetState(@SW_SHOW)
  #EndRegion ### END Koda GUI section ###

  WinSetTitle($Form2, "", "Add User Call Tips   " & $version)
  Const $ReplaceNone = 0
  Const $ReplaceAll = 1
  Const $ReplaceWarn = 2
  Const $RemoveAll = -1

  Global $iniFile = @ScriptDir & "\UserCallTips.ini"
  Global $SavedKWDesc = IniRead($iniFile, "Keywords", "UserCallTip", "Description:"), $KWDesc
  Global $SavedKWExclude = IniRead($iniFile, "Keywords", "Exclude", "Internal Use"), $KWExclude
  Global $lastfile = IniRead($iniFile, "Keywords", "LastFile", "")
  Global $lastCTCommentLimit = IniRead($iniFile, "Keywords", "MaxComment", 100)
  Global $lastFolder = StringLeft($lastfile, StringInStr($lastfile, '\', 0, -1))

  GUICtrlSetData($IpUDFpath, $lastfile)
  GUICtrlSetData($IpKeyDescription, $SavedKWDesc)
  GUICtrlSetData($InputMaxChar, $lastCTCommentLimit)
  If StringInStr($lastfile, '\') Then
      $lastFolder = StringLeft($lastfile, StringInStr($lastfile, '\', 0, -1))
  EndIf

  Global $nMsg, $file
  While 1
      $nMsg = GUIGetMsg()
      Switch $nMsg
          Case $GUI_EVENT_CLOSE
              Exit
          Case $ChkCommentLimit
              HideOrShowCommentLimit(GUICtrlRead($ChkCommentLimit) = $GUI_UNCHECKED)
          Case $BtnBrowse
              $file = FileOpenDialog("Select the UDF", $lastFolder, "AutoIt File (*.au3)", 3)
              GUICtrlSetData($IpUDFpath, $file)
          Case $BtnApply
              If Not ApplyChanges() Then GUICtrlSetData($EdResults, "No action taken")

      EndSwitch
  WEnd

;ApplyChanges
;returns true if changes made otherwise it returns false
  Func ApplyChanges()
      Local $CTresult, $UDFfile, $CallTipCommentLimit, $opt

      If GUICtrlRead($IpKeyDescription) = '' Then
          If WarningA(262144 + 4, "WARNING", "You have no Key Word for comments." & @CRLF & " No Comments will be added." & @CRLF _
                   & "Do you wish to proceed?", $IpKeyDescription) <> 6 Then Return False
      EndIf




      $UDFfile = GUICtrlRead($IpUDFpath)
      If $UDFfile = '' Or Not FileExists($UDFfile) Then
          If WarningA(262144, "ERROR", "You must enter a valid AU3 script!", $IpUDFpath) <> 6 Then Return False
      EndIf

      GUICtrlSetData($EdResults, "")
      $KWDesc = GUICtrlRead($IpKeyDescription)
      $KWExclude = GUICtrlRead($IpKeyExclude)

      If GUICtrlRead($RadioReplaceAll) = $GUI_CHECKED Then
          $opt = $ReplaceAll
      ElseIf GUICtrlRead($RadioReplaceWarn) = $GUI_CHECKED Then
          $opt = $ReplaceWarn
      ElseIf GUICtrlRead($RadioRemoveAll) = $GUI_CHECKED Then
          $opt = $RemoveAll
      Else
          $opt = $ReplaceNone;do not replace any
      EndIf

      If $opt <> $RemoveAll Then;a new Call Tip could be written
          If GUICtrlRead($ChkCommentLimit) = $GUI_CHECKED Then
              $CallTipCommentLimit = GUICtrlRead($InputMaxChar)
              If $CallTipCommentLimit < 40 Then
                  If WarningA(262144 + 4, "WARNING", "You have set fewer than 40 characters for comments." & @CRLF _
                           & "Do you wish to proceed?", $InputMaxChar) <> 6 Then Return False
              EndIf
          Else
              $CallTipCommentLimit = 0;0 means no maximum
          EndIf
      EndIf

      If $opt = $RemoveAll Then
          If MsgBox(262144 + 4, "Confirm", "Do you want to remove all the Call Tips for " & _
                  @CRLF & $UDFfile & "?") <> 6 Then Return False
      EndIf

      $CTresult = UpDateCallTips($UDFfile, $KWDesc, $opt, $CallTipCommentLimit)
      GUICtrlSetData($EdResults, $CTresult)
;MsgBox(262144, "Reult", $CTresult)
      If $SavedKWDesc <> $KWDesc Then
          If MsgBox(262144 + 4, "Key word for Description changed to '" & $KWDesc & "'", "Do yu want to save the new keyword for next time?") = 6 Then
              IniWrite($iniFile, "Keywords", "UserCallTip", $KWDesc)
          EndIf
      EndIf

      If $SavedKWExclude <> $KWExclude Then
          If MsgBox(262144 + 4, "Key word for Exclusion has changed to '" & $KWExclude & "'", "Do yu want to save the new keyword for next time?") = 6 Then
              IniWrite($iniFile, "Keywords", "Exclude", $KWExclude)
          EndIf
      EndIf

      IniWrite($iniFile, "Keywords", "LastFile", GUICtrlRead($IpUDFpath))
      If $CallTipCommentLimit <> $lastCTCommentLimit Then
          IniWrite($iniFile, "Keywords", "MaxComment", $CallTipCommentLimit)
          $lastCTCommentLimit = $CallTipCommentLimit
      EndIf

      Return True
  EndFunc;==>ApplyChanges


  Func HideOrShowCommentLimit($dohide)
      Local $hidestate = $GUI_SHOW
      If $dohide Then $hidestate = $GUI_HIDE

      GUICtrlSetState($Label4, $hidestate)
      GUICtrlSetState($Label7, $hidestate)
      GUICtrlSetState($InputMaxChar, $hidestate)
  EndFunc;==>HideOrShowCommentLimit


  Func UpDateCallTips($udf, $CallTipKeyword, $mode, $maxComment)
      Local $sCT, $fudf, $CallTipComment = '', $Func = '', $s2
      Local $sect, $line, $fw, $CTadded = 0, $CTdoneBefore = 0, $CTreplaced = 0, $CTremoved = 0
      Local $AutoItProdexePath = "C:\Users\" & @UserName & "\AppData\Local\AutoIt v3\SciTE"
      Local $UCTFile = $AutoItProdexePath & "\au3.user.calltips.api"
      Local $result, $backupNum, $backPath, $fnName, $ArraysCT
      Local $Usemode

      $sCT = FileRead($UCTFile)
      $ArraysCT = StringSplit(StringReplace($sCT, @CR, ''), @LF)
      $fudf = FileOpen($udf, 0)

      While 1
          $line = FileReadLine($fudf)
          If @error <> 0 Then ExitLoop


          $line = StringStripWS($line, 3)
          If StringLeft($line, 1) = ';' Then;it's a comment line so see if it's got the Call Tip in it
              $sect = StringLeft(StringStripWS(StringRight($line, StringLen($line) - 1), 1), StringLen($CallTipKeyword))
              If $sect = $CallTipKeyword Then;yes found Call Tip key word so strip out the Call Tip
                  $CallTipComment = StringRight($line, StringLen($line) - StringInStr($line, $CallTipKeyword) - StringLen($CallTipKeyword))
                  $CallTipComment = StringStripWS($CallTipComment, 3)
                  If $maxComment > 0 Then $CallTipComment = StringLeft($CallTipComment, $maxComment)
              EndIf
              If StringInStr($line, $KWExclude)  And $mode <> $ReplaceNone Then
                  $Usemode = $RemoveAll
              Else
                  $Usemode = $mode
              EndIf

          EndIf

;look for functions
          If StringLeft($line, 5) = 'Func ' Then
              $Func = getfunc($line);get the function name with parameter details
;$sline includes the parameters, but these could change so we must only look for
;     the fn name preceded by any number of spaces and followed by any number of spaces then '('
              $fnName = StringStripWS(StringLeft($Func, StringInStr($Func, '(') - 1), 3)

              replaceUCT($fnName, $ArraysCT, $Func & $CallTipComment, $Usemode, $CTadded, $CTdoneBefore, $CTreplaced, $CTremoved)

          EndIf

          If StringLeft($line, 7) = 'EndFunc' Then $CallTipComment = ''


      WEnd
      FileClose($fudf)

      If $CTreplaced + $CTremoved + $CTadded > 0 Then;if a change made
;backup calltips file before changing it
          $backupNum = 1
          $backPath = $AutoItProdexePath & "\BAKs\"
          If Not FileExists($backPath) Then
              DirCreate($backPath)
          EndIf

          While FileExists($backPath & "BAK" & $backupNum & "au3.user.calltips.api")
              $backupNum += 1
          WEnd
          FileCopy($AutoItProdexePath & "\au3.user.calltips.api", $backPath & "BAK" & $backupNum & "au3.user.calltips.api")
          $fw = FileOpen($UCTFile, 2)
          For $n = 1 To UBound($ArraysCT) - 1
              If StringStripWS($ArraysCT[$n], 3) <> '' Then FileWriteLine($fw, $ArraysCT[$n])
          Next

          FileClose($fw)
      EndIf
      $result &= "Call Tip changes" & @CRLF
      If $CTadded > 0 Then

          $result &= $CTadded & ' added.' & @CRLF
      Else
          $result = "No User Call Tips were added!" & @CRLF
      EndIf
      If $CTdoneBefore > 0 Then
          $s2 = ' function was'
          If $CTdoneBefore > 1 Then $s2 = ' functions were'
          $result &= $CTdoneBefore & $s2 & ' already included.' & @CRLF
      EndIf
      If $CTreplaced > 0 Then
          $s2 = ' function was'
          If $CTreplaced > 1 Then $s2 = ' functions were'
          $result &= $CTreplaced & $s2 & ' replaced.' & @CRLF
      EndIf
      $result &= @CRLF
      If $CTremoved > 0 Then
          $result &= $CTremoved & " removed" & @CRLF
      Else
          $result &= "No call tips removed"
      EndIf
      Return $result

  EndFunc;==>UpDateCallTips

;Returns the function name and parameters from the line without the preceding "Func" text
  Func getfunc($sf)
      Local $sm, $spos, $lBracket = 0, $rBracket = 0, $endfunc = 0
      $sf = StringReplace($sf, "Func ", '')
      For $spos = 1 To StringLen($sf)
          $sm = StringMid($sf, $spos, 1)
          Switch $sm
              Case '('
                  $lBracket += 1
              Case ')'
                  $rBracket += 1
          EndSwitch

          If $lBracket > 0 And $lBracket = $rBracket Then
              $endfunc = $spos
              ExitLoop
          EndIf
      Next
      If $endfunc > 0 Then Return StringLeft($sf, $endfunc)
      SetError(-1)
      Return ''

  EndFunc;==>getfunc

;Replace the line beginning with $sline in $sText with $NewUCT
  Func replaceUCT($sline, ByRef $sText, $NewUCT, $repMode, ByRef $PCTadded, ByRef $PCTdoneBefore, ByRef $PCTreplaced, ByRef $PCTremoved)
      Local $n
      For $n = 1 To UBound($sText) - 1
          If StringRegExp($sText[$n], "(?i)\h*\Q" & $sline & "\E\h*\(") Then
              $PCTdoneBefore += 1
              Switch $repMode
                  Case $RemoveAll
                      $sText[$n] = '';$NewUCT
                      $PCTremoved += 1
                  Case $ReplaceNone
;already included so do nthing
                  Case $ReplaceAll
                      $sText[$n] = $NewUCT
                      $PCTreplaced += 1
                  Case $ReplaceWarn
                      If MsgBox(262144 + 4, "Confirmation", "Replace the call tip for function" & @CRLF & $sline & ' ?') <> 6 Then Return
                      $sText[$n] = $NewUCT
                      $PCTreplaced += 1
              EndSwitch
              Return;found and actioned so job done

          EndIf

      Next
;haven't found the function so what do we do?
      Switch $repMode
          Case $RemoveAll
              Return
          Case $ReplaceAll, $ReplaceNone
              $PCTadded += 1
              $n = UBound($sText)
              ReDim $sText[$n + 1]
              $sText[$n] = $NewUCT
      EndSwitch

  EndFunc;==>replaceUCT

  Func WarningA($Wtype, $WTitle, $WMessage, $WID, $thenFocus = True)
      Local $result
      GUICtrlSetBkColor($WID, 0xff3300)
      $result = MsgBox($Wtype, $WTitle, $WMessage)
      Sleep(500)
      GUICtrlSetBkColor($WID, 0xffffff)
      If $thenFocus Then GUICtrlSetState($WID, $GUI_FOCUS)
      Return $result
  EndFunc;==>WarningA