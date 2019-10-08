#include <_FileFindEx.au3>
#include <excel.au3>

Func _FileGetShortFilename($sPath)
	Local $sShortFilename,$sCurrentFilename=StringMid($sPath,StringInStr($sPath,'\',1,-1)+1)
	$sPath=FileGetShortName($sPath)
	$sShortFilename=StringMid($sPath,StringInStr($sPath,'\',1,-1)+1)
	If $sCurrentFilename=$sShortFilename Then Return ""
	Return $sShortFilename
EndFunc

Global $aFileFindArray,$iTotalCount,$iFolderCount,$iTimer,$sFilename,$sAttrib,$hFindFileHandle
Global $sFolder= @ScriptDir & '\' ;(@WorkingDir);@SystemDir&'\'	; @UserProfileDir&'\'	; (the latter results in some REPARSE Points)
Global $sSearchWildCard="*.*"
Global $s_FileFindExStats="",$sFileFindFileStats="",$bDuplicateResults=True,$bShowStats=True,$iIsFolder

;Info Message==============================================================
MsgBox( "" ,"About Me","This Tools Made By Md. Siduzzaman Sagor" & @CRLF & "Contact for any issue: 01731284258; Thanks" & @CRLF & "" & @CRLF & "To start press 'OK'")

; If file Exist============================================================
If FileExists(@WorkingDir & "\output.txt")  then
   FileRecycle(@WorkingDir & "\output.txt")
EndIf

;Function Execute==========================================================
FFEXTest()

; -------------- _FileFindEx Test ------------

Func FFEXTest()

	$iTotalCount=0
	$iFolderCount=0
	$iTimer=TimerInit()
	$aFileFindArray=_FileFindExFirstFile($sFolder & $sSearchWildCard)
	If $aFileFindArray=-1 Then Exit

	Do
		If BitAND($aFileFindArray[2],16) Then
			If $aFileFindArray[0]='.' Or $aFileFindArray[0]='..' Then ContinueLoop
			; Increase found-folder count
			$iFolderCount+=1
		EndIf
		; Total file+folder count
		$iTotalCount+=1
		; 'Long' Filename: $aFileFindArray[0]
		; Attributes: $aFileFindArray[2]

		If $bShowStats Then
			$s_FileFindExStats &= "" & $aFileFindArray[0] & @CRLF

		ElseIf $bDuplicateResults Then

			_FileFindExTimeConvert($aFileFindArray,1,1)
			; Last-Access Time
			_FileFindExTimeConvert($aFileFindArray,2,1)
			; Last-Write Time
			_FileFindExTimeConvert($aFileFindArray,0,1)
		EndIf
	Until Not _FileFindExNextFile($aFileFindArray)
	_FileFindExClose($aFileFindArray)

EndFunc

; ------------ FileFindTest --------------

;Export to text=====================================================
Global $re = $s_FileFindExStats ;Test('example string')
ConsoleWrite($re & @CRLF)
FileWrite(@WorkingDir & "\output.txt", $re)

ReplaceExt()

; Open by Excel ====================================================
Local $sTxt = @WorkingDir & "\Output.txt"
Local $oExcel = ObjCreate("Excel.Application")
If IsObj($oExcel) Then
    $oExcel.Visible = 1
    $oExcel.Workbooks.OpenText($sTxt, 2, 1, 1, -4142, False, False, False, False, False, True, "|")
 EndIf

; Save Books and Close
;_Excel_BookSaveAs($oBook, $FilePath)
_Excel_Close($oExcel, True)

clearTemp()