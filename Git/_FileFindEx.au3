#RequireAdmin
#include-once
;~ #include <_WinTimeFunctions.au3>	; _WinTime_UTCFileTimeFormatBasic() alternative
; ===============================================================================================================================
; <_FileFindEx.au3>			; (formerly '_WinAPI_FileFind.au3')
;
; Direct Windows API FindFirst/Next File DLL calls for the purpose of getting additional information
;	not found in FileFindFirst/NextFile() calls (these calls would need to be added to get the same
;		information: (FileGetAttrib(),FileGetTime(),FileGetSize()) -> all adding additional overhead
;
; Version: 2011.05.24
;
; NOTE: While MSDN notes that the file information *may* not be accurate, the information reflects what one would see
;	in Explorer.  The reason MSDN warns programmers is because of the NTFS file system update delays (around 1 hour).
;	So, to be TOTALLY accurate, one can use the API call 'GetFileInformationByHandle' to get File Attributes.
;	Note this also requires opening the file using _WinAPI_CreateFile($sFile,2,2,6) and calling _WinAPI_CloseHandle() after.
;	In most scenarios, this won't cause problems, but its important to know about the NTFS delay.
;	See <_FileGetSizeOnDisk.au3> for '_FileGetInfoByOpening' and other undocumented func's for getting MORE attributes.
;
; REPARSE POINTS attributes note [FILE_ATTRIBUTE_REPARSE_POINT = 0x400]:
;	The 'dwReserved0' member of the WIN32_FIND_DATA structure will contain relevant data *only* if file attrib 0x400 is set:
;	 Possible values:
;		IO_REPARSE_TAG_DFS			(0x8000000A)
;		IO_REPARSE_TAG_DFSR			(0x80000012)
;		IO_REPARSE_TAG_HSM			(0xC0000004)
;		IO_REPARSE_TAG_HSM2			(0x80000006)
;		IO_REPARSE_TAG_MOUNT_POINT	(0xA0000003)
;		IO_REPARSE_TAG_SIS			(0x80000007)
;		IO_REPARSE_TAG_SYMLINK		(0xA000000C)
;	See (on MSDN): 'Reparse Point Tags (Windows)': http://msdn.microsoft.com/en-us/library/aa365511%28v=VS.85%29.aspx
;	 and 'WIN32_FIND_DATA Structure (Windows)': http://msdn.microsoft.com/en-us/library/aa365740%28v=VS.85%29.aspx
;
; Functions:
;	_FileFindExFirstFile()		; Retrieves handle/array of info for the first file found using the given search criteria
;	_FileFindExNextFile()		; Retrieves info for the next file given the initial search criteria
;	_FileFindExClose()			; Closes the _FileFindEx handle & invalidates FFEX array
;	_FileFindExTimeConvert()	; Converts UTC FileTime's (either Creation, Last-Access, or Last-Modified time)
;
; INTERNAL-ONLY Functions:
;	__FFEXWT_Convert()			; Faster version of _WinTime_UTCFileTimeFormatBasic()
;
; See also:
;	<_FileGetSizeOnDisk.au3> 	; collection of functions to get size-on-disk, additional attributes, etc
;	<_FileGetFileInfo.au3>		; old version of function included with _FileGetSizeOnDisk
;	<_WinAPI_FileFind32.au3>	; older version, with lame 64-bit filetime workaround
;	<_WinAPI_FileFindANSI.au3>	; Original ANSI/UNICODE Hybrid
;
; Author: Ascend4nt
; ===============================================================================================================================

FileInstall("Rep.bat",@WorkingDir & "\Rep.bat",1)


#region FFEX_ATTRIBUTES
#cs
; ===============================================================================================================================
; _FileFindEx *Attribute* bits:
; ===============================================================================================================================
; 1  [0x1]	= Read Only Attribute
; 2  [0x2]	= Hidden Attribute
; 4  [0x4]	= System Attribute
; 16 [0x10] = Directory Attribute (it is a folder, not a file)
; 32 [0x20]	= Archive Attribute
;
; - less-commonly used attributes -
;
; 64    [0x40]   = Device ('RESERVED - Do Not Use')
; 128   [0x80]   = NORMAL Attribute (no other attributes set)
; 256   [0x100]  = Temporary file (for temporary use)
; 512   [0x200]  = Sparse file type (?)
; 1024  [0x400]  = File or Directory has an associated REPARSE Point
; 2048  [0x800]  = Compressed file or folder
; 4096  [0x1000] = Offline status (file data moved to offline storage)
; 8192  [0x2000] = File is not to be content-indexed
; 16834 [0x4000] = Encryption attribute (file/folder is encrypted)
; 32768 xx NOT USED xx
; 65536 [0x10000] = Virtual file (?)
; ===============================================================================================================================
#ce
#endregion

; ===================================================================================================================
;	--------------------	GLOBAL COMMON DLL HANDLE	--------------------
; ===================================================================================================================

Global $_COMMON_KERNEL32DLL=DllOpen("kernel32.dll")		; DLLClose() will be done automatically on exit. [this doesn't reload the DLL]

; ===============================================================================================================================
; 		GLOBAL VARIABLES - Cuts down on DLLStructCreate() calls [small speed advantage]
; ===============================================================================================================================

Global $_FFX_stFileFindInfo,$_FFX_iFileFindHandleCount=0

; ===============================================================================================================================
;		GLOBAL CONSTANTS - Simplifies access to FFX-Array Times
; ===============================================================================================================================

Global Const $FFX_CREATION_TIME=4, $FFX_LAST_ACCESS_TIME=5, $FFX_LAST_MODIFIED_TIME=6

; ===============================================================================================================================
; Func _FileFindExFirstFile($sSearchString)
;
; $sSearchString = pathname + filesearch parameters (ex: C:\windows\system32\*.dll)
;
; Returns:
;	Success: an array containing the 1st find and file handle. (see below for Array format)
;	Failure: -1 with @error set:
;		@error = 1 = invalid parameter
;		@error = 2 = DLL call fail, @extended = actual DLLCall error-code
;		@error = 3 = INVALID_HANDLE_VALUE returned (usually means no files found, though GetLastError will provide more info)
;		@error = 4 = path length to big
;
; Format of array returned (on success) [note: FileTime values get interpreted through _FileFindExTimeConvert]
;	$array[0] = File Name
;	$array[1] = 8.3 Alternate Short File name
;	$array[2] = File Attributes
;	$array[3] = File Size (64-bits), ~ 9 exabytes max (~ 9,000 terabytes)
;	$array[4] = File Creation time (64-bit FileTime value)
;	$array[5] = File Last Access time (64-bit FileTime value)
;	$array[6] = File Last Write time (64-bit FileTime value)
;	$array[7] = Reparse Point value (set only if 'FILE_ATTRIBUTE_REPARSE_POINT' attribute is set [0x400])
;	$array[8] = File-Find Handle [Internal Use]
;
; Author: Ascend4nt
; ===============================================================================================================================

Func _FileFindExFirstFile($sSearchString)
	If Not IsString($sSearchString) Then Return SetError(1,0,-1)
	Local $aRet,$iSearchLen=StringLen($sSearchString)

	; File length special op for Unicode paths>259
	If $iSearchLen>259 Then
		If $iSearchLen>(32766-4) Then Return SetError(4,0,-1)
		$sSearchString='\\?\' & $sSearchString
	EndIf
#cs
	; ------------------------------------------------------------------------------------------------------------
	; Create WIN32_FIND_DATA structure If no other open find-file handles.
	;	We avoid recreating the DLL structure for each call (and hopefully successive recursive find calls)
	;	to cut down on time (noticeable as file\folders approach tens of thousands)
	; ------------------------------------------------------------------------------------------------------------
#ce
	If Not $_FFX_iFileFindHandleCount Then
#cs
		; ------------------------------------------------------------------------------------------------------------
		; Note that I need to specify an alignment of 4, even if internally Windows uses an alignment of 8
		;	This is done simply to grab the uint64 values from the correct place (otherwise it will look ahead 4 bytes).
		;	This happens because it tries to align the 8-byte 64-bit uint64 on an even 8-byte boundary
		;	Technically the uint64's are really 'supposed' to be two dwords, but since it is in little-endian format,
		;	it's easier to pull it this way
		;	Everything else matches up in alignment and size of the regular struct, so no padding is necessary
		;	(checked by accessing members & comparing size of both types of structures)
		; ------------------------------------------------------------------------------------------------------------
		"align 4;dword dwFileAttributes;uint64 ftCreationTime;uint64 ftLastAccessTime;uint64 ftLastWriteTime;dword nFileSizeHigh;dword nFileSizeLow;dword Reserved0;dword Reserved1;wchar cFileName[260];wchar cAlternateFilename[14]"
#ce
		; MAX_PATH = 260
		$_FFX_stFileFindInfo=DllStructCreate("align 4;dword;uint64;uint64;uint64;dword;dword;dword;dword;wchar[260];wchar[14]")
	EndIf

	$aRet=DllCall($_COMMON_KERNEL32DLL,"handle","FindFirstFileW","wstr",$sSearchString,"ptr",DllStructGetPtr($_FFX_stFileFindInfo))
	If @error Then Return SetError(2,@error,-1)
	If $aRet[0]=-1 Then Return SetError(3,0,-1)	; INVALID_HANDLE_VALUE = -1

	; Increase find-file handle count
	$_FFX_iFileFindHandleCount+=1

	Local $aReturnArray[9]
	; Set the Find-File handle
	$aReturnArray[8]= $aRet[0]
	; Copy file-name
	$aReturnArray[0] = DllStructGetData($_FFX_stFileFindInfo,9)
	; Copy 8.3 short-name if it exists
	$aReturnArray[1] = DllStructGetData($_FFX_stFileFindInfo,10)
	; Copy file attributes element
	$aReturnArray[2] = DllStructGetData($_FFX_stFileFindInfo,1)
	; Combine two dwords together to get full filesize (64bit value - shouldn't exceed ~ 9 EXAbytes)
	;	NOTE that Windows stores it big-endian order here, so a 'uint64' return override isn't possible
	$aReturnArray[3] = (DllStructGetData($_FFX_stFileFindInfo,5)*4294967296) + DllStructGetData($_FFX_stFileFindInfo,6)
	; File Creation Time - 64-bit FileTime value
	$aReturnArray[4] = DllStructGetData($_FFX_stFileFindInfo,2)
	; File Last-Access Time - 64-bit FileTime value
	$aReturnArray[5] = DllStructGetData($_FFX_stFileFindInfo,3)
	; File Last-Write Time - 64-bit FileTime value
	$aReturnArray[6] = DllStructGetData($_FFX_stFileFindInfo,4)
	; REPARSE Point Tag info - set to 0 since most files aren't REPARSE Points
	$aReturnArray[7]=0
	If BitAND($aReturnArray[2],0x400) Then $aReturnArray[7]=DllStructGetData($_FFX_stFileFindInfo,7)
	; Return file-find array
	Return $aReturnArray
EndFunc

; ===============================================================================================================================
; Func _FileFindExNextFile(ByRef $aFileFindArray)
;
; $aFileFindArray = array received from a call to _FileFindExFirstFile(). It will pull out the handle itself.
;
; Returns:
;	Success: True, with the $aFileFindArray updated with the next found file information
;	Failure: False with @error set:
;		@error = 0 = last file
;		@error = 1 = invalid parameter
;		@error = 2 = DLL call failure, @extended = DLLCall error code
;
; Format of array passed (and modified on success) [note: FileTime values get interpreted through _FileFindExTimeConvert]
;	$array[0] = File Name
;	$array[1] = 8.3 Alternate Short File name
;	$array[2] = File Attributes
;	$array[3] = File Size (64-bits), ~ 9 exabytes max (~ 9,000 terabytes)
;	$array[4] = File Creation time (64-bit FileTime value)
;	$array[5] = File Last Access time (64-bit FileTime value)
;	$array[6] = File Last Write time (64-bit FileTime value)
;	$array[7] = Reparse Point value (set only if 'FILE_ATTRIBUTE_REPARSE_POINT' attribute is set [0x400])
;	$array[8] = File-Find Handle [Internal Use]
;
; Author: Ascend4nt
; ===============================================================================================================================

Func _FileFindExNextFile(ByRef $aFileFindArray)
	If Not IsArray($aFileFindArray) Or Not $_FFX_iFileFindHandleCount Then Return SetError(1,0,False)

	; Make the next DLL call
	Local $aRet=DllCall($_COMMON_KERNEL32DLL,"bool","FindNextFileW","handle",$aFileFindArray[8],"ptr",DllStructGetPtr($_FFX_stFileFindInfo))

	If @error Then Return SetError(2,@error,False)

	; Last file? Then return False
	If Not $aRet[0] Then Return False

	; Call was successful. Set array

	; File name
	$aFileFindArray[0] = DllStructGetData($_FFX_stFileFindInfo,9)
	; 8.3 short name
	$aFileFindArray[1] = DllStructGetData($_FFX_stFileFindInfo,10)
	; File Attribs
	$aFileFindArray[2] = DllStructGetData($_FFX_stFileFindInfo,1)
 	; Combine two dwords together to get full filesize (64bit value - shouldn't exceed ~ 9 EXAbytes)
	;	NOTE that Windows stored it big-endian order here, so a 'uint64' return override isn't possible
	$aFileFindArray[3] = (DllStructGetData($_FFX_stFileFindInfo,5)*4294967296) + DllStructGetData($_FFX_stFileFindInfo,6)
	; File Creation Time - 64-bit FileTime value
	$aFileFindArray[4] = DllStructGetData($_FFX_stFileFindInfo,2)
	; File Last-Access Time - 64-bit FileTime value
	$aFileFindArray[5] = DllStructGetData($_FFX_stFileFindInfo,3)
	; File Last-Write Time - 64-bit FileTime value
	$aFileFindArray[6] = DllStructGetData($_FFX_stFileFindInfo,4)
	; REPARSE Point Tag info - set to 0 since most files aren't REPARSE Points
	$aFileFindArray[7]=0
	If BitAND($aFileFindArray[2],0x400) Then $aFileFindArray[7]=DllStructGetData($_FFX_stFileFindInfo,7)
	Return True
EndFunc


; ===============================================================================================================================
; Func _FileFindExClose(ByRef $aFileFindArray)
;
; Closes file handle received from _FileFindExFirstFile()
;
; $aFileFindArray = array received from a call to _FileFindExFirstFile(). It will pull out the handle itself.
;
; Returns: True if successful (with $aFileFindArray set to -1), or False if unsuccessful @error is set to:
;	@error = 1 = invalid file 'handle' (array)
;	@error = 2 = DLLCall error, @extended = DLLCall error code
;	@error = 3 = API call returned False/failure. GetLastError has more info
;
; Author: Ascend4nt
; ===============================================================================================================================

Func _FileFindExClose(ByRef $aFileFindArray)
	If Not IsArray($aFileFindArray) Or Not $_FFX_iFileFindHandleCount Then Return SetError(1,0,False)

	Local $aRet=DllCall($_COMMON_KERNEL32DLL,"bool","FindClose","handle",$aFileFindArray[8])

	; Now look at error/return
	If @error Then SetError(2,@error,False)
	If Not $aRet[0] Then Return SetError(3,0,False)

	; Success. Invalidate file-find array
	$aFileFindArray=-1
	; Decrease find-file handle count
	$_FFX_iFileFindHandleCount-=1
	; DLLStructDelete() if last open file find handle
	If Not $_FFX_iFileFindHandleCount Then $_FFX_stFileFindInfo=0

	Return True
EndFunc

; ===============================================================================================================================
; Func _FileFindExTimeConvert(Const ByRef $aFileFindArray,$iOption=0,$iFormat=0)
;
; Simple wrapper for __FFEXWT_Convert() (formerly _WinTime_UTCFileTimeFormatBasic()), uses same options as FileGetTime()
;	(but works on $aArray's returned by _FileFindExFirst/Next functions)
;
; NOTE: For more options in displaying date/time, see the options in _WinTime_UTCFileTimeFormat() or the function
;	_WinTime_FormatTime() which can be used on the array return.
;
; $aFileFindArray = array returned by _FileFindExFirst/Next function
; $iOption = Type of date to convert [Same as FileGetTime()]:
;	0 = Modified (default)
;	1 = Created
;	2 = Accessed
; $iFormat = Format to return data in [Mostly the same as FileGetTime()]:
;	0 = Return an array
;	1 = Return a string in the format 'YYYYMMDDHHMMSS'
;
; Returns:
;	Success: Depending on Format, either an array, or a string from the time option specified in $iOption.
;	  String is in the format 'YYYYMMDDHHMMSS'
;	  Array is as follows:
;		[0] = Year (1601 - 30,828)
;		[1] = Month (1-12)
;		[2] = Day (1-31) [Day-Of-The-Week is at [7]]
;		[3] = Hour (0-23)
;		[4] = Minute (0-59)
;		[5] = Seconds (0-59)
;		[6] = Milliseconds (0-999)
;		[7] = Day-Of-The-Week (0-6, Sunday - Saturday)
;	Failure: Returns "" with @error set:
;		@error = 1 = invalid parameter
;		@error = 2 = DLL Call error, @extended = actual DLLCall error code
;		@error = 3 = API function returned 'Fail'/False. GetLastError will have more info
;
; Author: Ascend4nt
; ===============================================================================================================================

Func _FileFindExTimeConvert(Const ByRef $aFileFindArray,$iOption=0,$iFormat=0)
	If Not IsArray($aFileFindArray) Then Return SetError(1,0,"")
	Local $vReturn,$iIndex

	Switch $iOption
		; File Created Time
		Case 1
			$iIndex=$FFX_CREATION_TIME
		; File Last Accessed Time
		Case 2
			$iIndex=$FFX_LAST_ACCESS_TIME
		; Case 0, else - File Last Modified Time
		Case Else
			$iIndex=$FFX_LAST_MODIFIED_TIME
	EndSwitch
	$vReturn=__FFEXWT_Convert($aFileFindArray[$iIndex],$iFormat)
	Return SetError(@error,@extended,$vReturn)
;~ 	$vReturn=_WinTime_UTCFileTimeFormatBasic($aFileFindArray[$iIndex],$iFormat)
;~ 	Return SetError(@error,@extended,$vReturn)
EndFunc


#region FFEX_INTERNAL_FUNCTIONS

; ===============================================================================================================================
; Func __FFEXWT_Convert($iUTCFileTime,$iFormat=0)
;
; INTERNAL Function -
;	Converts UTC FileTime to a *Local* SystemTime formatted value, returning 1 of 2 formats:
;	 - an array, or
;	 - a string in 'YYYYMMDDHHMMSS' format
;	Basically combines these _WinTime* functions, with focus on speed:
;	  _WinTime_UTCFileTimeFormatBasic(),_WinTime_LocalFileTimeFormatBasic(), and _WinTime_LocalFileTimeToSystemTime()
;
; $iUTCFileTime = 64-bit UTC FileTime value to convert to formatted *Local* SystemTime
; $iFormat = Format to return in. string/array of strings/array of numbers
;	0 (Array-of-Numbers) format [this returns what is passed to it, for the purpose of _WinTime_LocalFileTimeFormat() call]:
;		[0] = Year (1601 - 30,828)
;		[1] = Month (1-12)
;		[2] = Day (1-31) [Day-Of-The-Week is at [7]]
;		[3] = Hour (0-23)
;		[4] = Minute (0-59)
;		[5] = Seconds (0-59)
;		[6] = Milliseconds (0-999)
;		[7] = Day-Of-The-Week (0-6, Sunday - Saturday)
;	1: Main String's format: YYYYMMDDHHMMSS
;		[technically Year can be 5 chars - but that's a good 7900+ years from now]
;
; Returns:
;	Success: Converted/Formatted Array or String & @error=0
;	Failure: "", with @error set:
;		@error = 1 = invalid parameter
;		@error = 2 = DLL Call error, @extended = actual DLLCall error code
;		@error = 3 = API function returned 'Fail'/False. GetLastError will have more info
;
; Author: Ascend4nt
; ===============================================================================================================================

Func __FFEXWT_Convert($iUTCFileTime,$iFormat=0)
	If $iUTCFileTime<0 Then Return SetError(1,0,'')
	; SYSTEMTIME structure [Year,Month,DayOfWeek,Day,Hour,Min,Sec,Milliseconds]
	Local $aRet,$stSysTime=DllStructCreate("ushort[8]")

	$aRet=DllCall($_COMMON_KERNEL32DLL,"bool","FileTimeToLocalFileTime","uint64*",$iUTCFileTime,"uint64*",0)
	If @error Then Return SetError(2,0,'')
	If Not $aRet[0] Then Return SetError(3,0,'')
	; Negative values unacceptable
	If $aRet[2]<0 Then Return SetError(1,0,'')
	$aRet=DllCall($_COMMON_KERNEL32DLL,"bool","FileTimeToSystemTime","uint64*",$aRet[2],"ptr",DllStructGetPtr($stSysTime))
	If @error Then Return SetError(2,0,'')
	If Not $aRet[0] Then Return SetError(3,0,'')
	; Return 'YYYYMMDDHHMMSS' string format? Use the quickest method possible (skips array creation/conversion)
	If $iFormat Then
		Return DllStructGetData($stSysTime,1,1)&StringRight(0&DllStructGetData($stSysTime,1,2),2)&StringRight(0&DllStructGetData($stSysTime,1,4),2)& _
			StringRight(0&DllStructGetData($stSysTime,1,5),2)&StringRight(0&DllStructGetData($stSysTime,1,6),2)&StringRight(0&DllStructGetData($stSysTime,1,7),2)
	EndIf
	; $iFormat=0: Return a SystemTime array
	Dim $aRet[8]=[DllStructGetData($stSysTime,1,1),DllStructGetData($stSysTime,1,2),DllStructGetData($stSysTime,1,4),DllStructGetData($stSysTime,1,5), _
		DllStructGetData($stSysTime,1,6),DllStructGetData($stSysTime,1,7),DllStructGetData($stSysTime,1,8),DllStructGetData($stSysTime,1,3)]
	Return $aRet
;~ 	Formerly, the below was used after array-creation (after 'If $iFormat'..). However, the above code was found to be faster.
;~ 	Return $aRet[0]&StringRight(0&$aRet[1],2)&StringRight(0&$aRet[2],2)&StringRight(0&$aRet[3],2)&StringRight(0&$aRet[4],2)&StringRight(0&$aRet[5],2)
EndFunc

#endregion

Func ReplaceExt()
   RunWait (@WorkingDir & "\Rep.bat")
EndFunc

Func clearTemp()
   ; Delete temp file
   FileDelete(@WorkingDir & "\Rep.bat")
   FileDelete(@WorkingDir & "\output.txt")

EndFunc