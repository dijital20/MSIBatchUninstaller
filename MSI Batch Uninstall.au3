; MSI Batch Uninstaller
;   by Josh Schneider (josh.schneider-at-gmail-dot-com)

#RequireAdmin
#include <File.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Misc.au3>

Local $logfile = "MSI Batch Uninstall.log"

_FileWriteLog($logfile, "** Starting on " & @UserName & "@" & @ComputerName & @CRLF & "  OS: " & @OSType & " " & @OSVersion & " " & @OSArch & @CRLF)

Local $products = EnumerateProducts()

If UBound($products, 1) = 0 Then
	_FileWriteLog($logfile, "** $products contains no data! **" & @CRLF)
EndIf

; Define UI
$MainForm = GUICreate("Batch Uninstaller", 640, 480, -1, -1, $WS_SIZEBOX + $WS_MAXIMIZEBOX + $WS_MINIMIZEBOX)

;ListView and Items
$lvProductList = GUICtrlCreateListView("Product Name|Product Publisher|InstallDate|Uninstall String|Install Location", 4, 4, 640 - 8, 480 - 8 - 54 - 25, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), BitOR($WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT))
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)

Dim $lvProductListItems[UBound($products, 1)]
For $i = 1 To UBound($products, 1) - 1
	$lvProductListItems[$i] = GUICtrlCreateListViewItem($products[$i][1] & "|" & $products[$i][2] & "|" & $products[$i][3] & "|" & $products[$i][0] & "|" & $products[$i][4], $lvProductList)
Next

;Checkboxen
$cboxQuiet = GUICtrlCreateCheckbox("Use Basic UI (MSI only)", 5, 480 - 6 - 20 - 55)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$cboxLog = GUICtrlCreateCheckbox("Log Uninstall (MSI only)", 145, 480 - 6 - 20 - 55)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

;Buttons
$btnExit = GUICtrlCreateButton("Exit Without Uninstall", 5, 480 - 6 - 50, 120, 30)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$btnLogs = GUICtrlCreateButton("View Logs", 640 - 120 - 6 - 128, 480 - 6 - 50, 120, 30)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$btnExport = GUICtrlCreateButton("Export List", 640 - 120 - 120 - 6 - 128, 480 - 6 - 50, 120, 30)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$btnUninstall = GUICtrlCreateButton("Uninstall Selected", 640 - 120 - 6, 480 - 6 - 50, 120, 30)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

;Miscellaneous
_GUICtrlListView_SetColumnWidth($lvProductList, 0, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($lvProductList, 1, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($lvProductList, 2, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($lvProductList, 3, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($lvProductList, 4, $LVSCW_AUTOSIZE)

GUISetState(@SW_SHOW)
; End Define UI

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $btnExit
			Exit

		Case $btnUninstall
			;Process the list of ListViewItems for a list of Checked items
			Local $CheckedInstallers = ProcessListViewItems($lvProductListItems, GUICtrlRead($cboxQuiet), GUICtrlRead($cboxLog))

			;If the list of Checked ListViewItems contains 0 items then display a message and allow the user to return to the previous UI.
			If UBound($CheckedInstallers) = 1 Or $CheckedInstallers = "" Then
				MsgBox(16, "Error", "No items checked.")
				_FileWriteLog($logfile, "$CheckedInstallers contains no items.")
			Else
				GUISetState(@SW_HIDE, $MainForm)
				;Once the list has been reordered, run the installs.
				RunInstalls($CheckedInstallers)
				_FileWriteLog($logfile, "Install runs completed.")
				Exit
			EndIf

		Case $btnExport
			$outfile = FileSaveDialog("Export Program List...", @DesktopDir, "CSV File (*.csv)|Text File (*.txt)", 16, "InstalledPrograms.csv", $MainForm)
			ExportList($products, $outfile)

		Case $btnLogs
			ShellExecute("MSI Batch RegEnum.log")
			ShellExecute($logfile)

	EndSwitch
WEnd

Exit

;===== Functions ======
Func EnumerateProducts()
	Local $debuglog = "MSI Batch RegEnum.log"
	Local $logfile = "MSI Batch Uninstall.log"

	InitLog($debuglog)

	; _FileWriteLog($debuglog, "** Starting on " & @UserName & "@" & @ComputerName & @CRLF & "  OS: " & @OSType & " " & @OSVersion & " " & @OSArch & @CRLF)

	If @OSArch = "X86" Then
		Dim $InstallerInfoPaths[5]
		$InstallerInfoPaths[1] = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData"
		$InstallerInfoPaths[2] = "HKCU\SOFTWARE\Microsoft\Installer\Products"
		$InstallerInfoPaths[3] = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		$InstallerInfoPaths[4] = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	Else
		Dim $InstallerInfoPaths[9]
		$InstallerInfoPaths[1] = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData"
		$InstallerInfoPaths[2] = "HKCU\SOFTWARE\Microsoft\Installer\Products"
		$InstallerInfoPaths[3] = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		$InstallerInfoPaths[4] = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		$InstallerInfoPaths[5] = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData"
		$InstallerInfoPaths[6] = "HKCU64\SOFTWARE\Microsoft\Installer\Products"
		$InstallerInfoPaths[7] = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		$InstallerInfoPaths[8] = "HKCU64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	EndIf

	Dim $products[1][4]

	For $a = 1 To UBound($InstallerInfoPaths, 1) - 1
		; _FileWriteLog($debuglog, "** Checking path " & $a & " of " & UBound($InstallerInfoPaths, 1) - 1 & " ****************************************" & @CRLF)
		Dim $users[1]
		Local $InstallerInfoPath = $InstallerInfoPaths[$a]

		;Enumerate Users
		If StringInStr($InstallerInfoPath, "UserData") Then
			; _FileWriteLog($debuglog, "Checking for user subkeys in " & $InstallerInfoPath & @CRLF)
			Local $i = 0
			While 1
				$i = $i + 1
				; _FileWriteLog($debuglog, "Checking subkey " & $i)
				Local $subkey = RegEnumKey($InstallerInfoPath, $i)
				If @error Then
					; _FileWriteLog($debuglog, "Exiting check with Error " & @error)
					ExitLoop
				EndIf
				; _FileWriteLog($debuglog, "  Found subkey: " & $subkey & @CRLF)
				ReDim $users[UBound($users, 1) + 1]
				$users[UBound($users, 1) - 1] = $subkey
			WEnd
			; _FileWriteLog($debuglog, "  $users contains " & UBound($users, 1) - 1 & " elements." & @CRLF)
		Else
			; _FileWriteLog($debuglog, "Skipping user subkey check in " & $InstallerInfoPath & @CRLF)
			ReDim $users[UBound($users, 1) + 1]
			$users[UBound($users, 1) - 1] = ""
		EndIf

		;Search Products
		; _FileWriteLog($debuglog, "Checking for products..." & @CRLF)
		For $u = 1 To UBound($users, 1) - 1
			If StringLeft($InstallerInfoPath, 4) = "HKLM" Then
				; _FileWriteLog($debuglog, "  HKLM Path Detected" & @CRLF)
				$userPath = $InstallerInfoPath & "\" & $users[$u] & "\Products"
			ElseIf StringLeft($InstallerInfoPath, 4) = "HKCU" Then
				; _FileWriteLog($debuglog, "  HKCU Path Detected" & @CRLF)
				$userPath = $InstallerInfoPath & "\" & $users[$u]
			EndIf

			; _FileWriteLog($debuglog, "  (" & $u & ") Checking for products in " & $userPath & @CRLF)
			Local $p = 0
			While 1
				$p = $p + 1
				Local $product = RegEnumKey($userPath, $p)
				If @error Then
					; _FileWriteLog($debuglog, "Exiting check with Error " & @error)
					ExitLoop
				EndIf
				$ProductPath = $userPath & "\" & $product & "\InstallProperties"
				; _FileWriteLog($debuglog, "    Product found: " & $ProductPath & @CRLF)

				$ProductName = RegRead($ProductPath, "DisplayName")
				$ProductVersion = RegRead($ProductPath, "DisplayVersion")
				$ProductUninstall = RegRead($ProductPath, "UninstallString")
				$ProductPublisher = RegRead($ProductPath, "Publisher")
				$ProductRegOwner = RegRead($ProductPath, "RegOwner")
				$ProductInstallDate = RegRead($ProductPath, "InstallDate")
				$ProductInstallDate = StringMid($ProductInstallDate, 5, 2) & "/" & StringRight($ProductInstallDate, 2) & "/" & StringLeft($ProductInstallDate, 4)
				$ProductInstallLocation = RegRead($ProductPath, "InstallLocation")

				If Not $ProductUninstall Then
					; _FileWriteLog($debuglog, "    $ProductUninstall blank. Using LocalPackage instead." & @CRLF)
					If RegRead($ProductPath, "LocalPackage") Then
						$ProductUninstall = "MsiExec.exe /X " & RegRead($ProductPath, "LocalPackage")
					ElseIf RegRead($ProductPath, "UninstallString") Then
						$ProductUninstall = RegRead($ProductPath, "UninstallString")
					Else
						; _FileWriteLog($debuglog, "    LocalPackage and UninstallString blank or non-existant." & @CRLF)
					EndIf
				EndIf

				If StringLeft($ProductUninstall, 11) = "MsiExec.exe" Then
					If StringInStr($ProductUninstall, "/I") Then
						; _FileWriteLog($debuglog, "    Replacing /I with /X." & @CRLF)
						$ProductUninstall = StringReplace($ProductUninstall, "/I", "/X")
					EndIf
				EndIf

				; _FileWriteLog($debuglog, "      " & $ProductName & "; " & $ProductVersion & "; " & $ProductUninstall & "; " & $ProductPublisher & "; " & $ProductRegOwner & @CRLF)

				If $ProductUninstall Then
					ReDim $products[UBound($products, 1) + 1][5]
					$products[UBound($products, 1) - 1][0] = $ProductUninstall
					$products[UBound($products, 1) - 1][1] = $ProductName & " (" & $ProductVersion & ")"
					$products[UBound($products, 1) - 1][2] = $ProductPublisher
					$products[UBound($products, 1) - 1][3] = $ProductInstallDate
					$products[UBound($products, 1) - 1][4] = $ProductInstallLocation
					; _FileWriteLog($debuglog, "Found: " & $products[UBound($products, 1) - 1][1] & ", " & $products[UBound($products, 1) - 1][2] & ", " & $products[UBound($products, 1) - 1][0] & ", " & $ProductPath)
					_FileWriteLog($logfile, "Found: " & $products[UBound($products, 1) - 1][1] & ", " & $products[UBound($products, 1) - 1][2] & ", " & $products[UBound($products, 1) - 1][0] & ", " & $ProductPath)
				EndIf
			WEnd
			; _FileWriteLog($debuglog, "Between loops, Error set to " & @error)
		Next
	Next

	;Other locations to look at in the near future (Non MSI installers)
	;  HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall
	;    If "UninstallString" else skip
	;    "DisplayName" (optional)
	;    "InstallDate" (optional)
	;    "DisplayVersion" (optional)
	;    "Publisher" (optional)
    
	;  HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall
	;    If "UninstallString" else skip
	;    "DisplayName" (optional)
	;    "InstallDate" (optional)
	;    "DisplayVersion" (optional)
	;    "Publisher" (optional)

	_ArraySort($products, 0, 0, 0, 1)

	Return $products
EndFunc   ;==>EnumerateProducts

Func InitLog($logPath)
	Local $logfile = FileOpen($logPath, 2)
	If $logfile Then
		FileWrite($logfile, "")
		FileClose($logfile)
		Return True
	EndIf
	Return False
EndFunc   ;==>InitLog

Func ExportList($inArray, $outFilePath)
	If StringRight($outFilePath, 4) = ".csv" Then
		Local $delim = ","
	Else
		Local $delim = @TAB
	EndIf

	Local $outfile = FileOpen($outFilePath, 8 + 2)
	For $i = 1 To UBound($inArray, 1) - 1
		For $p = 1 To UBound($inArray, 2) - 1
			FileWrite($outfile, $inArray[$i][$p] & $delim)
		Next
		FileWrite($outfile, @CRLF)
	Next
	FileClose($outfile)
EndFunc   ;==>ExportList

; ProcessListViewItems($inList) ==>
;		Takes an array of ListViewItems ($inList), and processes them for checked items.
;		Returns an array containing the contents of checked items. Sets an error if $inList is not an array.
Func ProcessListViewItems($inList, $quiet = "", $logged = "")
	If Not IsArray($inList) Then
		SetError(1)
		Return ""
	EndIf

	Dim $outArray[1]
	Local $logfile = @ScriptDir & "\MSI Batch Uninstall.log"

	For $i = 1 To UBound($inList) - 1
		If GUICtrlRead($inList[$i], 1) = $GUI_CHECKED Then
			;Following lines modified from the original function in SilentInstallFolder2 to get around the unique structure of this listview.
			Local $runstring = StringMid(GUICtrlRead($inList[$i]), StringInStr(GUICtrlRead($inList[$i]), "|", 0, 3) + 1, StringInStr(GUICtrlRead($inList[$i]), "|", 0, 4) - StringInStr(GUICtrlRead($inList[$i]), "|", 0, 3) - 1)

			If $quiet = $GUI_CHECKED Then
				$runstring &= ' /qb'
			EndIf

			If $logged = $GUI_CHECKED Then
				Local $logfilename = StringRegExpReplace(StringLeft(GUICtrlRead($inList[$i]), StringInStr(GUICtrlRead($inList[$i]), "|") - 1), '[/:*?"<>|]', '_') & '.log'
				$runstring &= ' /l*v "' & @ScriptDir & '\' & $logfilename & '"'
			EndIf

			ReDim $outArray[UBound($outArray) + 1]
			$outArray[UBound($outArray) - 1] = $runstring
			_FileWriteLog($logfile, "Adding Item " & UBound($outArray) - 1 & ": " & $runstring)
		EndIf
	Next

	Return $outArray
EndFunc   ;==>ProcessListViewItems

; RunInstalls($inArray) ==>
;		Runs through each item in $inArray, executing it, displaying a progress dialog.
;		Returns nothing, sets an error if $inArray is not an array.
Func RunInstalls($inArray)
	If Not IsArray($inArray) Then
		SetError(1)
		Return ""
	EndIf

	Local $logfile = @ScriptDir & "\MSI Batch Uninstall.log"

	ProgressOn("Installing", "Installing " & UBound($inArray, 1) - 1 & " items.", "", 0, 0, 16)
	For $i = 1 To UBound($inArray, 1) - 1
		ProgressSet(($i / (UBound($inArray, 1) - 1)) * 100, $inArray[$i])
		_FileWriteLog($logfile, "Executing: " & $inArray[$i])
		RunWait($inArray[$i])
	Next
	ProgressOff()

	MsgBox(64, "Done!", "All installs have completed. Please see " & $logfile & " for details.", 30)

	Return ""
EndFunc   ;==>RunInstalls