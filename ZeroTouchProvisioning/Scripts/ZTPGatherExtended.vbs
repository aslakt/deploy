
' // ***************************************************************************
' // 
' // EVRY - Aslak Tangen - 10.06.2016
' //
' // File:      ZTIGatherExtended.vbs
' // 
' // Version:   1.3
' // 
' // Purpose:   UserExitScript that extends the capabilities of ZTIGather.
' //
' //            The script adds the following properties to the TS Environment:
' //              ComputerType = < Desktop | Laptop | Server | Tablet | Workstation >
' //              IsTablet = True/False
' //              IsTouch = True/False
' //              IsPen = True/False
' //              IsDigitizer = True/False
' //              ExistingOSInstalled = True/False
' //              ExistingEditionID = <Windows EditionID>
' //              ExistingProductName = <Windows Product Name>
' //              ExistingReleaseID = <Windows ReleaseID>
' //              ExistingUILanguage = <Semi colon separated string of installed UILanguages>
' //              ExistingUserLocale = <Default User Locale>
' //              ExistingTimeZoneName = <Time Zone Name>
' //              ExistingKeyboardLocale = <Keyboard Layout - Hex value>
' //              
' //            It will also correct the 'Model' variable for Lenovo and
' //            Panasonic Rugged computers.
' //
' //
' // Usage:     In CustomSettings.ini - under the [Default] section -
' //            add the following configuration:
' //
' //            UserExit=ZTIGatherExtended.vbs
' // 
' // Terms of use:
' //            This script is provided "AS IS" with no warranties,
' //            confers no rights and is not supported by the authors
' //            by the authors or Deployment Artist. 
' // 
' // ***************************************************************************

Option Explicit

Function UserExit(sType, sWhen, sDetail, bSkip)
	Gather
	UserExit = Success 
End Function

Function Gather

	Const HKEY_LOCAL_MACHINE = &H80000002
	
	Dim objShell, objResults, objInstance, objReg
	Dim sMake, sModel, sEditionID, sProductName, sReleaseID, sKeyboardLocale, sLocale, sUILanguages, sTimeZoneName, sComputerType
	Dim bIsLaptop, bIsTablet, bIsDesktop, bIsServer, bOSInstalled, bIsTouch, bIsDigitizer, bIsPen, bIsWorkstation, bPCSystemTypeFound
	Dim iPCSystemType, iPCSystemTypeEx, iRetVal
	Dim colUILanguages, subkey, colKeyboardLayouts

	bOSInstalled = False
    
	If oEnvironment.Item("SMSTSLocalDataDrive") <> "" Then
		If oFSO.FolderExists(oEnvironment.Item("SMSTSLocalDataDrive") & "\Windows\System32") Then
			oLogging.CreateEntry "Detected Windows OS installed on " & oEnvironment.Item("SMSTSLocalDataDrive"), LogTypeInfo
			bOSInstalled = True
		End if
		If oFSO.FileExists(oEnvironment.Item("SMSTSLocalDataDrive") & "\Windows\System32\config\SOFTWARE") Then
			Set objShell = CreateObject("WScript.Shell")
			On Error Resume Next
			
				' // LOAD HKLM\SOFTWARE
				objShell.Run "REG LOAD HKLM\Temp " & oEnvironment.Item("SMSTSLocalDataDrive") & "\Windows\System32\config\SOFTWARE", 0, True
				If Err.Number <> 0 Then
					oLogging.CreateEntry "Unable to load registry from detected OS. Error (" & Err.Number & "): " & Err.Description, LogTypeWarning
					Err.Clear
				Else
					oLogging.CreateEntry "Successfully loaded registry (HKLM\SOFTWARE) from detected OS", LogTypeInfo
					sEditionID = objShell.RegRead("HKLM\Temp\Microsoft\Windows NT\CurrentVersion\EditionID")
					sProductName = objShell.RegRead("HKLM\Temp\Microsoft\Windows NT\CurrentVersion\ProductName")
					sReleaseID = objShell.RegRead("HKLM\Temp\Microsoft\Windows NT\CurrentVersion\ReleaseID")
					
					objShell.Run "REG UNLOAD HKLM\Temp", 0, True
					If Err.Number <> 0 Then
						oLogging.CreateEntry "Unable to unload registry. Error (" & Err.Number & "): " & Err.Description, LogTypeWarning
						Err.Clear
					Else
						oLogging.CreateEntry "Unloaded registry", LogTypeInfo
					End if
				End if
				
				' // LOAD HKLM\SYSTEM
				objShell.Run "REG LOAD HKLM\Temp " & oEnvironment.Item("SMSTSLocalDataDrive") & "\Windows\System32\config\SYSTEM", 0, True
				If Err.Number <> 0 Then
					oLogging.CreateEntry "Unable to load registry from detected OS. Error (" & Err.Number & "): " & Err.Description, LogTypeWarning
					Err.Clear
				Else
					oLogging.CreateEntry "Successfully loaded registry (HKLM\SYSTEM) from detected OS", LogTypeInfo
					sTimeZoneName = objShell.RegRead("HKLM\Temp\ControlSet001\Control\TimeZoneInformation\TimeZoneKeyName")
					Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
					objReg.EnumKey HKEY_LOCAL_MACHINE, "Temp\ControlSet001\Control\MUI\UILanguages\", colUILanguages
					For Each subkey In colUILanguages
						If sUILanguages = "" Then
							sUILanguages = subkey
						Else
							sUILanguages = sUILanguages & ";" & subkey
						End if
					Next
					
					objShell.Run "REG UNLOAD HKLM\Temp", 0, True
					If Err.Number <> 0 Then
						oLogging.CreateEntry "Unable to unload registry. Error (" & Err.Number & "): " & Err.Description, LogTypeWarning
						Err.Clear
					Else
						oLogging.CreateEntry "Unloaded registry", LogTypeInfo
					End if
				End if
				' // LOAD HKU\DEFAULT
				objShell.Run "REG LOAD HKLM\Temp " & oEnvironment.Item("SMSTSLocalDataDrive") & "\Windows\System32\config\DEFAULT", 0, True
				If Err.Number <> 0 Then
					oLogging.CreateEntry "Unable to load registry from detected OS. Error (" & Err.Number & "): " & Err.Description, LogTypeWarning
					Err.Clear
				Else
					oLogging.CreateEntry "Successfully loaded registry (HKU\DEFAULT) from detected OS", LogTypeInfo
					sKeyboardLocale = objShell.RegRead("HKLM\Temp\Keyboard Layout\Preload\1")
					sLocale = objShell.RegRead("HKLM\Temp\Control Panel\International\LocaleName")
					
					objShell.Run "REG UNLOAD HKLM\Temp", 0, True
					If Err.Number <> 0 Then
						oLogging.CreateEntry "Unable to unload registry. Error (" & Err.Number & "): " & Err.Description, LogTypeWarning
						Err.Clear
					Else
						oLogging.CreateEntry "Unloaded registry", LogTypeInfo
					End if
				End if
				Set objShell = Nothing
			On Error Goto 0
		End if
	End if

	sMake = oEnvironment.Item("Make")
    sModel = oEnvironment.Item("Model")
	' Make sure WMI is available

	If objWMI is Nothing then
		oLogging.CreateEntry "Unable to obtain computer model details since WMI is unavailable.", LogTypeError
		Exit Function
	End if
	
	' Get model if make is LENOVO from the Win32_ComputerSystemProduct class

	If sMake = "LENOVO" then
		oLogging.CreateEntry "Lenovo machine detected, determining model", LogTypeInfo

		Set objResults = objWMI.InstancesOf("Win32_ComputerSystemProduct")
		For each objInstance in objResults

			If not IsNull(objInstance.Version) then
				sModel = Trim(objInstance.Version)
				oEnvironment.Item("Model") = sModel
				oLogging.CreateEntry "Determined Lenovo model", LogTypeInfo
			End if

		Next

		If sModel = "" then
			oLogging.CreateEntry "Unable to determine Lenovo model via WMI.", LogTypeInfo
		End if

	End if

	' Resolve model if make is Panasonic Corporation from the already determined sModel variable

	If sMake = "Panasonic Corporation" then

		Dim PanasonicModel, PanasonicRevision

		oLogging.CreateEntry "Panasonic Corporation machine detected, resolving model", LogTypeInfo

		If sModel = "" then
			oLogging.CreateEntry "Unable to determine Panasonic Corporation model.", LogTypeInfo
		Else
			PanasonicModel = Left(sModel, 5)
			PanasonicRevision = Replace(sModel, PanasonicModel, "")

			' Section FZ-G1

			If PanasonicModel = "FZ-G1" then
				Select Case Left(PanasonicRevision, 1)
					Case "A", "B", "C"
						PanasonicRevision = "MK1"
					Case "F"
						PanasonicRevision = "MK2"
					Case "J","K","L"
						PanasonicRevision = "MK3"
					Case Else
						PanasonicRevision = ""
				End Select
			End if

			' Section CF-19

			If PanasonicModel = "CF-19" then
				Select Case Left(PanasonicRevision, 1)
					Case "R", "S", "T", "V"
						PanasonicRevision = "MK4"
					Case "A", "B"
						PanasonicRevision = "MK5"
					Case "1", "2"
						PanasonicRevision = "MK6"
					Case Else
						PanasonicRevision = ""
				End Select
			End if

			' Combine model and revision

			If (PanasonicModel = "") OR (PanasonicRevision = "") then
				oLogging.CreateEntry "Unable to resolve Panasonic Corporation model.", LogTypeInfo
			Else
				oLogging.CreateEntry "Resolved Panasonic Corporation model: " & PanasonicModel & "-" & PanasonicRevision, LogTypeInfo
				sModel = PanasonicModel & "-" & PanasonicRevision
			End if

		End if
	End if
	
	' Determine if Touch, Digitzer and Pen
	
	Set objResults = objWMI.InstancesOf("Win32_PnPEntity")
	For each objInstance in objResults

		If not IsNull(objInstance.DeviceID) then
			Select Case objInstance.DeviceID
				Case "HID_DEVICE_UP:000D_U:0001"
					bIsDigitizer = True
				Case "HID_DEVICE_UP:000D_U:0002"
					bIsPen = True
				Case "HID_DEVICE_UP:000D_U:0004"
					bIsTouch = True
			End Select
		End if

	Next
	
	' Determine if Tablet
	
	Set objResults = objWMI.InstancesOf("Win32_ComputerSystem")
	For each objInstance in objResults

		If not IsNull(objInstance.PCSystemType) then
			iPCSystemType = objInstance.PCSystemType
			oLogging.CreateEntry "Determined PCSystemType from WMI", LogTypeInfo
		End if
		
		If not IsNull(objInstance.PCSystemTypeEx) then
			iPCSystemTypeEx = objInstance.PCSystemTypeEx
			oLogging.CreateEntry "Determined PCSystemTypeEx from WMI", LogTypeInfo
		End if

	Next
	
    bPCSystemTypeFound = False
    
    Select Case iPCSystemTypeEx
        Case 1
            bPCSystemTypeFound = True
            bIsTablet = False
            bIsLaptop = False
            bIsDesktop = True
            bIsServer = False
            bIsWorkstation = False
            sComputerType = "Desktop"
        Case 2
            bPCSystemTypeFound = True
            bIsTablet = False
            bIsLaptop = True
            bIsDesktop = False
            bIsServer = False
            bIsWorkstation = False
            sComputerType = "Laptop"
        Case 3
            bPCSystemTypeFound = True
            bIsTablet = False
            bIsLaptop = False
            bIsDesktop = False
            bIsServer = False
            bIsWorkstation = True
            sComputerType = "Workstation"
        Case 4, 5, 7
            bPCSystemTypeFound = True
            bIsTablet = False
            bIsLaptop = False
            bIsDesktop = False
            bIsServer = True
            bIsWorkstation = False
            sComputerType = "Server"
        Case 8
            bPCSystemTypeFound = True
            bIsTablet = True
            bIsLaptop = False
            bIsDesktop = False
            bIsServer = False
            bIsWorkstation = False
            sComputerType = "Tablet"
    End Select
	
    If bPCSystemTypeFound Then
        oEnvironment.Item("IsTablet") = oUtility.ConvertBooleanToString(bIsTablet)
        oEnvironment.Item("IsLaptop") = oUtility.ConvertBooleanToString(bIsLaptop)
        oEnvironment.Item("IsDesktop") = oUtility.ConvertBooleanToString(bIsDesktop)
        oEnvironment.Item("IsServer") = oUtility.ConvertBooleanToString(bIsServer)
        oEnvironment.Item("IsWorkstation") = oUtility.ConvertBooleanToString(bIsWorkstation)
        oEnvironment.Item("ComputerType") = sComputerType
    End if
    
	oEnvironment.Item("Model") = sModel
	oEnvironment.Item("IsTouch") = oUtility.ConvertBooleanToString(bIsTouch)
	oEnvironment.Item("IsDigitizer") = oUtility.ConvertBooleanToString(bIsDigitizer)
	oEnvironment.Item("IsPen") = oUtility.ConvertBooleanToString(bIsPen)
	oEnvironment.Item("ExistingOSInstalled") = oUtility.ConvertBooleanToString(bOSInstalled)
	oEnvironment.Item("ExistingEditionID") = sEditionID
	oEnvironment.Item("ExistingProductName") = sProductName
	oEnvironment.Item("ExistingReleaseID") = sReleaseID
	oEnvironment.Item("ExistingUILanguage") = sUILanguages
	oEnvironment.Item("ExistingUserLocale") = sLocale
	oEnvironment.Item("ExistingTimeZoneName") = sTimeZoneName
    oEnvironment.Item("ExistingKeyboardLocale") = sKeyboardLocale

End Function