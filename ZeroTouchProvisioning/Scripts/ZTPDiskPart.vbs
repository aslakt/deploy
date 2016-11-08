' // ***************************************************************************
' // 
' // EVRY - Aslak Tangen - 06.10.2016
' //
' // File:      ZTPDiskPart.vbs
' // 
' // Version:   2.0
' // 
' // Purpose:	This script exposes the diskpart command line to VBScript.
' //            It also provides functions for some common diskpart commands and
' //            returns easy to use objects while logging the process.
' //
' // Requirements:
' //            ZTIUtility.vbs must be loaded prior to this script for it to work.
' //
' // Terms of use:
' //            This script is provided "AS IS" with no warranties,
' //            confers no rights and is not supported by the authors
' //            by the authors or Deployment Artist. 
' // 
' // ***************************************************************************

Option Explicit

Dim objExec, objDict, objDisk, objPart, objVol
Dim sRetVal, line, temp, IgnoreThis

Set objExec = oShell.Exec("diskpart.exe")

Function ExitDiskPart

    objExec.StdIn.Write "exit" & VbCrLf

End Function

Function ExecuteDiskPartCommand (strCommand)

    objExec.StdIn.Write strCommand & VbCrLf

    Do While True
        IgnoreThis = objExec.StdOut.ReadLine & vbcrlf              
        If InStr(IgnoreThis, "DISKPART>") <> 0 Then Exit Do
    Loop

    objExec.StdIn.Write VbCrLf

    ExecuteDiskPartCommand = ""    

    Do While True
        ExecuteDiskPartCommand = ExecuteDiskPartCommand & objExec.StdOut.ReadLine & vbcrlf              
        If InStr(ExecuteDiskPartCommand, "DISKPART>") <> 0 Then Exit Do
    Loop
	
End Function

' Rescan returns True/False determined by rescan success
Function Rescan()
    oLogging.CreateEntry "DiskPart Command: RESCAN", LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("rescan")
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart has finished scanning your configuration.") > 0 Then
		Rescan = True
	Else
		Rescan = False
	End if
End Function


' SelectDisk returns True/False determined by disk selection success
Function SelectDisk(index)
    oLogging.CreateEntry "DiskPart Command: SELECT DISK " & index, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("select disk " & index)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "Disk " & index & " is now the selected disk.") Then
		SelectDisk = True
	Else
		SelectDisk = False
	End if
End Function


' SelectVolume returns True/False determined by volume selection success
Function SelectVolume(index)
    oLogging.CreateEntry "DiskPart Command: SELECT VOLUME " & index, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("select volume " & index)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "Volume " & index & " is the selected volume.") Then
		SelectVolume = True
	Else
		SelectVolume = False
	End if
End Function

' SelectPartition returns True/False determined by partition selection success
Function SelectPartition(index)
    oLogging.CreateEntry "DiskPart Command: SELECT PARTITION " & index, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("select partition " & index)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "Partition " & index & " is now the selected partition.") Then
		SelectPartition = True
	Else
		SelectPartition = False
	End if
End Function

' ShrinkPart returns True/False determined by partition shrink success
Function Shrink(argSize)
    oLogging.CreateEntry "DiskPart Command: SHRINK DESIRED=" & argSize & " MINIMUM=" & argSize, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("shrink desired=" & argSize & " minimum=" & argSize)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully shrunk the volume by:") Then
		Shrink = True
	Else
		Shrink = False
	End if
End Function

' CreatePart returns True/False determined by create partition success
Function CreatePart(args)
    oLogging.CreateEntry "DiskPart Command: CREATE PARTITION " & args, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("create partition " & args)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart succeeded in creating the specified partition.") Then
		CreatePart = True
	Else
		CreatePart = False
	End if
End Function

' Format returns True/False determined by format success
Function Format(args)
    oLogging.CreateEntry "DiskPart Command: FORMAT " & args, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("format " & args)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully formatted the volume.") Then
		Format = True
	Else
		Format = False
	End if
End Function

' Convert returns True/False determined by convert success
Function Convert(args)
    oLogging.CreateEntry "DiskPart Command: CONVERT " & UCase(args), LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("convert " & UCase(args))
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully converted the selected disk to " & UCase(args) & " format.") Then
		Convert = True
	Else
		Convert = False
	End if
End Function

' SetActive returns True/False determined by active success
Function SetActive()
    oLogging.CreateEntry "DiskPart Command: ACTIVE", LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("active")
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart marked the current partition as active.") Then
		SetActive = True
	Else
		SetActive = False
	End if
End Function

' DeletePartition returns True/False determined by delete partition success
Function DeletePartition()
    oLogging.CreateEntry "DiskPart Command: DELETE PART OVERRIDE", LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("delete part override")
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully deleted the selected partition.") Then
		DeletePartition = True
	Else
		DeletePartition = False
	End if
End Function

' Extend returns True/False determined by extend success
Function Extend()
    oLogging.CreateEntry "DiskPart Command: EXTEND", LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("extend")
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully extended the volume.") Then
		Extend = True
	Else
		Extend = False
	End if
End Function

' HideVolume returns True/False determined by attribute set success
Function HideVolume()
    oLogging.CreateEntry "DiskPart Command: ATTRIBUTES VOLUME SET HIDDEN", LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("ATTRIBUTES VOLUME SET HIDDEN")
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "Volume attributes set successfully.") Then
		HideVolume = True
	Else
		HideVolume = False
	End if
End Function

' UnHideVolume returns True/False determined by attribute clear success
Function UnHideVolume()
    oLogging.CreateEntry "DiskPart Command: ATTRIBUTES VOLUME CLEAR HIDDEN", LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("ATTRIBUTES VOLUME CLEAR HIDDEN")
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "Volume attributes cleared successfully.") Then
		UnHideVolume = True
	Else
		UnHideVolume = False
	End if
End Function

' Assign returns True/False determined by assign success
Function Assign(argDriveLetter)
    oLogging.CreateEntry "DiskPart Command: ASSIGN LETTER=" & argDriveLetter, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("assign letter=" & argDriveLetter)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully assigned the drive letter or mount point.") Then
		Assign = True
	Else
		Assign = False
	End if
End Function

' SetID returns True/False determined by SetID success
Function SetID(argID)
    oLogging.CreateEntry "DiskPart Command: SET ID=" & Chr(34) & argID & Chr(34), LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("set id=" & Chr(34) & argID & Chr(34))
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully set the partition ID.") Then
		SetID = True
	Else
		SetID = False
	End if
End Function

' GPTAttributes returns True/False determined by SetID success
Function GPTAttributes(argAttribute)
    oLogging.CreateEntry "DiskPart Command: GPT ATTRIBUTES=" & argAttribute, LogTypeInfo
    sRetVal = ExecuteDiskPartCommand("gpt attributes=" & argAttribute)
    If sRetVal <> "" AND InStr(Trim(sRetVal),"DISKPART>") = 0 Then
        oLogging.CreateEntry sRetVal, LogTypeInfo
    End if
	If InStr(sRetVal, "DiskPart successfully assigned the attributes to the selected GPT partition.") Then
		GPTAttributes = True
	Else
		GPTAttributes = False
	End if
End Function

' ListDisk returns a Dictinary of Disk objects
Function ListDisk()
	Set objDict = CreateObject("Scripting.Dictionary")
    oLogging.CreateEntry "DiskPart Command: LIST DISK", LogTypeInfo
	sRetVal = Split(ExecuteDiskPartCommand("list disk"), VbCrLf)
	For Each line in sRetVal
		If Trim(line) <> "" Then
			If InStr(Trim(line),"Disk ###") = 0 AND InStr(Trim(line),"------") = 0 AND InStr(Trim(line),"DISKPART>") = 0 Then
                oLogging.CreateEntry Trim(line), LogTypeInfo
                If Left(Trim(line),1) = "*" Then
                    Temp = True
                    line = Mid(Trim(line),2)
                Else
                    Temp = False
                End if
				Set objDisk = (New Disk)(Trim(Mid(Trim(line), 6, 3)), Trim(Mid(Trim(line), 11, 13)), Trim(Mid(Trim(line), 26, 7)), Trim(Mid(Trim(line), 35, 7)), Trim(Mid(Trim(line), 44, 3)), Trim(Mid(Trim(line), 49, 3)), Temp)
                objDict.Add Trim(Mid(line, 6, 3)), objDisk
			End if
		End if
	Next
	Set ListDisk = objDict
End Function

' ListPartition returns a Dictionary of Partition objects
Function ListPartition()
	Set objDict = CreateObject("Scripting.Dictionary")
    oLogging.CreateEntry "DiskPart Command: LIST PARTITION", LogTypeInfo
    sRetVal = Split(ExecuteDiskPartCommand("list partition"), VbCrLf)
    For Each line in sRetVal
        If Trim(line) <> "" Then
            If Trim(line) = "There is no disk selected to list partitions." Then
                oLogging.CreateEntry Trim(line), LogTypeWarning
                Exit Function
            ElseIf InStr(Trim(line),"Partition ###") = 0 AND InStr(Trim(line),"------") = 0 AND InStr(Trim(line),"DISKPART>") = 0 Then
                oLogging.CreateEntry Trim(line), LogTypeInfo
                If Left(Trim(line),1) = "*" Then
                    Temp = True
                    line = Mid(Trim(line),2)
                Else
                    Temp = False
                End if
                Set objPart = (New Partition)(Trim(Mid(Trim(line), 11, 3)), Trim(Mid(Trim(line), 16, 16)), Trim(Mid(Trim(line), 34, 7)), Trim(Mid(Trim(line), 43, 7)), Temp)
                objDict.Add Trim(Mid(line, 11, 3)), objPart
            End if
        End if
    Next
	Set ListPartition = objDict
End Function

' ListVolume returns a Dictionary of Volume objects
Function ListVolume()
    oLogging.CreateEntry "DiskPart Command: LIST VOLUME", LogTypeInfo
	Set objDict = CreateObject("Scripting.Dictionary")
	sRetVal = Split(ExecuteDiskPartCommand("list volume"), VbCrLf)
	For Each line in sRetVal
		If Trim(line) <> "" Then
			If InStr(Trim(line),"Volume ###") = 0 AND InStr(Trim(line),"------") = 0 AND InStr(Trim(line),"DISKPART>") = 0 Then
                oLogging.CreateEntry Trim(line), LogTypeInfo
                If Left(Trim(line),1) = "*" Then
                    Temp = True
                    line = Mid(Trim(line),2)
                Else
                    Temp = False
                End if
				Set objVol = (New Volume)(Trim(Mid(Trim(line), 8, 3)), Trim(Mid(Trim(line), 13, 3)), Trim(Mid(Trim(line), 18, 11)), Trim(Mid(Trim(line), 31, 5)), Trim(Mid(Trim(line), 38, 10)), Trim(Mid(Trim(line), 50, 7)), Trim(Mid(Trim(line), 59, 9)), Trim(Mid(Trim(line), 70, 8)), Temp)
				objDict.Add Trim(Mid(Trim(line), 8, 3)), objVol
			End if
		End if
	Next
	Set ListVolume = objDict
End Function


Class Disk

	Public Index
	Public Status
	Public Size
	Public Free
    Public Dynamic
    Public GPT
    Public Selected
	
	Public Default Function Init(argIndex, argStatus, argSize, argFree, argDynamic, argGPT, argSelected)
		Index = argIndex
		Status = argStatus
		Size = argSize
		Free = argFree
        If argDynamic <> "" Then
            Dynamic = True
        Else
            Dynamic = False
        End if
        If argGPT <> "" Then
            GPT = True
        Else
            GPT = False
        End if
        Selected = argSelected
		Set Init = Me
	End Function

    Public Function ToString()
        ToString =  "Disk " & Index & ":" & VbCrLf &_
                    " " & "Status: " & Status & VbCrLf &_
                    " " & "Size: " & Size & VbCrLf &_
                    " " & "Free: " & Free & VbCrLf
    End Function
	
End Class

Class Partition

	Public Index
	Public PartitionType
	Public Size
	Public Offset
    Public Selected
	
	Public Default Function Init(argIndex, argType, argSize, argOffset, argSelected)
		Index = argIndex
		PartitionType = argType
		Size = argSize
		Offset = argOffset
        Selected = argSelected
		Set Init = Me
	End Function

    Public Function ToString()
        ToString =  "Partition " & Index & VbCrLf &_
                    " " & "Type: " & PartitionType & VbCrLf &_
                    " " & "Size: " & Size & VbCrLf &_
                    " " & "Offset: " & Offset & VbCrLf
    End Function
	
End Class

Class Volume

	Public Index
    Public Letter
	Public Label
	Public FileSystem
	Public VolumeType
	Public Size
	Public Status
	Public Info
    Public Selected
	
	Public Default Function Init(argIndex, argLetter, argLabel, argFileSystem, argType, argSize, argStatus, argInfo, argSelected)
		Index = argIndex
		Letter = argLetter
	    Label = argLabel
	    FileSystem = argFileSystem
	    VolumeType = argType
	    Size = argSize
	    Status = argStatus
	    Info = argInfo
        Selected = argSelected
		Set Init = Me
	End Function

    Public Function ToString()
        ToString =  "Volume " & Index & ":" & VbCrLf &_
                    " " & "Letter: " & Letter & VbCrLf &_
                    " " & "Label: " & Label & VbCrLf &_
                    " " & "FileSystem: " & FileSystem & VbCrLf &_
                    " " & "Type: " & VolumeType & VbCrLf &_
                    " " & "Size: " & Size & VbCrLf &_
                    " " & "Status: " & Status & VbCrLf &_
                    " " & "Info: " & Info & VbCrLf
    End Function
	
End Class