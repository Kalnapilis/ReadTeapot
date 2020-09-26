Option Strict Off
Option Explicit On
Module GlobalsSSM
	
    Const gstrSEP_URLDIR As String = "/" ' Separator for dividing directories in URL
	'                                   addresses.
	Const gstrNULL As String = "" ' Empty string
	Const gstrSEP_DIR As String = "\" ' Directory separator character
	Const gstrQUOTE As String = """"
	
	Public Const kName As Short = 0 ' column # of name from GetAllSettings
	Public Const kValue As Short = 1 ' column # of name from GetAllSettings
	
	
    Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, ByRef lSize As Integer) As Integer
	
	Public Function ReplaceStr(ByRef TextIn As Object, ByRef SearchStr As Object, ByRef Replacement As Object, ByRef CompMode As Short) As Object
		Dim WorkText As String
		Dim Pointer As Short
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(TextIn) Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ReplaceStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ReplaceStr = System.DBNull.Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object TextIn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			WorkText = TextIn
			'UPGRADE_WARNING: Couldn't resolve default property of object SearchStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Pointer = InStr(1, WorkText, SearchStr, CompMode)
			Do While Pointer > 0
				'UPGRADE_WARNING: Couldn't resolve default property of object Replacement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
				'UPGRADE_WARNING: Couldn't resolve default property of object SearchStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr, CompMode)
			Loop 
			'UPGRADE_WARNING: Couldn't resolve default property of object ReplaceStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ReplaceStr = WorkText
		End If
	End Function
	
	Public Function Drop(ByVal S As String, ByVal Count As Short) As String
		'Drop Count characters from the left end of S
        Drop = S.Remove(0, Count)
	End Function 'Drop
	
	Public Function Head(ByVal S As String, ByVal Count As Short) As String
		'Return the Count leftmost characters of S
        Head = S.Substring(0, Count)
	End Function 'Drop
	
	Public Function Parse(ByRef S As String, ByVal Delim As String) As String
		'drops token and delimiter from head of S and returns Token
		Dim Count As Short
		Dim bQuoted As Boolean
		Dim token As String
		
		bQuoted = False
		token = ""
		S = LTrim(S)
		
		If Left(S, 1) = """" Then
			S = Drop(S, 1)
			Count = InStr(S, """")
			If Count > 0 Then 'as expected, there's an ending one also
				bQuoted = True
				token = Left(S, Count - 1)
				token = Trim(token)
				S = Drop(S, Count) 'drop token and quote
				S = LTrim(S)
				'S is now probably headed by a delimiter
				'else ? if there isn't an ending quote
			End If
		End If
		
		Count = InStr(S, Delim)
		
		If Count > 0 Or bQuoted Then
			If Not bQuoted Then
				token = Left(S, Count - 1)
				token = Trim(token)
			End If
			S = Drop(S, Count) 'drop token and delimiter
			S = LTrim(S)
		Else 'no ending token, return the rest
			token = Trim(S)
			S = ""
		End If
		
		Parse = token
	End Function 'Parse
	
	'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function DelimCount(ByVal S As String, ByVal Char_Renamed As String) As Short
		'determine the number of times the leftmost character of Char occurs in S
		'UPGRADE_NOTE: Loc was upgraded to Loc_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Loc_Renamed, Start, Count As Short
		
		Start = 1
		DelimCount = 0
		If Len(Char_Renamed) > 1 Then Char_Renamed = Left(Char_Renamed, 1)
		While Len(S) - Start > -1
			If Len(S) > 0 Then
				If Mid(S, Start, 1) = """" Then
					Count = InStr(Start + 1, S, """")
					If Count > 0 Then 'as expected, there's an ending one also
						Start = Count + 1 'step over token and quote
						'start probably now points to a delimiter
						'else ? if there isn't an ending quote
					End If
				End If
			End If 'checking for quoted string
			Loc_Renamed = InStr(Start, S, Char_Renamed)
			If Loc_Renamed > 0 Then
				DelimCount = DelimCount + 1
				Start = Loc_Renamed + 1
			Else
				Start = Len(S) + 1
			End If
		End While
	End Function 'DelimCount
	
   'Public Function StrToHex(ByRef S As Object) As Object
   '   '
   '   ' Converts a string to a series of hexadecimal digits.
   '   ' Useful if you want a true ASCII sort in your query.
   '   '
   '   ' StrToHex(Chr(9) & "A~") returns "09417E"
   '   '
   '   Dim Temp As String
   '   Dim I As Short
   '   If VarType(S) <> VariantType.String Then
   '      StrToHex = S
   '   Else
   '      Temp = ""
   '      For I = 1 To Len(S)
   '         'UPGRADE_WARNING: Couldn't resolve default property of object S. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
   '         Temp = Temp & Microsoft.VisualBasic.Format(Hex(Asc(Mid(S, I, 1))), "00")
   '      Next I
   '      'UPGRADE_WARNING: Couldn't resolve default property of object StrToHex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
   '      StrToHex = Temp
   '   End If
   'End Function
	
	Public Function HexToStr(ByRef S As Object) As Object
		'
		' Converts hexadecimal digits to a string.
		'
		Dim Temp As String
		Dim I As Short
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If VarType(S) <> VariantType.String Then
			'UPGRADE_WARNING: Couldn't resolve default property of object S. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object HexToStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			HexToStr = S
		Else
			Temp = ""
			For I = 1 To Len(S) Step 2
				'UPGRADE_WARNING: Couldn't resolve default property of object S. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Temp = Temp & Chr(Val("&h" & Mid(S, I, 2)))
			Next I
			'UPGRADE_WARNING: Couldn't resolve default property of object HexToStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			HexToStr = Temp
		End If
	End Function
	
	'-----------------------------------------------------------
	' FUNCTION: FileExists
	' Determines whether the specified file exists
	'
	' IN: [strPathName] - file to check for
	'
	' Returns: True if file exists, False otherwise
	'-----------------------------------------------------------
	'
	Public Function FileExists(ByVal strPathName As String) As Short
		Dim intFileNum As Short
		
		On Error Resume Next
		
		'
		' If the string is quoted, remove the quotes.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object strUnQuoteString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strPathName = strUnQuoteString(strPathName)
		'
		'Remove any trailing directory separator character
		'
		If Right(strPathName, 1) = gstrSEP_DIR Then
			strPathName = Left(strPathName, Len(strPathName) - 1)
		End If
		
		'
		'Attempt to open the file, return value of this function is False
		'if an error occurs on open, True otherwise
		'
		intFileNum = FreeFile
		FileOpen(intFileNum, strPathName, OpenMode.Input)
		
		FileExists = IIf(Err.Number = 0, True, False)
		
		FileClose(intFileNum)
		
		Err.Clear()
	End Function ' FileExists
	
   Public Function strUnQuoteString(ByVal strQuotedString As String) As String
      '
      ' This routine tests to see if strQuotedString is wrapped in quotation
      ' marks, and, if so, remove them.
      '
      strQuotedString = Trim(strQuotedString)

      If Mid(strQuotedString, 1, 1) = gstrQUOTE And Right(strQuotedString, 1) = gstrQUOTE Then
         '
         ' It's quoted.  Get rid of the quotes.
         '
         strQuotedString = Mid(strQuotedString, 2, Len(strQuotedString) - 2)
      End If
      strUnQuoteString = strQuotedString
   End Function
	
	'-----------------------------------------------------------
	' FUNCTION: DirExists
	'
	' Determines whether the specified directory name exists.
	' This function is used (for example) to determine whether
	' an installation floppy is in the drive by passing in
	' something like 'A:\'.
	'
	' IN: [strDirName] - name of directory to check for
	'
	' Returns: True if the directory exists, False otherwise
	'-----------------------------------------------------------
	'
	Public Function DirExists(ByVal strDirName As String) As Boolean
		Const strWILDCARD As String = "*.*"
		
		Dim strDummy As String
		
		On Error Resume Next
		
		AddDirSep(strDirName)
		strDirName = strDirName & strWILDCARD
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		strDummy = Dir(strDirName, FileAttribute.Directory)
		System.Diagnostics.Debug.Write("strDirName=" & strDirName & ". strDummy=>" & strDummy)
      Debug.Print("< for " & Microsoft.VisualBasic.Format(Len(strDummy), "###0"))
		'strDummy = Dir$(strDirName) ' , vbDirectory
		'Debug.Print "strDirName=" & strDirName; ". strDummy=>"; strDummy;
		'Debug.Print "< for "; Format(Len(strDummy), "###0")
		DirExists = Not (strDummy = gstrNULL)
		Debug.Print("Direxists :" & strDirName & " >" & strDummy & "<")
		
		Err.Clear()
	End Function
	
	'-----------------------------------------------------------
	' SUB: AddDirSep
	' Add a trailing directory path separator (back slash) to the
	' end of a pathname unless one already exists
	'
	' IN/OUT: [strPathName] - path to add separator to
	'-----------------------------------------------------------
	'
	Public Sub AddDirSep(ByRef strPathName As String)
		If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
			strPathName = RTrim(strPathName) & gstrSEP_DIR
		End If
	End Sub
	
	
	Private Function StartsWith(ByRef SI As String, ByRef Target As String) As Boolean
        'does SI start with target?
        StartsWith = False
		If StrComp(Left(UCase(SI), Len(Target)), UCase(Target)) = 0 Then
			StartsWith = True
		End If
	End Function
	
   Public Function Min(ByRef a As Object, ByRef b As Object) As Object
      'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object Min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      Min = b
      'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object Min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If a < b Then Min = a
   End Function
	Public Function Max(ByRef a As Object, ByRef b As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Max = b
		'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If a > b Then Max = a
	End Function
	
	Public Sub ShowModalForm(ByRef frmTarget As System.Windows.Forms.Form)
        Dim ofrm As Object
		
		
		'Disable all the forms
      For Each ofrm In Application.OpenForms
         'UPGRADE_WARNING: Couldn't resolve default property of object ofrm.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
         ofrm.Enabled = False
      Next ofrm
		
		'Now show the target form non-modal
		frmTarget.Show()
		'If the frmTarget was disabled by the loop above
		'make sure it is now enabled
		frmTarget.Enabled = True
		
		'Sit in a loop until the target form is dismissed
		Do While frmTarget.Visible = True
			System.Windows.Forms.Application.DoEvents()
		Loop 
		
		'We have left the loop, so the dialog has been dismissed
		'Now Enable the forms, and exit the procedure
      For Each ofrm In Application.OpenForms
         'UPGRADE_WARNING: Couldn't resolve default property of object ofrm.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
         ofrm.Enabled = True
      Next ofrm
	End Sub
	
	Sub FParseFullPath(ByVal FullPath As String, ByRef Drive As String, ByRef DirName As String, ByRef FName As String, ByRef Ext As String)
		'
		' Parses drive, directory, filename, and extension into separate variables.
		'
		' Assumptions/Gotcha's:
		' 1.  If a drive letter isn't specified, returns the current drive.
		' 2.  If the directory doesn't start from the root, prepends the current directory name
		'     for the current/selected drive.  This could cause problems if the selected drive
		'     doesn't exist on this machine.
		'
		Dim I, Found As Short
		Dim f As String
		Drive = Left(CurDir(), 2) ' Current drive if none explicitly specified
		DirName = ""
		FName = ""
		Ext = ""
		FullPath = Trim(FullPath)
		'
		' Get drive letter
		'
		If Mid(FullPath, 2, 1) = ":" Then
			Drive = Left(FullPath, 2)
			FullPath = Mid(FullPath, 3)
		End If
		'
		' Get directory name
		'
		f = ""
		Found = False
		For I = Len(FullPath) To 1 Step -1
			If Mid(FullPath, I, 1) = "\" Then
				f = Mid(FullPath, I + 1)
				DirName = Left(FullPath, I)
				Found = True
				Exit For
			End If
		Next I
		If Not Found Then
			f = FullPath
		End If
		'
		' Add current directory of selected drive if absolute directory not specified
		'
		If DirName = "" Or Left(DirName, 1) <> "\" Then
			DirName = Mid(CurDir(Left(Drive, 1)), 3) & "\" & DirName
		End If
		'
		' Get File name and extension
		'
		If f = "." Or f = ".." Then
			FName = f
		Else
			I = InStr(f, ".")
			If I > 0 Then
				FName = Left(f, I - 1)
				Ext = Mid(f, I)
			Else
				FName = f
			End If
		End If
	End Sub
	
	Sub FParsePath(ByVal FullPath As String, ByRef Drive As String, ByRef DirName As String, ByRef FName As String, ByRef Ext As String)
		'
		' Parses drive, directory, filename, and extension into separate variables.
		' Returns blank drive letter/path if none specified.
		'
		Dim I, Found As Short
		Dim f As String
		Drive = ""
		DirName = ""
		FName = ""
		Ext = ""
		FullPath = Trim(FullPath)
		'
		' Get drive letter
		'
		If Mid(FullPath, 2, 1) = ":" Then
			Drive = Left(FullPath, 2)
			FullPath = Mid(FullPath, 3)
		End If
		'
		' Get directory name
		'
		f = ""
		Found = False
		For I = Len(FullPath) To 1 Step -1
			If Mid(FullPath, I, 1) = "\" Then
				f = Mid(FullPath, I + 1)
				DirName = Left(FullPath, I)
				Found = True
				Exit For
			End If
		Next I
		If Not Found Then
			f = FullPath
		End If
		'
		' Get File name and extension
		'
		If f = "." Or f = ".." Then
			FName = f
		Else
			I = InStr(f, ".")
			If I > 0 Then
				FName = Left(f, I - 1)
				Ext = Mid(f, I)
			Else
				FName = f
			End If
		End If
	End Sub
	
	' Evalutate the 16-bit CRC (Cyclic Redundancy Checksum) of an array of bytes
	'
	' If you omit the second argument, the entire array is considered
	
	Public Function Crc16(ByRef cp() As Byte, Optional ByVal Size As Integer = -1) As Integer
		Dim I As Integer
		Dim fcs As Integer
		Static fcstab(256) As Integer
        'Dim L5, L3, L1, L2, L4, L6 As Integer
		
        If Size < 0 Then Size = UBound(cp) - LBound(cp) + 1
		
		If fcstab(1) = 0 Then
			' Initialize array once and for all
			fcstab(0) = &H0
			fcstab(1) = &H8005
			fcstab(2) = &H800F
			fcstab(3) = &HA
			fcstab(4) = &H801B
			fcstab(5) = &H1E
			fcstab(6) = &H14
			fcstab(7) = &H8011
			fcstab(8) = &H8033
			fcstab(9) = &H36
			fcstab(10) = &H3C
			fcstab(11) = &H8039
			fcstab(12) = &H28
			fcstab(13) = &H802D
			fcstab(14) = &H8027
			fcstab(15) = &H22
			fcstab(16) = &H8063
			fcstab(17) = &H66
			fcstab(18) = &H6C
			fcstab(19) = &H8069
			fcstab(20) = &H78
			fcstab(21) = &H807D
			fcstab(22) = &H8077
			fcstab(23) = &H72
			fcstab(24) = &H50
			fcstab(25) = &H8055
			fcstab(26) = &H805F
			fcstab(27) = &H5A
			fcstab(28) = &H804B
			fcstab(29) = &H4E
			fcstab(30) = &H44
			fcstab(31) = &H8041
			fcstab(32) = &H80C3
			fcstab(33) = &HC6
			fcstab(34) = &HCC
			fcstab(35) = &H80C9
			fcstab(36) = &HD8
			fcstab(37) = &H80DD
			fcstab(38) = &H80D7
			fcstab(39) = &HD2
			fcstab(40) = &HF0
			fcstab(41) = &H80F5
			fcstab(42) = &H80FF
			fcstab(43) = &HFA
			fcstab(44) = &H80EB
			fcstab(45) = &HEE
			fcstab(46) = &HE4
			fcstab(47) = &H80E1
			fcstab(48) = &HA0
			fcstab(49) = &H80A5
			fcstab(50) = &H80AF
			fcstab(51) = &HAA
			fcstab(52) = &H80BB
			fcstab(53) = &HBE
			fcstab(54) = &HB4
			fcstab(55) = &H80B1
			fcstab(56) = &H8093
			fcstab(57) = &H96
			fcstab(58) = &H9C
			fcstab(59) = &H8099
			fcstab(60) = &H88
			fcstab(61) = &H808D
			fcstab(62) = &H8087
			fcstab(63) = &H82
			fcstab(64) = &H8183
			fcstab(65) = &H186
			fcstab(66) = &H18C
			fcstab(67) = &H8189
			fcstab(68) = &H198
			fcstab(69) = &H819D
			fcstab(70) = &H8197
			fcstab(71) = &H192
			fcstab(72) = &H1B0
			fcstab(73) = &H81B5
			fcstab(74) = &H81BF
			fcstab(75) = &H1BA
			fcstab(76) = &H81AB
			fcstab(77) = &H1AE
			fcstab(78) = &H1A4
			fcstab(79) = &H81A1
			fcstab(80) = &H1E0
			fcstab(81) = &H81E5
			fcstab(82) = &H81EF
			fcstab(83) = &H1EA
			fcstab(84) = &H81FB
			fcstab(85) = &H1FE
			fcstab(86) = &H1F4
			fcstab(87) = &H81F1
			fcstab(88) = &H81D3
			fcstab(89) = &H1D6
			fcstab(90) = &H1DC
			fcstab(91) = &H81D9
			fcstab(92) = &H1C8
			fcstab(93) = &H81CD
			fcstab(94) = &H81C7
			fcstab(95) = &H1C2
			fcstab(96) = &H140
			fcstab(97) = &H8145
			fcstab(98) = &H814F
			fcstab(99) = &H14A
			fcstab(100) = &H815B
			fcstab(101) = &H15E
			fcstab(102) = &H154
			fcstab(103) = &H8151
			fcstab(104) = &H8173
			fcstab(105) = &H176
			fcstab(106) = &H17C
			fcstab(107) = &H8179
			fcstab(108) = &H168
			fcstab(109) = &H816D
			fcstab(110) = &H8167
			fcstab(111) = &H162
			fcstab(112) = &H8123
			fcstab(113) = &H126
			fcstab(114) = &H12C
			fcstab(115) = &H8129
			fcstab(116) = &H138
			fcstab(117) = &H813D
			fcstab(118) = &H8137
			fcstab(119) = &H132
			fcstab(120) = &H110
			fcstab(121) = &H8115
			fcstab(122) = &H811F
			fcstab(123) = &H11A
			fcstab(124) = &H810B
			fcstab(125) = &H10E
			fcstab(126) = &H104
			fcstab(127) = &H8101
			fcstab(128) = &H8303
			fcstab(129) = &H306
			fcstab(130) = &H30C
			fcstab(131) = &H8309
			fcstab(132) = &H318
			fcstab(133) = &H831D
			fcstab(134) = &H8317
			fcstab(135) = &H312
			fcstab(136) = &H330
			fcstab(137) = &H8335
			fcstab(138) = &H833F
			fcstab(139) = &H33A
			fcstab(140) = &H832B
			fcstab(141) = &H32E
			fcstab(142) = &H324
			fcstab(143) = &H8321
			fcstab(144) = &H360
			fcstab(145) = &H8365
			fcstab(146) = &H836F
			fcstab(147) = &H36A
			fcstab(148) = &H837B
			fcstab(149) = &H37E
			fcstab(150) = &H374
			fcstab(151) = &H8371
			fcstab(152) = &H8353
			fcstab(153) = &H356
			fcstab(154) = &H35C
			fcstab(155) = &H8359
			fcstab(156) = &H348
			fcstab(157) = &H834D
			fcstab(158) = &H8347
			fcstab(159) = &H342
			fcstab(160) = &H3C0
			fcstab(161) = &H83C5
			fcstab(162) = &H83CF
			fcstab(163) = &H3CA
			fcstab(164) = &H83DB
			fcstab(165) = &H3DE
			fcstab(166) = &H3D4
			fcstab(167) = &H83D1
			fcstab(168) = &H83F3
			fcstab(169) = &H3F6
			fcstab(170) = &H3FC
			fcstab(171) = &H83F9
			fcstab(172) = &H3E8
			fcstab(173) = &H83ED
			fcstab(174) = &H83E7
			fcstab(175) = &H3E2
			fcstab(176) = &H83A3
			fcstab(177) = &H3A6
			fcstab(178) = &H3AC
			fcstab(179) = &H83A9
			fcstab(180) = &H3B8
			fcstab(181) = &H83BD
			fcstab(182) = &H83B7
			fcstab(183) = &H3B2
			fcstab(184) = &H390
			fcstab(185) = &H8395
			fcstab(186) = &H839F
			fcstab(187) = &H39A
			fcstab(188) = &H838B
			fcstab(189) = &H38E
			fcstab(190) = &H384
			fcstab(191) = &H8381
			fcstab(192) = &H280
			fcstab(193) = &H8285
			fcstab(194) = &H828F
			fcstab(195) = &H28A
			fcstab(196) = &H829B
			fcstab(197) = &H29E
			fcstab(198) = &H294
			fcstab(199) = &H8291
			fcstab(200) = &H82B3
			fcstab(201) = &H2B6
			fcstab(202) = &H2BC
			fcstab(203) = &H82B9
			fcstab(204) = &H2A8
			fcstab(205) = &H82AD
			fcstab(206) = &H82A7
			fcstab(207) = &H2A2
			fcstab(208) = &H82E3
			fcstab(209) = &H2E6
			fcstab(210) = &H2EC
			fcstab(211) = &H82E9
			fcstab(212) = &H2F8
			fcstab(213) = &H82FD
			fcstab(214) = &H82F7
			fcstab(215) = &H2F2
			fcstab(216) = &H2D0
			fcstab(217) = &H82D5
			fcstab(218) = &H82DF
			fcstab(219) = &H2DA
			fcstab(220) = &H82CB
			fcstab(221) = &H2CE
			fcstab(222) = &H2C4
			fcstab(223) = &H82C1
			fcstab(224) = &H8243
			fcstab(225) = &H246
			fcstab(226) = &H24C
			fcstab(227) = &H8249
			fcstab(228) = &H258
			fcstab(229) = &H825D
			fcstab(230) = &H8257
			fcstab(231) = &H252
			fcstab(232) = &H270
			fcstab(233) = &H8275
			fcstab(234) = &H827F
			fcstab(235) = &H27A
			fcstab(236) = &H826B
			fcstab(237) = &H26E
			fcstab(238) = &H264
			fcstab(239) = &H8261
			fcstab(240) = &H220
			fcstab(241) = &H8225
			fcstab(242) = &H822F
			fcstab(243) = &H22A
			fcstab(244) = &H823B
			fcstab(245) = &H23E
			fcstab(246) = &H234
			fcstab(247) = &H8231
			fcstab(248) = &H8213
			fcstab(249) = &H216
			fcstab(250) = &H21C
			fcstab(251) = &H8219
			fcstab(252) = &H208
			fcstab(253) = &H820D
			fcstab(254) = &H8207
			fcstab(255) = &H202
		End If
		
		' The initial FCS value
		fcs = 0 'pppinitfcs16
		
		' evaluate the FCS
		For I = LBound(cp) To LBound(cp) + Size - 1
			'      L1 = fcs \ 256
			'      L2 = L1 Xor cp(I)
			'      L3 = fcstab(L2)
			'      L4 = fcs * 256
			'      L5 = L3 Xor L4
			'      L6 = L5 And &HFFFF&
			'      Debug.Print "I="; I; ", cp(i)="; Hex(cp(I)); "("; Chr(cp(I)); "), ";
			'      Debug.Print "fcs_in="; Hex(fcs); ", L1="; Hex(L1); ", L2="; Hex(L2);
			'      Debug.Print ", L3="; Hex(L3); ", L4="; Hex(L4); ", L5="; Hex(L5);
			'      Debug.Print ", L6="; Hex(L6);
			fcs = (fcstab((fcs \ &H100) Xor cp(I)) Xor (fcs * 256)) And &HFFFF
			'      Debug.Print ", fcs_out="; Hex(fcs)
		Next I
		
		' return the result
		Crc16 = fcs
	End Function
End Module