Dim i As Integer
Dim j As Integer
Dim r As integer
Dim c As Integer
Dim tmp As Variant

Private sub append(ByRef targetArray As Variant, argv As Variant _
				   Optional method As variant = "defalt_vertical")
	Dim targetDimension As Integer
	targetDimension = getArrayDimension(targetArray)
	If method = "v" Or "vertical" then
	End if
	If Not IsArray(argv) Then
		Select Case targetDimension		' case of target dimension.
			Case 1
				ReDim Preserve targetarray(LBound(targetArray) To UBound(targetArray) + 1)
				targetArray(UBound(targetArray)) = argv
			case 2
				ReDim Preserve targetarray(UBound(targetArray, 2), _
										   LBound(targetArray, 1) To UBound(targetArray, 1) + 1)
				targetArray(UBound(targetArray, 2), UBound(targetArray, 1)) = argv
			Case Else
				Err.raise 1000, "append()", "Too many array dimension. Must be less than 2."
		End Select

		Exit Sub

	Else								' if argv is array.
		Dim argvDimension As Integer
		Dim argv_nRow As Long
		Dim argv_nColumn As long
		Dim target_nRow As Long
		Dim target_nColumn As Long
		Dim argvArray_ As Variant
		Dim targetArray_ As variant

		argvDimension = getArrayDimension(argv)
		If argvDimension = 1 Then
			argv_nRow = 1
			argv_nColumn = UBound(argv) - lBound(argv)
			argvpArray_ = toTwoDimensionArray(argv)
		Else If argvDimension = 2 Then
			argv_nRow = UBound(argv, 1) - lbound(argv, 1)
			argv_nColumn = UBound(argv, 2) - lbound(argv, 2)
			argvArray_ = argvArray
		End If
		If targetDimension = 1 Then
			target_nRow = 1
			target_nColumn = UBound(target) - lBound(target)
			targetArray_ = toTwoDimensionArray(targetArray)
		Else If targetDimension = 2 Then
			target_nRow = UBound(target, 1) - lbound(target, 1)
			target_nColumn = UBound(target, 2) - lbound(target, 2)
			targetArray_ = targetArray
		End If

		If method = "defalt_vertical" Then
			Err.raise 1000, "append()", "If argv is array then option ""method"" must be specified."
			Exit Sub
		Else If method = "v" Or method = "vertical" Then
			If argv_nColumn <> target_nColumn Then
				Err.raise 1000, "append()", "Two arrays must be same number of Column."
				Exit Sub
			End If

			Dim tmpArray() As Variant
			ReDim Preserve tmpArray(lbound(targetArray_, 1) To ubound(targetArray_, 1), _
									LBound(targetArray_, 2) To UBound(targetArray_, 2) + target_nRow)

			For c = lbound(targetArray_, 2) To ubound(targetArray_, 2)
				r = LBound(targetArray_, 1)
				For j = lbound(targetArray_, 1) To ubound(targetArray_, 1)
					tmpArray(r, c) = targetArray_(i, j)
					r = r + 1
				Next

				r = LBound(targetArray_, 1) + 1
				For j = lbound(targetArray_, 2) To ubound(targetArray_, 2)
					tmpArray(r, c) = argv_(i, j)
					r = r + 1
				Next
			Next

		Else If method = "h" Or "horizontanl" Then
			If argv_nColumn <> target_nColumn Then
				Err.raise 1000, "append()", "Two arrays must be same number of row."
				Exit Sub
			End If
			Dim tmpArray() As Variant
			ReDim Preserve tmpArray(lbound(targetArray_, 1) To ubound(targetArray_, 1) + target_nColumn, _
									LBound(targetArray_, 2) To UBound(targetArray_, 2))

			For r = lbound(targetArray_, 1) To ubound(targetArray_, 1)
				c = LBound(targetArray_, 1)
				For j = lbound(targetArray_, 2) To ubound(targetArray_, 2)
					tmpArray(r, c) = targetArray_(i, j)
					c = c + 1
				Next

				c = uBound(targetArray_, 1) + 1
				For j = lbound(argvArray_, 2) To ubound(targetArray_, 2)
					tmpArray(r, c) = argv_(i, j)
					c = c + 1
				Next
			next
		Else
			Err.raise 1000, "append()", "Option ""method"" must be selected from ""v(vertical)"", ""h(horizontarl)""."
		End If
		Erase targetArray
		targetArray = tmpArray
End Sub

Public Function getArrayDimension(ByRef targetArray As Variant) As Integer
  If Not IsArray(targetArray) then
	  getArrayDimension = False
	  Exit Function
  End If

  Dim n As Long
  n = 0
  Dim tmp As Long
  On Error Resume Next
  Do While Err.Number = 0
    n = n + 1
    tmp = UBound(targetArray, n)
  Loop
  Err.Clear
  getArrayDimension = n - 1
End Function

Public Function toOneDimensionArray(ByRef targetArray() As variant, rowIdx As Integer) as variant
	Dim tmpArray() As Variant
	ReDim tmpArray(lbound(targetArray, 2), ubound(targetArray,2))
	for i = lbound(tmpArray) to ubound(tmpArray) step 1
		tmpArray(i) = targetArray(rowIdx, i)
	next i
	toOneDimensionArray = tmpArray
end Function

Public Function toTwoDimensionArray(ByRef targetArray() As Variant) As Variant
	Dim tmpArray() As Variant
	ReDim tmpArray(lbound(targetArray) To ubound(targetArray), 0)
	For i = LBound(targetArray) To ubound(targetArray)
		tmpArray(i, 0) = targetArray(i)
	Next
	toTwoDimensionArray = tmpArray
End Function

Public function transportArray(Byref targetArray() As Variant) as Variant
	Dim tmpArray() As Variant
	Dim dimention As Integer
	dimention = getArrayDimension(targetArray)
	If  dimention = 1 Then
		reDim tmpArray(0 To LBound(targetArray) - ubound(targetArray), 0)
		For i = 0 To ubound(targetArray) Step 1
			tmpArray(i, 0) = targetArray(i)
		Next
	Else If dimention = 2 then
		ReDim tmpArray(0 To LBound(targetArray, 2) - ubound(targetArray, 2),_
					   0 To LBound(targetArray, 1) - ubound(targetArray, 1))
		r = 0
		c = 0
		For i = lbound(targetArray, 2) To ubound(targetArray, 2)
			for j = LBound(targetArray, 1) to ubound(targetArray, 2)
				tmpArray(r, c) = targetArray(i, j)
				c = c + 1
			next j
			r = r + 1
		Next i
	Else
		Err.raise 1000, "transportArray", "To many dimension of array."
	End if
	transportArray = tmpArray
End function
