//匹配字符串,返回|||分隔的结果

	Function FindMatch(ByVal str, ByVal start, ByVal last)
		Dim Match
		Dim s
		Dim FilterStr
		Dim MatchStr
		Dim strContent
		Dim ArrayFilter()
		Dim i, n
		Dim bRepeat
		
		If Len(start) = 0 Or Len(last) = 0 Then Exit Function
		
		On Error Resume Next
		
		MatchStr = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(last) & ")"
		
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = MatchStr
		Set s = re.Execute(str)
		n = 0
		For Each Match In s
			If n = 0 Then
				n = n + 1
				ReDim ArrayFilter(n)
				ArrayFilter(n) = Match
			Else
				bRepeat = False
				For i = 0 To UBound(ArrayFilter)
					If UCase(Match) = UCase(ArrayFilter(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve ArrayFilter(n)
					ArrayFilter(n) = Match
				End If
			End If
		Next
		
		Set s = Nothing
		Set re = Nothing
		
		strContent = Join(ArrayFilter, "|||")
		strContent = Replace(strContent, start, "")
		strContent = Replace(strContent, last, "")
		
		FindMatch = Replace(strContent, "|||", vbNullString, 1, 1)
		Exit Function
	End Function
	
	
	Function CorrectPattern(ByVal str)
		str = Replace(str, "\\", "\\\\")
		str = Replace(str, "~", "\~")
		str = Replace(str, "!", "\!")
		str = Replace(str, "@", "\@")
		str = Replace(str, "#", "\#")
		str = Replace(str, "%", "\%")
		str = Replace(str, "^", "\^")
		str = Replace(str, "&", "\&")
		str = Replace(str, "*", "\*")
		str = Replace(str, "(", "\(")
		str = Replace(str, ")", "\)")
		str = Replace(str, "-", "\-")
		str = Replace(str, "+", "\+")
		str = Replace(str, "[", "\[")
		str = Replace(str, "]", "\]")
		str = Replace(str, "<", "\<")
		str = Replace(str, ">", "\>")
		str = Replace(str, ".", "\.")
		str = Replace(str, "/", "\/")
		str = Replace(str, "?", "\?")
		str = Replace(str, "=", "\=")
		str = Replace(str, "|", "\|")
		str = Replace(str, "$", "\$")
		CorrectPattern = str
	End Function