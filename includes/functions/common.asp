<%
Private Function hexValue(ByVal intHexLength)

	Dim intLoopCounter
	Dim strHexValue

	'Randomise the system timer
	Randomize Timer()

	'Generate a hex value
	For intLoopCounter = 1 to intHexLength

		'Genreate a radom decimal value form 0 to 15
		intHexLength = CInt(Rnd * 1000) Mod 16

		'Turn the number into a hex value
		Select Case intHexLength
			Case 1
				strHexValue = "1"
			Case 2
				strHexValue = "2"
			Case 3
				strHexValue = "3"
			Case 4
				strHexValue = "4"
			Case 5
				strHexValue = "5"
			Case 6
				strHexValue = "6"
			Case 7
				strHexValue = "7"
			Case 8
				strHexValue = "8"
			Case 9
				strHexValue = "9"
			Case 10
				strHexValue = "A"
			Case 11
				strHexValue = "B"
			Case 12
				strHexValue = "C"
			Case 13
				strHexValue = "D"
			Case 14
				strHexValue = "E"
			Case 15
				strHexValue = "F"
			Case Else
				strHexValue = "Z"
		End Select

		'Place the hex value into the return string
		hexValue = hexValue & strHexValue
	Next
End Function

Function RoundNumber(intNumber)
	intNumber = csng(Replace(intNumber, ".", strServerComma))
	
	if len(strRoundNumber) > 0 then
		roundtemp = Round(intNumber, strRoundNumber)
		comma_place = instr(roundtemp, strServerComma)
		
		if comma_place > 0 then
			x = 0
			for x = comma_place + 1 to (comma_place + strRoundNumber)
				if NOT isnumeric(mid(roundtemp, x, 1)) then
					roundtemp = roundtemp & "0"
				end if
			next
		else
			roundtemp = roundtemp & ","
			x = 0
			for x = 1 to strRoundNumber
				roundtemp = roundtemp & "0"
			next
		end if
	else
		roundtemp = Round(intNumber, 0)
	end if
	RoundNumber = roundtemp
	roundtemp = Replace(RoundNumber, strServerComma, strDecimalSign)
	
	if len(strSeparator) > 0 AND instr(roundtemp, strDecimalSign) > 0 then
		before_comma = getComma(left(roundtemp, instr(roundtemp, strDecimalSign) - 1))
		roundtemp = before_comma & mid(roundtemp, instr(roundtemp, strDecimalSign))
	end if
	RoundNumber = roundtemp
end function

function convertPost(postMsg)
	postMsg = Replace(postMsg, "<", "&lt;")
	postMsg = Replace(postMsg, ">", "&gt;")
	
	postMsg = Replace(postMsg, chr(10), "<br />")
	
	postMsg = Replace(postMsg, "[B]", "<b>")
	postMsg = Replace(postMsg, "[/B]", "</b>")
	postMsg = Replace(postMsg, "[U]", "<u>")
	postMsg = Replace(postMsg, "[/U]", "</u>")
	postMsg = Replace(postMsg, "[I]", "<i>")
	postMsg = Replace(postMsg, "[/I]", "</i>")
	postMsg = Replace(postMsg, "[LEFT]", "<div align=""left"">")
	postMsg = Replace(postMsg, "[RIGHT]", "<div align=""right"">")
	postMsg = Replace(postMsg, "[CENTER]", "<div align=""center"">")
	postMsg = Replace(postMsg, "[/LEFT]", "</div>")
	postMsg = Replace(postMsg, "[/CENTER]", "</div>")
	postMsg = Replace(postMsg, "[/RIGHT]", "</div>")
	
	do while instr(postMsg, "[HREF") > 0
		intBegin = instr(postMsg, "[HREF")
		intEnd   = instr(postMsg, "[/HREF]") + 7
		strBefore = left(postMsg, intBegin - 1)
		strAfter  = right(postMsg, len(postMsg) - (intEnd - 1))
		
		temp = Mid(postMsg, intBegin, intEnd - intBegin)
		temp = left(temp, len(temp) - 7)
		
		'Find URL
		intBeginUrl = instr(temp, "[HREF")
		intEnd = instr(temp, "]")
		url    = Mid(temp, intBeginUrl + 6, intEnd - (intBeginUrl + 6))
		temp   = Right(temp, len(temp) - intEnd)
		
		postMsg = strBefore & "<a href=""" & url & """>" & temp & "</a>" & strAfter
	loop
	
	convertPost = postMsg
end function

function getComma(intNumber)
	if strSeparator = "&nbsp;" then strSeparator = " "
	if len(intNumber) > 0 and isnumeric(intNumber) then
		length = len(intNumber)
		if length / 3 = length \ 3 then
			total_comma = (length \ 3) - 1
		else
			total_comma = length \ 3
		end if
			
		y = 0
		newNumber = ""
		if total_comma > 0 then
			for y = 1 to total_comma
				left_side = left(intNumber, length - (y*3))
				right_side = right(intNumber, ((y * 3) + (y - 1)))
				intNumber = left_side & strSeparator & right_side
			next
			newNumber = newNumber & intNumber
		end if		
		getComma = intNumber
	end if
end function
%>