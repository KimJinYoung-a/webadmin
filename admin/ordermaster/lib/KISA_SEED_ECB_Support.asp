<%

'// 문자열 -> 16진수
'// 12345abcde한글 => 31,32,33,34,35,61,62,63,64,65,D5,5C,AE,00
Public Function StringToHex(pStr)
    Dim i, one_hex, retVal, val

	For i = 1 To Len(pStr)
		val = Mid(pStr, i, 1)

		if (AscW(val) < 0) then
			val = Hex(AscW(val))
			val = Right("0000" & val, 4)
			val = Left(val, 2) & "," & Right(val,2)
		else
			val = Right("00" & Hex(Asc(val)), 2)
		end if
		retVal = retVal & val & ","
    Next

	if (retVal <> "") then
		retVal = Left(retVal, Len(retVal)-1)
	end if

    StringToHex = retVal
End Function

'// 16진수 -> 문자열
'// 31,32,33,34,35,61,62,63,64,65,D5,5C,AE,00 => 12345abcde한글
Public Function HexToString(pHex)
    Dim one_hex, tmp_hex, i, retVal, val

	pHex = Replace(pHex, ",", "")
    For i = 1 To Len(pHex)
        one_hex = Mid(pHex, i, 1)
        If IsNumeric(one_hex) Then
            If CInt(one_hex) < 8 Then
                tmp_hex = Mid(pHex, i, 2)
                i = i + 1
            Else
                tmp_hex = Mid(pHex, i, 4)
                i = i + 3
            End If
        Else
            tmp_hex = Mid(pHex, i, 4)
            i = i + 3
        End If
        retVal = retVal & ChrW("&H" & tmp_hex)
    Next
    HexToString = retVal
End Function

'// 31,32,33,34,35,61,62,63,64,65,D5,5C,AE,00 => Array
Public Function HexStringToArray(hStr)
	dim result(), i

	hStr = split(hStr,",")
	redim result(ubound(hStr))

	for i=0 to (ubound(hStr))
		result(i) = (Cbyte)("&H" & (right("0000" & hStr(i), 4)))
	Next
	HexStringToArray = result
end function

'// Array => 31,32,33,34,35,61,62,63,64,65,D5,5C,AE,00
Public Function HexArrayToString(hArr)
	dim result, i

	for i = 0 to (UBound(hArr))
		result = result & Right("00" & hex(hArr(i)), 2) & ","
	next
	if (result <> "") then
		result = Left(result, Len(result)-1)
	end if
	HexArrayToString = result
end function

function SeedECBEncrypt(key, plainText)
	dim result, i

	result = HexStringToArray(StringToHex(plainText))
	result = KISA_SEED_ECB.SEED_ECB_Encrypt(HexStringToArray(StringToHex(key)), result, 0, UBound(result)+1)
	SeedECBEncrypt = Replace(HexArrayToString(result), ",", "")
end function

function SeedECBDecrypt(key, cipherText)
	dim result, i

	for i = 0 to CLng(Len(cipherText) / 2) - 1
		result = result & "," & mid(cipherText, (i*2)+1, 2)
	next

	if (result <> "") then
		result = Mid(result, 2, Len(result))
	end if

	result = HexStringToArray(result)

	result = KISA_SEED_ECB.SEED_ECB_Decrypt(HexStringToArray(StringToHex(key)), result, 0, UBound(result)+1)
	SeedECBDecrypt = HexToString(HexArrayToString(result))
end function

''dim aaa : aaa = "12345abcde한글"
''aaa = SeedECBEncrypt("4e161da3b4b3c7fa", aaa)
''response.write aaa & "<br />"
''aaa = SeedECBDecrypt("4e161da3b4b3c7fa", aaa)
''response.write aaa & "<br />"

%>
