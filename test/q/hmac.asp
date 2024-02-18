<!-- #include file="./include/dvim_brix_crypto-js-master_VB.asp" -->

<%

Function generateHmac(url, method, querystring, apikey, seckey)
    dim signedDate
    dim encString
    dim requestData

    '[timestamp, httpMethod, requestPath, queryString]
    signedDate = generateSignedDate()
    requestData = signedDate & method & url & querystring
    set encString = mac256(requestData, seckey)
    signature = ", signature=" & cstr(encString)

    generateHmac = "CEA algorithm=HmacSHA256, access-key=" & apikey & ", signed-date=" & signedDate & signature
End Function

'/******  MAC Function ******/
' CEA algorithm=HmacSHA256, access-key=0a3a0f34-7852-4ad8-9368-766290b8b1ab, signed-date=190201T042152Z, signature=737374df2de01ad31cc85c14c42faa0711e7d156d8350e57c911b20916d4ba77
'Input String|WordArray , Returns WordArray
Function mac256(ent, seckey)
    Dim encWA
    Set encWA = ConvertUtf8StrToWordArray(ent)
    Dim keyWA
    Set keyWA = ConvertUtf8StrToWordArray(seckey)
    Dim resWA
    Set resWA = CryptoJS.HmacSHA256(encWA, keyWA)
    Set mac256 = resWA
End Function

'Input (Utf8)String|WordArray Returns WordArray
Function ConvertUtf8StrToWordArray(data)
    If (typename(data) = "String") Then
        Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data)
    Elseif (typename(data) = "JScriptTypeInfo") Then
        On error resume next
        'Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data.toString(CryptoJS.enc.Utf8))
        Set ConvertUtf8StrToWordArray = CryptoJS.lib.WordArray.create().concat(data) 'Just assert that data is WordArray
        If Err.number>0 Then
            Set ConvertUtf8StrToWordArray = Nothing
        End if
        On error goto 0
    Else
        Set ConvertUtf8StrToWordArray = Nothing
    End if
End Function

Function generateSignedDate()
    Dim nowDateTime

    'ISO TIMEZONE
    nowDateTime = DateAdd("H", -9, now())

    generateSignedDate = ToIsoDate(nowDateTime) & "T" & ToIsoTime(nowDateTime) & "Z"

End Function

Function ToIsoDate(datetime)
    ToIsoDate = CStr(Mid(Year(datetime), 3, 2)) & StrN2(Month(datetime)) & StrN2(Day(datetime))
End Function

Function ToIsoTime(datetime)
    ToIsoTime = StrN2(Hour(datetime)) & StrN2(Minute(datetime)) & StrN2(Second(datetime))
End Function

Function StrN2(n)
    If Len(CStr(n)) < 2 Then StrN2 = "0" & n Else StrN2 = n
End Function
%>
