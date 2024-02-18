<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<%
Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput
End Function

function tenDecW(byVal sPlainEnc)
    dim lLength, lCount, sTemp, buf
    dim mul, mul1, mul2, mul3
    lLength = Len(sPlainEnc)
    For lCount = 1 To lLength
        mul1 = lCount mod 9
        mul2 = lCount mod 4
        mul3 = lCount mod 30
        if (lCount mod 2)=0 then mul3 = 30-mul3
        mul  = mul1-mul2+mul3
        buf = AscW(Mid(sPlainEnc, lCount, 1))-mul
        sTemp = sTemp & ChrW(buf)
    Next

    tenDecW = HexToStrW(sTemp)
end function

Function HexToStrW(byVal strHex)
    Dim Length
    Dim Max
    Dim Str
    Max = Len(strHex)
    For Length = 1 To Max Step 4
        Str = Str & ChrW("&h" & Mid(strHex, Length, 4))
    Next
    HexToStrW = Str
End function

dim sqlStr, arrList, intLoop

sqlStr = "select top 2000 * from db_temp.dbo.tbl_rejectMail_T where email is NULL"
rsget.open sqlStr,dbget,1

if  not rsget.EOF  then
    arrList = rsget.getRows()
end if
rsget.Close

dim buf, pos
dim date1,time1,svr,urI,email
IF isArray(arrList) THEN
    For intLoop = 0 To UBound(arrList,2)
        'response.write arrList(0,intLoop)&","&arrList(1,intLoop)&","&arrList(2,intLoop)&","&URLDecode(arrList(3,intLoop))&"<br>"
        date1 = arrList(0,intLoop)
        time1 = arrList(1,intLoop)
        svr = arrList(2,intLoop)
        urI = arrList(3,intLoop)
        
        buf = arrList(3,intLoop)
        buf = replace(buf,"um=","")
        pos = InStr(buf,"&_=")-1
        buf = LEFT(buf,pos)
        email = tenDecW(URLDecode(buf))
        
        response.write date1&","&time1&","&svr&","&urI&","&email&"<br>"
        sqlStr="update db_temp.dbo.tbl_rejectMail_T"
        sqlStr=sqlStr&" set email='"&email&"'" &VbCRLF
        sqlStr=sqlStr&" where date1='"&date1&"'"&VbCRLF
        sqlStr=sqlStr&" and time1='"&time1&"'"&VbCRLF
        sqlStr=sqlStr&" and svr='"&svr&"'"&VbCRLF
        sqlStr=sqlStr&" and urI='"&urI&"'"&VbCRLF
        
       ' dbget.Execute sqlStr
    next
end if        
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
