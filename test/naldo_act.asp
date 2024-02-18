<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<script language="javascript" runat="server" src="/js/json2.js"></script>
<%

Function toUTF8(szSource)
    Dim szChar
    Dim WideChar
    Dim nLength
    Dim i

    nLength = Len(szSource)
    For i = 1 To nLength
        szChar = Mid(szSource, i, 1)

        If Asc(szChar) < 0 Then
            WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))

            If (WideChar And &HFF80) = 0 Then
                toUTF8 = toUTF8 & "%" & Hex(WideChar)
            ElseIf (WideChar And &HF000) = 0 Then
                toUTF8 = toUTF8 & _
                "%" & Hex(CLng((WideChar And &HFFC0) / 64) Or &HC0) & _
                "%" & Hex(WideChar And &H3F Or &H80)
            Else
                toUTF8 = toUTF8 & _
                "%" & Hex(CLng((WideChar And &HF000) / 4096) Or &HE0) & _
                "%" & Hex(CLng((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
                "%" & Hex(WideChar And &H3F Or &H80)
            End If
        Else
            toUTF8 = toUTF8 + szChar
        End If
    Next
    Exit Function
End Function

Function json_encode(ByVal dic)
	dim ret, k, x

    ret = "{"
    If TypeName(dic) = "Dictionary" Then
        For each k in dic
			Select Case VarType(dic.Item(k))
				Case vbString
					ret = ret & """" & k & """:""" & dic.Item(k) & ""","
				Case Else
					If VarType(dic.Item(k)) > vbArray Then
						ret = ret & """" & k & """:["
						For x = 0 to Ubound(dic.Item(k), 1)
							ret = ret & """" & dic.Item(k)(x) & ""","
						Next
						ret = Left(ret, Len(ret) - 1)   'Trim trailing comma
						ret = ret & "],"
					Else
						ret = ret & """" & k & """:" & dic.Item(k) & ","
					End If
			End Select
		Next
		ret = Left(ret, Len(ret) - 1)   'Trim trailing comma
	End If
	ret = ret & "}"
	json_encode = ret
End Function

Function make_url(ByVal dic)
	dim ret, k

	ret = "?"
	If TypeName(dic) = "Dictionary" Then
		For each k in dic
			ret = ret & k & "=" & toUTF8(dic.Item(k)) & "&"
		next
	end if

	make_url = ret
End Function


dim mode
dim vURL_LOGIN 			: vURL_LOGIN = "http://1.234.83.82:8080/open/api/authenticate?username=user&password=user"		'// 로그인
''dim vURL_CHECK_LOGIN 	: vURL_CHECK_LOGIN = "http://1.234.83.82:8080/open/api/authenticated"							'// 로그인 체크
dim vURL_CHECK_PRICE 	: vURL_CHECK_PRICE = "http://1.234.83.82:8080/open/rest/price"									'// 가격조회
dim vURL 				: vURL = "http://1.234.83.82:8080/open/rest/order"												'// 주문접수, 주문목록조회, 조회, 취소

dim jsonString, vAnswer, vStatus, dict, objJSON, objOrder
dim urlString
dim i, j

dim xmlhttp, authToken


'// ============================================================================
mode = request("mode")


Select Case mode
	Case "login"
		'// 로그인
		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "POST", vURL_LOGIN, false
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.send ""

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		set xmlhttp = Nothing

		if (vStatus = 200) then
			'// 로그인 성공
			session("naldo_token") = vAnswer
			Response.Write "<script>alert('로그인에 성공하였습니다.'); opener.focus(); window.close();</script>"
		else
			Response.Write "<script>alert('시스템팀 문의\n\n로그인 실패!! - " & vStatus & "');</script>"
			Response.Write "시스템팀 문의\n\n로그인 실패!! - " & vStatus
		end if
	Case "checkprice"
		''// 가격조회
		if (session("naldo_token") = "") then
			Response.Write "<script>alert('먼저 로그인하세요.');</script>"
			Response.end
		end if

		set dict = Server.CreateObject("Scripting.Dictionary")

		dict.Add "fromSido", request("fromSido")
		dict.Add "fromGugun", request("fromGugun")
		dict.Add "fromDong", request("fromDong")
		dict.Add "fromDetail", request("fromDetail")
		dict.Add "fromAddressType", request("fromAddressType")

		dict.Add "toSido", request("toSido")
		dict.Add "toGugun", request("toGugun")
		dict.Add "toDong", request("toDong")
		dict.Add "toDetail", request("toDetail")
		dict.Add "toAddressType", request("toAddressType")

		dict.Add "smallCount", request("smallCount")
		dict.Add "mediumCount", request("mediumCount")
		dict.Add "bigCount", request("bigCount")
		dict.Add "complexCount", request("complexCount")

		dict.Add "billType", request("billType")

		dict.Add "reservation", request("reservation")
		dict.Add "reservationTime", request("reservationTime")

		dict.Add "options", request("options")

		urlString = make_url(dict)

		set dict = Nothing

		authToken = JSON.parse(session("naldo_token")).token

		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "POST", vURL_CHECK_PRICE & urlString, false
		xmlhttp.setRequestHeader "Content-type","application/json"
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.setRequestHeader "X-Auth-Token", authToken
		xmlhttp.send ""

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		if (vStatus = 200) then
			Response.write vAnswer
		elseif (vStatus = 401) then
			Response.write "인증오류 : 다시 로그인하세요."
		elseif (vStatus = 403) then
			Response.write "오류 : " & JSON.parse(vAnswer).message
		else
			Response.write "ERR : " & vStatus & "<br /><br />"
			Response.write vAnswer
		end if

		set xmlhttp = Nothing
	Case "sendorder"
		''// 주문접수
		if (session("naldo_token") = "") then
			Response.Write "<script>alert('먼저 로그인하세요.');</script>"
			Response.end
		end if

		set dict = Server.CreateObject("Scripting.Dictionary")

		dict.Add "orderPhoneNumber", request("orderPhoneNumber")
		dict.Add "senderPhoneNumber", request("senderPhoneNumber")
		dict.Add "receiverPhoneNumber", request("receiverPhoneNumber")

		dict.Add "receiverName", request("receiverName")
		dict.Add "senderName", request("senderName")

		dict.Add "etc", request("etc")

		dict.Add "company", request("company")

		dict.Add "smsForward", request("smsForward")
		dict.Add "smsTarget", request("smsTarget")

		dict.Add "fromSido", request("fromSido")
		dict.Add "fromGugun", request("fromGugun")
		dict.Add "fromDong", request("fromDong")
		dict.Add "fromDetail", request("fromDetail")
		dict.Add "fromAddressType", request("fromAddressType")

		dict.Add "toSido", request("toSido")
		dict.Add "toGugun", request("toGugun")
		dict.Add "toDong", request("toDong")
		dict.Add "toDetail", request("toDetail")
		dict.Add "toAddressType", request("toAddressType")

		dict.Add "smallCount", request("smallCount")
		dict.Add "mediumCount", request("mediumCount")
		dict.Add "bigCount", request("bigCount")
		dict.Add "complexCount", request("complexCount")

		dict.Add "billType", request("billType")

		dict.Add "reservation", request("reservation")
		dict.Add "reservationTime", request("reservationTime")

		dict.Add "options", request("options")

		jsonString = json_encode(dict)

		set dict = Nothing

		authToken = JSON.parse(session("naldo_token")).token

		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "POST", vURL, false
		xmlhttp.setRequestHeader "Content-type","application/json"
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.setRequestHeader "X-Auth-Token", authToken
		xmlhttp.send jsonString

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		if (vStatus = 200) then
			Response.write vAnswer
		elseif (vStatus = 401) then
			Response.write "인증오류 : 다시 로그인하세요."
		elseif (vStatus = 403) then
			Response.write "오류 : " & JSON.parse(vAnswer).message
		else
			Response.write "ERR : " & vStatus & "<br /><br />"
			Response.write vAnswer
		end if

		set xmlhttp = Nothing
	Case "orderlist"
		''// 주문목록 조회
		if (session("naldo_token") = "") then
			Response.Write "<script>alert('먼저 로그인하세요.');</script>"
			Response.end
		end if

		authToken = JSON.parse(session("naldo_token")).token

		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "GET", vURL, false
		xmlhttp.setRequestHeader "Content-type","application/json"
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.setRequestHeader "X-Auth-Token", authToken
		xmlhttp.send ""

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		if (vStatus = 200) then
			set objJSON = JSON.parse(vAnswer)

			Response.write "조회건수 : " & objJSON.ordersSize & "<br /><br />"

			if (objJSON.ordersSize > 0) then
				for each objOrder in objJSON.orders
					Response.write objOrder.orderNumber & "<br />"
				next
			end if

			'' Response.write "<br /><br />"
			Response.write vAnswer
		elseif (vStatus = 401) then
			Response.write "인증오류 : 다시 로그인하세요."
		elseif (vStatus = 403) then
			Response.write "오류 : " & JSON.parse(vAnswer).message
		else
			Response.write "aaaa" & vStatus
		end if

		set xmlhttp = Nothing
	Case "vieworder"
		''// 주문조회
		if (session("naldo_token") = "") then
			Response.Write "<script>alert('먼저 로그인하세요.');</script>"
			Response.end
		end if

		authToken = JSON.parse(session("naldo_token")).token

		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "GET", vURL & "/" & request("orderNumber"), false
		xmlhttp.setRequestHeader "Content-type","application/json"
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.setRequestHeader "X-Auth-Token", authToken
		xmlhttp.send ""

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		if (vStatus = 200) then
			Response.write vAnswer
		elseif (vStatus = 401) then
			Response.write "인증오류 : 다시 로그인하세요."
		elseif (vStatus = 403) then
			Response.write "오류 : " & JSON.parse(vAnswer).message
		else
			Response.write "ERR : " & vStatus & "<br /><br />"
			Response.write vAnswer
		end if

		set xmlhttp = Nothing
	Case "cancelorder"
		''// 주문취소
		if (session("naldo_token") = "") then
			Response.Write "<script>alert('먼저 로그인하세요.');</script>"
			Response.end
		end if

		authToken = JSON.parse(session("naldo_token")).token

		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "PUT", vURL & "/" & request("orderNumber"), false
		xmlhttp.setRequestHeader "Content-type","application/json"
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.setRequestHeader "X-Auth-Token", authToken
		xmlhttp.send ""

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		if (vStatus = 200) then
			Response.write vAnswer
		elseif (vStatus = 401) then
			Response.write "인증오류 : 다시 로그인하세요."
		elseif (vStatus = 403) then
			Response.write "오류 : " & JSON.parse(vAnswer).message
		else
			Response.write "ERR : " & vStatus & "<br /><br />"
			Response.write vAnswer
		end if

		set xmlhttp = Nothing
	Case Else
		''
End Select

%>
