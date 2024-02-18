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

dim mode
dim vURL_LOGIN : vURL_LOGIN = "http://1.234.83.82:8080/open/api/authenticate?username=user&password=user"		'// �α���
dim vURL : vURL = "http://1.234.83.82:8080/open/rest/order"														'// �ֹ�����

dim jsonString, vAnswer, vStatus, dict

dim xmlhttp


'// ============================================================================
mode = request("mode")


Select Case mode
	Case "login"
		'// �α���
		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "POST", vURL_LOGIN, false
		xmlhttp.setRequestHeader "Accept","application/json"
		xmlhttp.send ""

		vAnswer = xmlhttp.responseText
		vStatus = xmlhttp.status

		set xmlhttp = Nothing

		if (vStatus = 200) then
			'// �α��� ����
%>
<script>
if(typeof(Storage) !== "undefined") {
	localStorage.setItem('is.token', '<%= vAnswer %>');
	alert('�α��� ����');
} else {
	alert('�α��� ����\n\n�ٸ� �������� �̿��ϼ���.');
}
</script>
<%
		else
			Response.Write "<script>alert('�ý����� ����\n\n�α��� ����!! - " & vStatus & "');</script>"
		end if
	Case "sendorder"
		''// �ֹ�����
		set dict = Server.CreateObject("Scripting.Dictionary")

		dict.Add "orderPhoneNumber", "010-111-1111"
		dict.Add "senderPhoneNumber", "010-111-2222"
		dict.Add "receiverPhoneNumber", "010-111-3333"

		dict.Add "receiverName", "�޴»��"
		dict.Add "senderName", "�������"

		dict.Add "etc", "��۽� ���ǻ���"

		dict.Add "company", "(��)�ٹ�����"

		dict.Add "smsForward", true
		dict.Add "smsTarget", "010-111-4444"

		dict.Add "fromSido", "�����"
		dict.Add "fromGugun", "���α�"
		dict.Add "fromDong", "���з�12��"
		dict.Add "fromDetail", "31 �������� 5��"
		dict.Add "fromAddressType", "NEW"

		dict.Add "toSido", "�����"
		dict.Add "toGugun", "���۱�"
		dict.Add "toDong", "��3��"
		dict.Add "toDetail", "279-508 ������� 201ȣ"
		dict.Add "toAddressType", "OLD"

		dict.Add "smallCount", 1
		dict.Add "mediumCount", 0
		dict.Add "bigCount", 0
		dict.Add "complexCount", 0

		dict.Add "billType", "CREDIT"

		dict.Add "reservation", true
		dict.Add "reservationTime", "2015-06-11 11:30"

		dict.Add "options", ""

		jsonString = json_encode(dict)

		set dict = Nothing


		set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "POST", vURL, false
		xmlhttp.setRequestHeader "Content-type","application/json"
		xmlhttp.setRequestHeader "Accept","application/json"
		''xmlhttp.setRequestHeader('X-Auth-Token', authToken);
		xmlhttp.send jsonString

		vAnswer = xmlhttp.responseText

		Response.write vAnswer

		set xmlhttp = Nothing
	Case Else
		''
End Select











%>
<html>
	<head>
	</head>
	<body>
		aaa
	</body>
</html>
