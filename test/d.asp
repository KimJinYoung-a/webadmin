<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="/webfonts/CoreSansC.css">
</head>
<body>
	<%
	function AstarPhoneNumber(phoneNumber)
		Dim regEx, result
		Set regEx = New RegExp

		With regEx
			.Pattern = "([0-9]+)-([0-9]+)-([0-9]+)"
			.IgnoreCase = True
			.Global = True
		End With

		result = regEx.Replace(phoneNumber,"$1-***-$3")

		if (result = phoneNumber) then
			if (Len(phoneNumber) >= 4) then
				result = Left(phoneNumber, (Len(phoneNumber) - 4)) & "****"

			end if
		end if

		set regEx = nothing

		AstarPhoneNumber = result
	end function

	function AstarUserName(userName)
		Dim result

		Select Case Len(userName)
			Case 0
				''
			Case 1
				result = "*"
			Case 2
				result = Left(userName,1) & "*"
			Case Else
				''3�̻�
				result = Left(userName,1) & "*" & Right(userName,1)
		End Select

		AstarUserName = result
	end function

	dim aaa, bbb
	aaa = "010-1111-3333"
	bbb = "ȫ�浿"
	Response.Write "AaBbCcDdEeFfGg<br/>"
	Response.write AstarPhoneNumber(aaa)
	Response.write AstarUserName(bbb) & "<br/>"


	%>
	<p style="font-family:CoreSansC-65Bold;">
		AaBbCcDdEeFfGg<br/>
		<%= Left("����Ư���� ���ϱ� ��õ��101�� 2",35) %>
	</p>
	<p style="font-family:CoreSansC-45Regular;">
	AaBbCcDdEeFfGg<br/>
		>>>> <%= CLng(Len("����Ư���� ���ϱ� ��õ��1002a")/2) %>
	</p>
</body>
</html>
