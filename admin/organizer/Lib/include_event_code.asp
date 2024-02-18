<%
	'####### 이벤트 번호 #######
	Dim vEventCode
	IF application("Svr_Info") = "Dev" THEN
		vEventCode = "20569"
	Else
		vEventCode = "22155"
	End If
	'####### 이벤트 번호 #######
%>