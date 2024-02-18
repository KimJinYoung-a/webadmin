<%
'####### 이벤트명이 "2010 다이어리 사은품" 인 이벤트코드 #######
Dim vEventCode, vEventCode2
Dim vGiftCode1,vGiftCode2,vGiftCode3
IF application("Svr_Info") = "Dev" THEN
	'vEventCode = "20719"
	'vEventCode2 = ""
	vGiftCode1=""
Else
	'vEventCode = "29488"
	'vEventCode2 = ""
	vGiftCode1="11501"      ''롯데닷컴 이벤트
	vGiftCode2="11503"
	vGiftCode3="11504"
End If
'####### 이벤트 번호 #######
%>