<%
	Dim g_TeamJang, g_PartJang, g_MyTeam, g_MyPart, g_MenuPos
	IF application("Svr_Info")="Dev" THEN
		g_MenuPos   = "1404"		'### �޴���ȣ ����.
	Else
		g_MenuPos   = "1404"		'### �޴���ȣ ����.
	End If
	g_TeamJang = "x"
	g_PartJang = "x"
	
	
	'####### ������, ������ ���� #######
	If CInt(session("ssAdminLsn")) =< CInt("2") Then
		g_TeamJang = "o"
	End If
	
	'####### ��Ʈ���ӱ��� #######
	If CInt(session("ssAdminLsn")) = CInt("3") Then
		g_PartJang = "o"
	End If
	



	g_MyPart = session("ssAdminPsn")
	
	'####### �ý�����, �濵������, �мǻ���� ������ #######
	If CInt(g_MyPart) = CInt("7") OR CInt(g_MyPart) = CInt("8") OR CInt(g_MyPart) = CInt("17") Then
		g_MyTeam = g_MyPart
	End IF

	'####### ���̶�������� #######
	If CInt(g_MyPart) = CInt("15") OR CInt(g_MyPart) = CInt("19") Then
		g_MyTeam = "15,19"
	End IF

	'####### �������������� #######
	If CInt(g_MyPart) = CInt("13") OR CInt(g_MyPart) = CInt("18") Then
		g_MyTeam = "13,18"
	End IF

	'####### ������� ������ #######
	If CInt(g_MyPart) = CInt("9") OR CInt(g_MyPart) = CInt("10") Then
		g_MyTeam = "9,10"
	End IF
	
	'####### �ٹ����ٻ����(��������,MD,WD) ������ #######
	If CInt(g_MyPart) = CInt("11") OR CInt(g_MyPart) = CInt("12") OR CInt(g_MyPart) = CInt("14") OR CInt(g_MyPart) = CInt("16") Then
		g_MyTeam = "11,12,14,16"
	End IF
	
	

	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
		'### test��
		'g_TeamJang = "o"
		'g_PartJang = "o"
	End If
%>
