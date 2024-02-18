<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 시스템관리 > VPN접속현황
' History : 서동석 생성
'			2017.05.19 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim i, j, vQuery, vIdx, vLogCont, vLogTmp, vGubun, vWhyCon
	vGubun = requestCheckVar(Request("gubun"),10)
	vIdx = requestCheckVar(Request("idx"),10)
	vLogCont = Request("logcont")
	vWhyCon = Request("whycon")
	
	If vGubun = "whycon" Then
		vQuery = "UPDATE [db_board].[dbo].[tbl_vpn_connect_log] SET "
		vQuery = vQuery & "whycon = '" & vWhyCon & "', whyuserid = '" & session("ssBctId") & "', whyregdate = getdate() "
		vQuery = vQuery & "WHERE idx = '" & vIdx & "'"
		dbget.execute vQuery
	ElseIf vGubun = "sign" Then
		vQuery = "UPDATE [db_board].[dbo].[tbl_vpn_connect_log] SET "
		vQuery = vQuery & "sign = '" & session("ssBctId") & "', signdate = getdate() "
		vQuery = vQuery & "WHERE idx = '" & vIdx & "'"
		dbget.execute vQuery
	ElseIf vGubun = "onedel" Then
		vQuery = "delete from [db_board].[dbo].[tbl_vpn_connect_log] where "
		vQuery = vQuery & " idx = '" & vIdx & "'"
		dbget.execute vQuery
	Else
		For i = LBound(Split(vLogCont,vbCrLf)) To UBound(Split(vLogCont,vbCrLf))
			vLogTmp = Split(vLogCont,vbCrLf)(i)
			
			If vLogTmp <> "" Then
				vQuery = "IF NOT EXISTS(select idx from [db_board].[dbo].[tbl_vpn_connect_log] where stime = '" & Split(vLogTmp,"	")(0) & "' and realip='"& Split(vLogTmp,"	")(2) &"') "
				vQuery = vQuery & " BEGIN "
				'vQuery = vQuery & "INSERT INTO [db_board].[dbo].[tbl_vpn_connect_log](stime, etime, equip, userid, realip, assignip, loginstate, constate, reguserid) "
				'vQuery = vQuery & "VALUES('" & Split(vLogTmp,"	")(0) & "','" & Split(vLogTmp,"	")(1) & "','" & Split(vLogTmp,"	")(2) & "','" & Split(vLogTmp,"	")(3) & "','" & Split(vLogTmp,"	")(4) & "',"
				'vQuery = vQuery & "'" & Split(vLogTmp,"	")(5) & "','" & Split(vLogTmp,"	")(6) & "','" & Split(vLogTmp,"	")(7) & "','" & session("ssBctId") & "') "
				vQuery = vQuery & " 	INSERT INTO [db_board].[dbo].[tbl_vpn_connect_log](stime, userid, realip, reguserid) "
				vQuery = vQuery & " 	VALUES('" & Split(vLogTmp,"	")(0) & "','" & Split(vLogTmp,"	")(1) & "','" & Split(vLogTmp,"	")(2) & "'"
				vQuery = vQuery & "		,'" & session("ssBctId") & "') "
				vQuery = vQuery & " END "

				'response.write vQuery & "<br>"
				'response.end
				dbget.execute vQuery
			End If
		Next
	End If
%>

<script>
	<% If vGubun = "sign" or vGubun = "onedel" Then %>
		parent.location.reload();
	<% Else %>
		opener.location.reload();
		window.close();
	<% End If %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->