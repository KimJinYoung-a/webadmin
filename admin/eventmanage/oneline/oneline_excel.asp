<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/event/onelineCls.asp"-->

<%
	Dim vEvtCode, vGubun, vQuery, vIdx, sqlStr, vSDate, vSubQuery
	vEvtCode 	= requestCheckVar(Request("eC"),10)
	vIdx		= requestCheckVar(Request("idx"),10)
	vSDate		= requestCheckVar(Request("esday"),10)
	vGubun		= Request("gubun")

	sqlStr = "SELECT O.idx, O.userid, O.comment, O.winYN, O.isusing, O.regdate, O.icon, U.userlevel " & _
			 "		FROM [db_contents].[dbo].[tbl_one_comment] AS O " & _
			 "	INNER JOIN [db_user].[dbo].[tbl_logindata] AS U ON O.userid = U.userid " & _
			 "	WHERE O.evt_code = '" & vEvtCode & "' " & _
			 "	" & vSubQuery & " " & _
			 "	ORDER BY O.idx DESC "
	rsget.Open sqlStr, dbget, 1
%>

<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=" & DatePart("m",vSDate) & "월" & getWeekSerial(vSDate) & "주차_한줄List.xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>번호</td>
	<td>icon</td>
	<td>아이디</td>
	<td>회원등급</td>
	<td>등록일</td>
	<td>내용</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr height="30">
	<td><%=rsget("idx")%></td>
	<td width="30" align="center" valign="middle"><img src="http://fiximage.10x10.co.kr/web2010/oneline/emoticon_0<%=rsget("icon")%>_s.gif" width="20" height="20"></td>
	<td><%=rsget("userid")%></td>
	<td><%= getUserLevelStrByDate(rsget("userlevel"), left(rsget("regdate"),10)) %></td>
	<td><%=rsget("regdate")%></td>
	<td><%=rsget("comment")%></td>
</tr>
<%
		rsget.MoveNext
		loop
	else
%>
<tr><td colspan="13" align="center">참여자가 없습니다</td></tr>
<%	end if %>
</table>
</body>
</html>
<% rsget.close %>

<!-- #include virtual="/lib/db/dbclose.asp" -->