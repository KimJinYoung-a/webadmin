<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'           2014.03.03 허진원 ; 개인정보 데이터 제거
'			2014.03.10 한용민 수정
'			2015.06.26 유태욱(초능력자들 이벤트용으로 임시 생성-이벤트 종료후 삭제예정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim eCode, Sdate, Edate, limitLevel, strSql
	eCode = Request("eC")	'이벤트코드
	Sdate = Request("Sdate")	'적용시작일
	Edate = Request("Edate")	'적용종료일
	limitLevel = Request("limitLevel")	'회원등급제한

'다운로드 기록 로그 저장
strSql = "insert into db_log.[dbo].[tbl_caution_event_log] (evt_code, userid, refip, device ) values " &_
		" ('"& eCode &"'" &_
		", '"& session("ssBctId") &"'" &_
		", '"& Left(request.ServerVariables("REMOTE_ADDR"),32) & "'" &_
		", 'S')"
dbget.Execute strSql

'// DB에서 목록접수
strSql = "select " &_
		"	t1.sub_idx, t1.userid , t2.username, t2.usercell, t1.regdate, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3 " &_
		"	, Case t3.userlevel  " &_
		"		When '0' Then 'Yellow'  " &_
		"		When '1' Then 'Green'  " &_
		"		When '2' Then 'Blue'  " &_
		"		When '3' Then 'VIP Siver'  " &_
		"		When '4' Then 'VIP Gold'  " &_
		"		When '5' Then 'Orange'  " &_
		"		When '6' Then 'Friends' " &_
		"		When '7' Then 'Staff' " &_
		"		When '9' Then '감성매니아'  " &_
		"	 End as userlevel  " &_
		"	,isnull((select count(*) from db_event.dbo.tbl_event_prize where t2.userid=evt_winner),0) as eventcount" &_
		" from db_event.dbo.tbl_event_subscript as t1 " &_
		"	Join db_user.[dbo].tbl_user_n as t2 " &_
		"		on t1.userid=t2.userid " &_
		"	Join db_user.[dbo].tbl_logindata as t3 " &_
		"		on t2.userid=t3.userid " &_
		" left join db_user.dbo.tbl_invalid_user iu" &_
		" 	on t1.userid=iu.invaliduserid" &_
		" 	and iu.isusing='Y'" &_
		" 	and iu.gubun='ONEVT'" &_
		" where iu.idx is null and t1.sub_opt2<>0 and t1.evt_code=" & eCode &_
		"	and t1.regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "

	Select Case limitLevel
		Case "orange"
			strSql = strSql & "	and t3.userlevel not in ('5') "
		Case "yellow"
			strSql = strSql & "	and t3.userlevel not in ('0','5') "
	end Select
'	response.write strsql
'	response.end
	rsget.Open strSql, dbget, 1
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=event" & eCode & "_" & Date() & ".xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>번호</td>
	<td>회원ID</td>
	<td>이름</td>
	<td>응모일시</td>
	<td>연락처</td>
	<td>응모여부(횟수)</td>
	<td>당첨상품</td>
	<td>카카오초대여부</td>
	<td>회원등급</td>
	<td>이벤트당첨횟수</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td><%=rsget("sub_idx")%></td>
	<td style='mso-number-format:"\@";'><%=rsget("userid")%></td>
	<td><%=rsget("username")%></td>
	<td><%=rsget("regdate")%></td>
	<td><%=rsget("usercell")%></td>
	<td><%=rsget("sub_opt1")%></td>
	<td><%=rsget("sub_opt2")%></td>
	<td><%=rsget("sub_opt3")%></td>
	<td><%=rsget("userlevel")%></td>
	<td><%=rsget("eventcount")%></td>
</tr>
<%
		rsget.MoveNext
		loop
	else
%>
<tr><td colspan="13" align="center">조건에 맞는 참여자가 없습니다</td></tr>
<%	end if %>
</table>
</body>
</html>
<% rsget.close %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
