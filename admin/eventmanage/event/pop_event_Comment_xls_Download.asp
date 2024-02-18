<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'           2014.03.10 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
response.write "사용중지 매뉴입니다. 신매뉴로 사용 부탁드립니다."
response.end

dim eCode, Sdate, Edate, limitLevel, strSql
	eCode = Request("eC")	'이벤트코드
	Sdate = Request("Sdate")	'적용시작일
	Edate = Request("Edate")	'적용종료일
	limitLevel = Request("limitLevel")	'회원등급제한

'// DB에서 목록접수
strSql = "select " &_
		"	t1.evtcom_idx, t1.userid " &_
		"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
		"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
		"		end as userAge " &_
		"	, t3.userlevel " &_
		"	, t1.evtcom_regdate, t2.regdate as joindate " &_
		"	, t1.evtcom_txt, t1.evtcom_point, t1.blogurl " &_
		"	,(select count(*) FROM [db_event].[dbo].[tbl_event_prize] as p WHERE p.evt_winner = t2.userid) as wincnt  " &_
		"	,(select top 1 evt_regdate FROM [db_event].[dbo].[tbl_event_prize] as p WHERE p.evt_winner = t2.userid order by evt_regdate desc) as windate " &_
		"	,t2.username " &_
		" from db_event.dbo.tbl_event_comment as t1 " &_
		"	Join db_user.[dbo].tbl_user_n as t2 " &_
		"		on t1.userid=t2.userid " &_
		"	Join db_user.[dbo].tbl_logindata as t3 " &_
		"		on t2.userid=t3.userid " &_
		" left join db_user.dbo.tbl_invalid_user iu" &_
		" 	on t1.userid=iu.invaliduserid" &_
		" 	and iu.isusing='Y'" &_
		" 	and iu.gubun='ONEVT'" &_
		" where iu.idx is null and t1.evt_code=" & eCode &_
		"	and t1.evtcom_using='Y' " &_
		"	and t1.evtcom_regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "

	Select Case limitLevel
		Case "orange"
			strSql = strSql & "	and t3.userlevel not in ('5') "
		Case "yellow"
			strSql = strSql & "	and t3.userlevel not in ('0','5') "
		Case "white"
			strSql = strSql & "	and t3.userlevel not in ('0') "
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
	<td>나이</td>
	<td>회원등급</td>
	<td>작성일</td>
	<td>회원가입일</td>
	<td>코멘트 내용</td>
	<td>선택번호</td>
	<td>블로그주소</td>
	<td>최근당첨일</td>
	<td>이벤트당첨횟수</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td><%=rsget("evtcom_idx")%></td>
	<td><%=rsget("userid")%></td>
	<td><%=rsget("username")%></td>
	<td><%=rsget("userAge")%></td>
	<td><%= getUserLevelStrByDate(rsget("userlevel"), left(rsget("evtcom_regdate"),10)) %></td>
	<td><%=rsget("evtcom_regdate")%></td>
	<td><%=rsget("joindate")%></td>
	<td><%=rsget("evtcom_txt")%></td>
	<td><%=rsget("evtcom_point")%></td>
	<td><%=rsget("blogurl")%></td>
	<td><%=rsget("windate")%></td>
	<td><%=rsget("wincnt")%></td>
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
