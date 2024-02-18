<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  이벤트 응모자 리스트
' History : 2007.09.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventuserclass.asp"-->

<% 
dim seachbox,eventbox ,page , i
	seachbox = request("seachbox")
	eventbox = request("eventbox")
	page = Request("page")

If page="" Then page = 1

dim oeventuserlist
set oeventuserlist = new Ceventuserlist
	if eventbox = "3" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist3()
	end if 
	if eventbox = "5" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist5()
	end if 
	if eventbox = "7" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist7()
	end if 
	if eventbox = "9" then
		set oeventuserlist = new Ceventuserlist
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectinvaliduseryn="N"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist9()
	end if 

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=event_userlist" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>
<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25" valign="top">
	<td>
		<font color="red"><strong><%= seachbox %> 이벤트 응모자 리스트</strong></font>
	</td>
</tr>
</table>
<!--표 헤드끝-->
	
<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
<tr bgcolor=#DDDDFF>
	<td align="center">id</td>
	<!--<td align="center">주민번호앞자리</td>-->
	<td align="center">성별</td>
	<td align="center">이름</td>
	<td align="center">이메일</td>
	<td align="center">전화번호</td>
	<td align="center">핸드폰번호</td>
	<td align="center">주소</td>
	<td align="center">레벨</td>
	<td align="center">코멘트</td>

	<% If eventbox = "9" Then %>
		<td align="center">당첨횟수</td>
		<td align="center">최근당첨일</td>
	<% End If %>
</tr>
<% if oeventuserlist.ftotalcount >0 then %>
<% for i = 0 to oeventuserlist.ftotalcount - 1 %>
<tr bgcolor=#FFFFFF>
	<td><%= oeventuserlist.flist(i).fuserid %></td>
	<!--<td><%'= left(oeventuserlist.flist(i).fjuminno,6) %>-->
	</td>
	<td>
		<%
		if mid(oeventuserlist.flist(i).fjuminno,8,1) = "1" then
		response.write "남성"
		else
		response.write "여성"
		end if
		%>
	</td>
	<td><%= oeventuserlist.flist(i).fusername %></td>
	<td><%= oeventuserlist.flist(i).fusermail %></td>
	<td><%= oeventuserlist.flist(i).fuserphone %></td>
	<td><%= oeventuserlist.flist(i).fusercell %></td>
	<td align="left">
		[<%= oeventuserlist.flist(i).fzipcode %>]
		&nbsp;
		<%= oeventuserlist.flist(i).faddress1 %>
		&nbsp;
		<%= oeventuserlist.flist(i).fuseraddr %>
	</td>
	<td><%= getUserLevelStr(oeventuserlist.flist(i).fLevel) %></td>
	<td><%= replace(oeventuserlist.flist(i).fevtcom_txt,"<","&lt;") %></td>

	<% If eventbox = "9" Then %>
		<td><%= oeventuserlist.flist(i).fWcnt %></td>
		<td><%= oeventuserlist.flist(i).fWdate %></td>
	<% End If %>
</tr>
<% next %>

<% else %>
<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF" colspan=15>검색 결과가 없습니다.</td>
</tr>
<% end if %>	
</table>

</body>
</html>

<%
set oeventuserlist=nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

