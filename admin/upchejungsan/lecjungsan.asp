<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드ID</td>
		<td>스트리트명(한글)</td>
		<td>스트리트명(영문)</td>
		<td>사용여부</td>
		<td>비고</td>
	</tr>
<%
dim sqlStr
sqlStr = "select userid, socname_kor, socname, isusing from [db_user].[dbo].tbl_user_c"
sqlStr = sqlStr + " where userdiv='14'"
sqlStr = sqlStr + " order by isusing desc"

rsget.Open sqlStr,dbget,1
do until rsget.eof
%>

	<% if rsget("isusing")="N" then %>
	<tr align="center" bgcolor="<%= adminColor("gray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="/admin/upchejungsan/mijungsanlist.asp?designer=<%= rsget("userid") %>&gubun=lecture"><%= rsget("userid") %></a></td>
		<td><%= db2html(rsget("socname_kor")) %></td>
		<td><%= db2html(rsget("socname")) %></td>
		<td><%= db2html(rsget("isusing")) %></td>
		<td></td>
	</tr>
<%
	rsget.moveNext
loop
rsget.Close
%>


</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
