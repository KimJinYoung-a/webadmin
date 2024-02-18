<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150" align="center">택배사</td>
		<td width="150" align="center">택배사 코드</td>
	</tr>
<%
dim sql
sql = " SELECT divcd,divname,findurl,isUsing, isTenUsing,tel " &_
			" FROM db_order.[dbo].tbl_songjang_div " &_
			" Where isusing='Y'" &_
			" ORDER BY isTenUsing desc ,divcd "

rsget.open sql,dbget,1

if not (rsget.eof or rsget.bof) then
	do until rsget.eof
%>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><%= db2html(rsget("divname")) %></td>
		<td><%= rsget("divcd") %></td>
	</tr>

<%
	rsget.movenext
	loop
end if

rsget.close
%>
    <tr bgcolor="#FFFFFF">
        <td colspan="2" align="center"><input type="button" value="닫기" onClick="window.close();"></td>
    </tr>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
