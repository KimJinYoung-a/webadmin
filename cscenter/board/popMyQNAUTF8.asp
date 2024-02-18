<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%

'==============================================================================
'나의 1:1질문답변
dim boardqna
set boardqna = New CMyQNA

boardqna.read(request("idx"))

%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
    	<td height="25" width="90" align="center" bgcolor="#FFFFFF"><b>제목</b></td>
    	<td width="570" bgcolor="#FFFFFF">
			<%= nl2br(db2html(boardqna.results(0).title)) %>
    	</td>
    </tr>
	<tr>
    	<td height="25" width="90" align="center" bgcolor="#FFFFFF"><b>내용</b></td>
    	<td width="570" bgcolor="#FFFFFF">
			<%= nl2br(db2html(boardqna.results(0).contents)) %>
    	</td>
    </tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
