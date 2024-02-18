<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/10x10_boardcls.asp"-->

<%
dim oboard,idx
idx = request("idx")

set oboard = new CHopeBoardDetail
oboard.read idx


%>
<script language="JavaScript">
<!--
	function GotoDel(){
		if (confirm("삭제하시겠습니까?")){
			location.href = "10x10_board_act.asp?idx=<% = idx %>&mode=delete";
		}
	}
//-->
</script>
<table border="0" cellpadding="5" cellspacing="0" width="650" class="a">

	<tr>
		<td bgcolor="white" style="padding:2" align="right" valign="bottom" colspan="2">
			<% = FormatDateTime(oboard.Fregdate,1) %>
		</td>
	</tr>
	<tr>
		<td bgcolor="#46699c" style="font-weight:bold;color:white">
			&nbsp; 제목 :  <%=oboard.FTitle %> </td>
		<td bgcolor="#e0e0e0" align="right" width="150">
		 글쓴이 : <%=oboard.Fusername %>(<%=oboard.Fuserid %>)&nbsp;
		 </td>
	</tr>
	<tr>
		<td bgcolor="#EFEFEF" style="padding: 20 20 20 20;border-bottom:1 solid #99a9bc" colspan="2" class=a>
		<%=oboard.FContents %>
		<% if oboard.Fuserid = session("ssBctId") then %>
		<br><br>
		<div align="right"><a href="10x10_board_modify.asp?idx=<% = idx %>">수정</a> | <a href="javascript:GotoDel();">삭제</a></div>
		<% end if %>
		</td>
	</tr>
</table><br>
<!-- #include virtual ="/admin/board/10x10_board_comment.asp" -->
<%
set oboard = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->