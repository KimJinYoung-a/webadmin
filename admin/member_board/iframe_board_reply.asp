<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� �亯
' History : 2011.02.28 ������ ����
' 			2018.11.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<%
	Dim brd_sn, rBoard, page, cmt_content
	Dim rcmt_sn, rid, cidx, i
	brd_sn = request("brd_sn")
	rid	= NullFillWith(requestCheckVar(Request("rid"),50),"")
	cidx = NullFillWith(requestCheckVar(Request("cidx"),10),"")

	page = request("page")
	If page = "" Then page = 1

	Set rBoard = new Board
		rBoard.Fbrd_sn = brd_sn
		rBoard.fnBoardreplylist

	If cidx <> "" Then
		rBoard.Fcmt_sn = cidx
		rBoard.fnBoardreplymodify
		cmt_content = rBoard.FOnecmt.Fcmt_content
	End If

%>
<script language="javascript">
function form_check(){
	var frm = document.frm
	if(document.frm.cmt_content.value == ""){
		alert("����� �Է��ϼ���");
		frm.focus();
		return false;
	}
	frm.action = "iframe_board_reply_proc.asp";
	frm.submit();
}
function gosubmit(page){
    frm.page.value=page;
    frm.submit();
}
function reply_edit(cidx){
	location.href = "iframe_board_reply.asp?brd_sn=<%=brd_sn%>&page=<%=page%>&cidx="+cidx;
}
function reply_del(cidx){
	if(confirm("�����Ͻ� ���� �����Ͻðڽ��ϱ�?") == true) {
		location.href = "iframe_board_reply_proc.asp?mode=del&brd_sn=<%=brd_sn%>&page=<%=page%>&cidx="+cidx;
	} else {
		return false;
	}
}
</script>
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<form name="frm" method="post" style="margin:0px;">
<table width="810" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type = "hidden" name="brd_sn" value=<%=brd_sn%>>
<input type = "hidden" name="mode" value="add">
<input type = "hidden" name="page" value="">
<input type = "hidden" name="cidx" value="<%=cidx%>">
<tr bgcolor="#FFFFFF">
	<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��۳���</td>
	<td align="left"><textarea class="textarea" name="cmt_content" cols="100" rows="5"><%= cmt_content %></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right">
		<img src= "http://webadmin.10x10.co.kr/images/icon_reply.gif" onclick="form_check()" style="cursor:hand">
	</td>
</tr>
</table>
</form>
<br>
<table width="810" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" width="140">�ۼ���</td>
	<td align="center">��&nbsp;&nbsp;&nbsp;��</td>
</tr>
<%
If rBoard.FTotalCount = "0" Then
%>
<tr bgcolor="#FFFFFF" height="30">
	<td colspan="2" align="center" class="page_link">[����� �����ϴ�.]</td>
</tr>
<%
Else
%>
<%
For i = 0 To rBoard.fresultcount -1
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="center" valign="top" style="padding:3 0 0 3" width="170">
		<%
		Response.Write rBoard.FcmtList(i).Fusername & "(" & rBoard.FcmtList(i).fid & ")"
		Response.Write "<br>" & rBoard.FcmtList(i).Fcmt_regdate
		If rBoard.FcmtList(i).Fid = session("ssBctId") Then
			Response.Write "<br><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_modify.gif' style='cursor:pointer' onClick='reply_edit(" & rBoard.FcmtList(i).Fcmt_sn & ")'>"
			Response.Write "&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='reply_del(" & rBoard.FcmtList(i).Fcmt_sn & ")'>"
		ElseIf C_ADMIN_AUTH Then
			Response.Write "<br><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='reply_del(" & rBoard.FcmtList(i).Fcmt_sn & ")'>"
		
		' cs������ �ϰ��
		elseif C_CSPowerUser Then
			' cs���� �� ���� �ϰ�� ���������ϰ� ó��.
			if rBoard.FcmtList(i).fpart_sn="10" Then
				Response.Write "<br><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='reply_del(" & rBoard.FcmtList(i).Fcmt_sn & ")'>"
			end if
		End If
		%>
	</td>
	<td align="left" style="padding:3 3 3 3"><%=ReplaceBracket(replace(rBoard.FcmtList(i).Fcmt_content,vbCrLf,"<br>"))%></td>
</tr>
<% Next %>
<%
End If
%>
</table>
<%
set rBoard = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
