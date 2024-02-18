<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  담당업무 수정
' History : 2011.04.25 김진영 생성
'			2017.04.10 한용민 수정
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim oMember
Set oMember = new CTenByTenMember
	oMember.Fempno = requestCheckVar(request("empno"),32)
	oMember.fnGetMemberData
%>
<script type='text/javascript'>

function modify(emp){
	var frm = document.frm;
	frm.action = "/admin/member/tenbyten/domodifymemberinfo.asp?mode=mywork&empno="+ emp;
	frm.submit();
}

</script>
<b><center>담당업무 수정</center></b><p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="frm" method="post">
<tr bgcolor="#DDDDFF">
	<td width="130" height="30">담당업무</td>
	<td bgcolor="#FFFFFF" height="30"><input type="text" class="text" name="mywork" value="<%=oMember.Fmywork%>" maxlength="80" size="60"></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="center" colspan="3" height="30">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_modify.gif" onclick="modify('<%=oMember.Fempno%>');" style="cursor:hand">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->