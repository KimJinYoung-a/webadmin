<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMenuCls.asp"-->
<%
	Dim cDisp, i, vCateCode
	vCateCode = Request("catecode")
	
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function goCateCode(c){
	location.href = "<%=CurrURL()%>?menupos=<%=Request("menupos")%>&catecode="+c+"";
}
</script>
<style type="text/css">
.box1 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FFF8F8; padding:7px 10px;}
</style>
<div class="box1">
############################################ <b>�� ��</b> ############################################<br>
* <b>�Է� �� ���� ��</b> �ݵ�� <b>�����ϱ⸦ ����</b>�ּ���. �׷��� ������ ������ ���� �ʽ��ϴ�.<br>
* <b>�����ϱ� ��</b> ���� �����Ϸ��� �ݵ�� <b>�̸����⸦ ���� <font color="red">�Ǽ��� �����ϱ⸦ ����</font></b>�ּ���. �׷��� ������ ���� ������ ���� �ʽ��ϴ�.<br>
* �� �������� �������� �귣��� ���������� ���̵� �������� ���� ����Ǵ� ���������� ��Ʈ��Ʈ��(����)���� ��ȯ�Ǿ� �����˴ϴ�.<br>
###############################################################################################<br>
</div>
<br>
ī�װ����� : 
<%
If cDisp.FResultCount > 0 Then
	Response.Write "<select name=""catecode"" class=""select"" onChange=""goCateCode(this.value);"">" & vbCrLf
	Response.Write "<option value="""">����</option>" & vbCrLf
	For i=0 To cDisp.FResultCount-1
		Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
	Next
	Response.Write "</select>"
End If
Set cDisp = Nothing

If vCateCode <> "" Then
%>
	<iframe name="menuiframe" id="menuiframe" src="/admin/CategoryMaster/displaycate/menu/templete_root.asp?catecode=<%=vCateCode%>" width="100%" height="410px" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<%
Else
	'Response.Write "<br><br>"
End If
%>

<!--
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr align="center" bgcolor="#F3F3FF" height="30">
	<td width="4%"></td>
	<td width="6%"></td>
	<td width="10%">Maker ID</td>
	<td>��ǰ��</td>
	<td width="35%">������ī�װ�</td>
</tr>
<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td></td>
	<td></td>
</tr>
</table>
//-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->