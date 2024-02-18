<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/popManageColorCode.asp
' Description :  ��ǰ �÷� �ڵ���
' History : 2009.03.24 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim oitem, sMode, iColorCD, lp
Dim sColorName, sColorIcon, iSortNo, sIsUsing
iColorCD = Request.Querystring("iCD")

'// �⺻��
sMode = "I"	'���

'// �����ڵ尡 ������ �������
if iColorCD<>"" then
	sMode = "U"	'����
	set oitem = new CItemColor
	oitem.FRectColorCD = iColorCD
	oitem.GetColorList

	if oitem.FResultCount>0 then
		sColorName	= oitem.FItemList(0).FcolorName
		sColorIcon	= oitem.FItemList(0).FcolorIcon
		iSortNo		= oitem.FItemList(0).FsortNo
		sIsUsing	= oitem.FItemList(0).FisUsing
	else
		Alert_return("�߸��� ��ȣ�Դϴ�.")
		dbget.close()	:	response.End
	end if

	set oitem = Nothing
end if
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.scName.value){
			alert("�÷����� �Է����ּ���.");			
			return false;
		}

		if((!document.frmImg.scIcon.value)&&document.frmImg.mode.value=="I"){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �÷�Ĩ �̹����� ������ �ּ���.");			
			return false;
		}

		if(!document.frmImg.icSort.value){
			alert("���Ĺ�ȣ�� ���ڷ� �Է����ּ���.");			
			return false;
		}
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��ǰ �����ڵ� ����</div>
<table width="350" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/items/itemColorCodeProcess.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=sMode%>">
<% if sMode="U" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�÷��ڵ�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="icCode" size="4" readonly value="<%=iColorCD%>"></td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�÷���</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scName" size="10" maxlength="12" value="<%=sColorName%>"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�÷�Ĩ������</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="scIcon">
		<% IF sColorIcon <> "" THEN %>
			<br>���� ���ϸ� : <%=right(sColorIcon,len(sColorIcon)-instrRev(sColorIcon,"/"))%>
		<% END IF %>
	</td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="icSort" size="4" maxlength="4" style="text-align:right" value="<%=iSortNo%>"></td>
</tr>
<% if sMode="U" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="scUse" value="Y" <% if sIsUsing="Y" then Response.Write "checked" %>>���
		<input type="radio" name="scUse" value="N" <% if sIsUsing="N" then Response.Write "checked" %>>����
	</td>
</tr>
<% end if %>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<% if sMode="I" then %>
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		<% Else %>
		<a href="popManageColorCode.asp"><img src="/images/icon_cancel.gif" border="0"></a>
		<% end if %>
	</td>
</tr>
</form>
</table>
<br>
<%
	'####### �÷�Ĩ ��� #######
	set oitem = new CItemColor
	oitem.FPageSize = 50
	oitem.FRectUsing = "Y"
	oitem.GetColorList
%>
<table width="350" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="#DDDDFF">�ڵ�</td>
	<td bgcolor="#DDDDFF">Icon</td>
	<td bgcolor="#DDDDFF">�ڵ��</td>
	<td bgcolor="#DDDDFF">���Ĺ�ȣ</td>
	<td bgcolor="#DDDDFF">���</td>
</tr>
<%
	if oitem.FResultCount>0 then
		for lp=0 to oitem.FResultCount-1
%>
<tr align="center">
	<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FcolorCode%></td>
	<td bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="1" bgcolor="#dddddd"><tr><td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="12" height="12" hspace="2" vspace="2"></td></tr></table></td>
	<td bgcolor="#FFFFFF"><a href="popManageColorCode.asp?iCD=<%=oitem.FItemList(lp).FcolorCode%>"><%=oitem.FItemList(lp).FcolorName%></a></td>
	<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FsortNo%></td>
	<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FisUsing%></td>
</tr>
<%
		next
	else
		Response.Write "<tr><td colspan=5 height=50 align=center bgcolor=#F8F8F8>��ϵ� ������ �����ϴ�.</td></tr>"
	end if
%>
</table>
<%
	set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->