<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : Category_left_bestBrand.asp
' Discription : ī�װ� ���� ����Ʈ �귣��
' History : 2008.04.02 �ѿ�� �ٹ����پ��� ����/����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/diycategorycls.asp" -->
<%
dim cdl,idx,mode,i , fmainitem
	cdl=request("cdl")
	mode=request("mode")
	idx=request("idx")
%>

<script language="javascript">

function subcheck(){
	var frm=document.inputfrm;

	if (frm.cdl.value.length<1) {
		alert('ī�װ��� ������ �ּ���..');
		frm.cdl.focus();
		return;
	}

	if (frm.makerid.value.length< 1 ){
		 alert('��ü�� ���� ���ּ���');
	frm.makerid.focus();
	return;
	}

	if(!frm.sortNo.value) {
		alert("ǥ�� ������ �Է����ּ���.\n�� ������ �����̸� �������� ������ �����ϴ�.");
		frm.sortNo.focus();
		return;
	}

	frm.submit();
}

function popBrandSearch(fm,tg){
	var popup_item = window.open("/admin/member/popBrandSearch.asp?frmName=" + fm + "&compName=" + tg, "popup_brand", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

function changecontent()
{}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="<%=imgFingers%>/linkweb/items/bestbrand/doLeftBestBrand.asp" enctype="multipart/form-data">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>����Ʈ �귣�� ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">ī�װ�����</td>
	<td bgcolor="#FFFFFF"><% DrawSelectBoxacademyCategoryLarge "cdl", cdl %></td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�� ü</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="makerid" value="">
		<input type="button" class="button" value="ã��" onClick="popBrandSearch('inputfrm','makerid')">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="img1" value="" size="55" class="text"><br>
		(�̹��� Size�� 180x75 �Դϴ�..)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="sortNo" value="0" size="3">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" checked>Y
		<input type="radio" name="isusing" value="N">N
	</td>
</tr>
<% elseif mode="edit" then %>
<%
set fmainitem = New CCatemanager
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectIdx=idx
	fmainitem.GetBestBrandList()
%>
<tr>
	<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">ī�װ� ����</td>
	<td bgcolor="#FFFFFF"><% DrawSelectBoxacademyCategoryLarge "cdl", fmainitem.FItemList(0).Fcdl %></td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ü</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="makerid" value="<%=fmainitem.FItemList(0).Fmakerid%>">
		<input type="button" class="button" value="ã��" onClick="popBrandSearch('inputfrm','makerid')">
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
	<td bgcolor="#FFFFFF"><input type="file" name="img1" size="55"><br>
		(�̹��� size�� 180x240 �Դϴ�..)<br>
		<table border="1" cellpadding="0" cellspacing="0" width="180" height="212" class="a">
		<tr><td><img src="<%= imgFingers & "/left/bestbrand/" & fmainitem.FItemList(0).FImage %>" border="0" name="imgv1"></td></tr>
		<tr><td bgcolor="#303030" align="center"><font color="white"><%= fmainitem.FItemList(0).FImage %></font></td></tr>
		</table>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="sortNo" value="<%= fmainitem.FItemList(0).FsortNo %>" size="3">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%if fmainitem.FItemList(0).FIsusing="Y" then response.write "checked" %> checked>Y
		<input type="radio" name="isusing" value="N" <%if fmainitem.FItemList(0).FIsusing="N" then response.write "checked" %>>N
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>

<%
	set fmainitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->