<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

<%
	Response.CharSet = "euc-kr"
	
	Dim cDisp, vDepth, vItemID, vItemName, vCateCode, vCateName, vSortNo, vIsDefault, vResultCount
	vItemID		= Request("itemid")
	vDepth		= Request("depth")
	vCateCode 	= Request("catecode")
	
	SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.FRectItemID = vItemID
	cDisp.GetDispCateItemDetail()
	
	vCateName 	= cDisp.FCateFullName
	vItemName	= cDisp.FItemName
	vSortNo		= cDisp.FSortNo
	vIsDefault	= cDisp.FIsDefault
	vResultCount = cDisp.FResultCount
	SET cDisp = Nothing
%>
<% If vResultCount > 0 Then %>
<script>
$(function() {
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

function updateCate(){
	$('input[name="action"]').val('update');
	frmCateItem.submit();
	location.reload();
}

function deleteCate(){
	if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?") == true) {
		$('input[name="action"]').val('delete');
		frmCateItem.submit();
		location.reload();
	}
}
</script>

<form name="frmCateItem" action="display_cate_item_proc.asp" method="post" style="margin:0px;" target="cateitemproc">
<input type="hidden" name="action" value="update">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="itemid" value="<%=vItemID%>">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr>
	<td bgcolor="#F3F3FF" width="25%">��ǰ</td>
	<td bgcolor="#FFFFFF">[<%=vItemID%>]<%=vItemName%></td>
</tr>
<tr>
	<td bgcolor="#F3F3FF">ī�װ�</td>
	<td bgcolor="#FFFFFF">[<%=vCateCode%>]<%=vCateName%></td>
</tr>
<tr>
	<td bgcolor="#F3F3FF">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortno" style="width:70px;" value="<%=vSortNo%>"><br />(�� ���ڰ� �������� ���)</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF">�⺻����</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isdefault" id="useyn_1" value="y" <%=CHKIIF(vIsDefault="y","checked","")%> /><label for="useyn_1">�⺻ī�װ�</label>
		<input type="radio" name="isdefault" id="useyn_2" value="n" <%=CHKIIF(vIsDefault="n","checked","")%> /><label for="useyn_2">�⺻�ƴ�</label>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" border="0" cellpadding="2" cellspacing="2" class="a">
		<tr>
			<td id="lyrSbmBtn"><input type="button" value="��    ��" onClick="updateCate();"></td>
			<td id="lyrSbmBtn2" align="right"><input type="button" value="��    ��" onClick="deleteCate();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<script>
	$("#lyrSbmBtn input").button();
	$("#lyrSbmBtn2 input").button();
</script>
<% Else
	SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateDetail()
	
	vCateName 	= cDisp.FCateFullName
	SET cDisp = Nothing
	Response.Write "<b>" & vCateName & "</b><br>ī�װ��� ������ �� �ֽ��ϴ�."
%>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->