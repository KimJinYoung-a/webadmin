<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/admin/etc/my11st/my11stcls.asp"-->
<%
Dim oitem, vitemid, i, oitemoption
vItemID = Request("itemid")

Dim vOriginListImage, vOriginItemName, vOriginMakerID, vOriginSellCash, vOriginOrgPrice, vTransItemname
set oitem = new CItemInfo
	oitem.FRectItemID = vItemID
	oitem.GetOneItemInfo

	vOriginListImage = oitem.FOneItem.FListImage
	vOriginItemName = oitem.FOneItem.FItemName
	vOriginMakerID = oitem.FOneItem.FMakerid
	vOriginSellCash = oitem.FOneItem.FSellcash
	vOriginOrgPrice = oitem.FOneItem.FOrgPrice
set oitem = Nothing

set oitemoption = new CMy11st
	oitemoption.FRectItemID = vItemID
	If vItemID <> "" Then
		oitemoption.getItemOptionInfo
		vTransItemname = oitemoption.getTransItemname
	End If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function autocapypaste(){
	var OriginItemName='<%=vOriginItemName%>';

	document.frmreg.itemname.value = OriginItemName;
}
$(function(){
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});
//����
function goSubmit(){
	if(document.frmreg.itemname.value == ""){
		alert("��ǰ���� �Է��ϼ���.");
		document.frmreg.itemname.focus();
		return;
	}
	
	if(confirm("���� �Ŀ��� ��Ͽ��ΰ� 11st ��Ͽ������� �ٲ�ϴ�.\n\n���� �Ͻðڽ��ϱ�?")){
		document.frmreg.submit();
	}
}
</script>
<form name="frmreg" method="post" action="/admin/etc/my11st/my11stManagerProc.asp" style="margin:0px;">
<input type="hidden" name="itemid" value="<%=vItemID%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" border="0" class="a">
		<tr>
			<td width="100"><img src="<%=vOriginListImage%>" width="100" height="100"></td>
			<td valign="top">
				<table width="100%" border="0" class="a">
				<tr>
					<td height="23">��ǰ�� : <%=vOriginItemName%>&nbsp;&nbsp;&nbsp;<input type="button" value="��ǰ�� �Է¶��� �ֱ�" class="button" style="width:130px;" onClick="autocapypaste();"></td>
				</tr>
				<tr>
					<td height="23">��ǰ�ڵ� : <%=vItemID%> - [<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=vItemID%>" target="_blank">��ǰ�󼼺���������</a>]</td>
				</tr>
				<tr>
					<td height="23">�귣��ID : <%=vOriginMakerID%></td>
				</tr>
				<tr>
					<td height="23">�Һ��ڰ� : <%=FormatNumber(vOriginOrgPrice,0)%> / �ǸŰ� : <%=FormatNumber(vOriginSellCash,0)%></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100> ��ǰ��</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="itemname" value="<%=vTransItemname%>" size="95" maxlangth="60"></td>
</tr>
<% i = 0 %>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�ɼ�</td>
	<td valign="top" bgcolor="#FFFFFF">
		�� ó����Ͻ� �ɼ��� �ִ� ��ǰ�� ���, ����Ʈ�� �ɼ��� �������� �ɼǵ�� ���ΰ� �̵�ϻ��·� ��Ÿ���ϴ�. ������ �ʿ��� ��쿡�� ��ġ�ø� �˴ϴ�.<br><br>
		<table cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tabletop") %>">
	<%
		If oitemoption.FResultCount > 0 Then
	%>
		<% For i=0 To oitemoption.FResultCount - 1 %>
			<tr>
				<td bgcolor="#FFFFFF" align="center">
					<input type="hidden" name="itemoption<%=i%>" value="<%= oitemoption.FITemList(i).FItemOption %>" /><%= oitemoption.FITemList(i).FItemOption %>
					<% if oitemoption.FItemList(i).Fitemoption="0000" then %>
						* �ɼǾ���
						<input type="hidden" name="optiontypename<%=i%>" value="<%= oitemoption.FITemList(i).FOptionTypeName %>" />
						<input type="hidden" name="optionname<%=i%>" value="<%= oitemoption.FITemList(i).FOptionName %>" />
						<input type="hidden" name="optisusing<%=i%>" value="<%= oitemoption.FITemList(i).FOptIsUsing %>" />
					<% else %>
						<input type="text" name="optiontypename<%= i %>" value="<%= oitemoption.FITemList(i).FOptionTypeName %>" size="10" />
						<input type="text" name="optionname<%= i %>" value="<%= oitemoption.FITemList(i).FOptionName %>" size="30" />
						<span class="rdoUsing">
							<input type="radio" name="optisusing<%= i %>" id="rdoUsing<%= i %>_1" value="Y" <%= CHKIIF(oitemoption.FITemList(i).FOptIsUsing="Y","checked","") %> /><label for="rdoUsing<%= i %>_1">���</label>
							<input type="radio" name="optisusing<%= i %>" id="rdoUsing<%= i %>_2" value="N" <%= CHKIIF(oitemoption.FITemList(i).FOptIsUsing="N","checked","") %> /><label for="rdoUsing<%= i %>_2">������</label>
						</span>
						* �ɼǵ�� : <%= CHKIIF(oitemoption.FItemList(i).FNotReg="o" ,"<font color=red><b>�̵��</b></font>","<font color=blue><b>��ϿϷ�</b></font>") %>
					<% end if %>
				</td>
			</tr>
		<% Next %>
	<% End IF %>
		</table>
	</td>
</tr>
<input type="hidden" name="optioncount" value="<%=i%>" />
</table>
</form>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="window.close();" style="cursor:pointer">
    </td>
    <td align="right">
    	<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="goSubmit();" style="cursor:pointer">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<% set oitemoption = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->