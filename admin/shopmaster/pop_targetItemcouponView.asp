<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/base64.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/newitemcouponcls.asp"-->
<%
	'��������
	dim itemcouponidx, itemid
	dim oitemcouponmaster, ocouponitemlist

	itemcouponidx	= request("icpidx")
	itemid			= request("iid")

	'Ÿ������ Ȯ��
	set oitemcouponmaster = new CItemCouponMaster
	oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
	oitemcouponmaster.GetOneItemCouponMaster

	if oitemcouponmaster.FResultCount<1 then
		Call Alert_Close("�߸��� �����Դϴ�.")
		response.End
	end if

	'������ǰ Ȯ��
	set ocouponitemlist = new CItemCouponMaster
	ocouponitemlist.FPageSize=1
	ocouponitemlist.FCurrPage=1
	ocouponitemlist.FRectItemCouponIdx = itemcouponidx
	ocouponitemlist.FRectsRectItemidArr = itemid
	ocouponitemlist.GetItemCouponItemList

	if ocouponitemlist.FResultCount<1 then
		Call Alert_Close("���ų� �߸��� ��ǰ�Դϴ�.")
		response.End
	end if
%>
<script type="text/javascript">
// Ŭ������� ����
function fnCBCopy(iid,dvc) {
	var doc, dmn
	switch(dvc) {
		case "w":
			dmn = "http://www.10x10.co.kr/shopping/category_prd.asp";
			break;
		case "m":
			dmn = "http://m.10x10.co.kr/category/category_itemprd.asp";
			break;
		case "a":
			dmn = "http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp";
			break;
	}
	doc = dmn + "?itemid=" + iid + "&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>";
	clipboardData.setData("Text", doc);
	alert('��ũ�� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.');
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
    <td colspan="4"><b>Ÿ������ ��ũ Ȯ��/����</b></td>
</tr>
<tr bgcolor="#E8E8EE">
	<td width="80">������</td>
	<td colspan="3" bgcolor="#FFFFFF"><%= oitemcouponmaster.FOneItem.Fitemcouponname %></td>
</tr>
<tr bgcolor="#E8E8EE">
	<td>������</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetDiscountStr %>
	</td>
</tr>
<tr bgcolor="#E8E8EE">
	<td>����Ⱓ</td>
	<td colspan="3" bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.Fitemcouponstartdate %> ~ <%= oitemcouponmaster.FOneItem.Fitemcouponexpiredate %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>��ǰ��ȣ</td>
	<td colspan="3" bgcolor="#FFFFFF"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>��ǰ��</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<%= ocouponitemlist.FitemList(0).FItemName %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="80">�����ǸŰ�</td>
	<td bgcolor="#FFFFFF">
		<%= FormatNumber(ocouponitemlist.FitemList(0).GetCouponSellcash,0) %>��
		<% if ocouponitemlist.FitemList(0).Fitemcoupontype="3" then %><font color="red">(������)</font><% end if %>
	</td>
	<td width="80">�����ǸŰ�</td>
	<td bgcolor="#FFFFFF">
		<%= FormatNumber(ocouponitemlist.FitemList(0).FSellcash,0) %>��
	</td>
</tr>
</table><br>
�� �Ʒ� ��ũ�� ���ؼ� �����ϸ� ���ΰ��ݰ� �����ٿ�ε尡 ǥ�õ˴ϴ�.<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#DDDDFF" align="center">
	<td>���ó</td>
	<td>��ũ</td>
	<td>����</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>PC��</td>
	<td><input type="text" value="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=itemid%>&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>" class="text" readonly style="width:100%" onfocus="this.select();"></td>
	<td><input type="button" onclick="fnCBCopy(<%=itemid%>,'w')" value="����"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>�������</td>
	<td><input type="text" value="http://m.10x10.co.kr/category/category_itemprd.asp?itemid=<%=itemid%>&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>" class="text" readonly style="width:100%" onfocus="this.select();"></td>
	<td><input type="button" onclick="fnCBCopy(<%=itemid%>,'m')" value="����"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>wishApp</td>
	<td><input type="text" value="http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp?itemid=<%=itemid%>&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>" class="text" readonly style="width:100%" onfocus="this.select();"></td>
	<td><input type="button" onclick="fnCBCopy(<%=itemid%>,'a')" value="����"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="3"><input type="button" value="â�ݱ�" onclick="window.close()"></td>
</tr>
</table>
<%
	set oitemcouponmaster = Nothing
	set ocouponitemlist = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->