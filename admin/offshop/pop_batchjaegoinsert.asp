<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���
' History : �̻� ����
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<%
dim shopid, idx
	shopid = requestCheckVar(request("shopid"),32)
	idx = requestCheckVar(request("idx"),10)

dim oshoporder
set oshoporder = new CShopOrder
oshoporder.FRectShopID = shopid
oshoporder.FRectIdx = idx
oshoporder.FPageSize = 1000
oshoporder.GetShopOrderDetail

dim i

%>



<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<form name="frm" method="post" action="return false;">
	<input type="hidden" name="shopid" value="<%= oshoporder.FItemList(i).Fjobshopid %>">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	        	<tr>
	        	  <td width="90">����ó :</td>
	        	  <td><%= oshoporder.FItemList(i).Fjobshopid %></td>
	        	</tr>
	        	<tr>
	        	  <td>����ó :</td>
	        	  <td><% SelectBoxOffShopSuplyer "suplyer", "-", "-", session("ssBctDiv") %></td>
	        	</tr>
	        	<tr>
	        	  <td>�԰��û�� :</td>
	        	  <td><input type=text name="yyyymmdd" value="" size=10 readonly > <a href="javascript:calendarOpen(frm.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a></td>
	        	</tr>
	        	<tr>
	        	  <td>��Ÿ��û���� :</td>
	        	  <td><textarea name=comment cols=60 rows=3></textarea></td>
	        	</tr>
	                </table>
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	          <input type=button name=tmp value=" �ֹ����ۼ� " onclick="SubmitInsert()">
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="30">NO</td>
      <td width="30">����</td>
      <td width="50">��ǰID</td>
      <td>��ǰ��</td>
      <td width="100">�ɼ�</td>
      <td width="50">�ǸŰ�</td>
      <td width="50">���԰�</td>
      <td width="50">����<br>����</td>
      <td width="50">���ް�</td>
      <td width="50">����<br>����</td>
      <td width="30">����</td>
    </tr>
<% for i=0 to oshoporder.FResultCount-1 %>
    <input type="hidden" name="itemgubun" value="<%= oshoporder.FItemList(i).Fitemgubun %>">
    <input type="hidden" name="itemid" value="<%= oshoporder.FItemList(i).Fitemid %>">
    <input type="hidden" name="itemoption" value="<%= oshoporder.FItemList(i).Fitemoption %>">
    <input type="hidden" name="itemno" value="<%= -1 * oshoporder.FItemList(i).Fitemno %>">
    <input type="hidden" name="sellcash" value="<%= oshoporder.FItemList(i).Frealsellprice %>">
    <input type="hidden" name="suplycash" value="<%= (oshoporder.FItemList(i).Frealsellprice * (1.0 - (oshoporder.FItemList(i).Fdefaultsuplymargin / 100))) %>">
    <input type="hidden" name="buycash" value="<%= (oshoporder.FItemList(i).Frealsellprice * (1.0 - (oshoporder.FItemList(i).Fdefaultmargin / 100))) %>">
    <input type="hidden" name="designer" value="<%= oshoporder.FItemList(i).Fmakerid %>">
    <tr align="center" bgcolor="#FFFFFF">
      <td><%= (i + 1) %></td>
      <td><%= oshoporder.FItemList(i).Fitemgubun %></td>
      <td><%= oshoporder.FItemList(i).Fitemid %></td>
      <td align="left"><%= oshoporder.FItemList(i).Fitemname %></td>
      <td align="left"><%= oshoporder.FItemList(i).Fitemoptionname %></td>
      <td align="right"><%= FormatNumber(oshoporder.FItemList(i).Frealsellprice,0) %></td>
      <td align="right"><%= FormatNumber((oshoporder.FItemList(i).Frealsellprice * (1.0 - (oshoporder.FItemList(i).Fdefaultmargin / 100))),0) %></td>
      <td><%= oshoporder.FItemList(i).Fdefaultmargin %></td>
      <td align="right"><%= FormatNumber((oshoporder.FItemList(i).Frealsellprice * (1.0 - (oshoporder.FItemList(i).Fdefaultsuplymargin / 100))),0) %></td>
      <td><%= oshoporder.FItemList(i).Fdefaultsuplymargin %></td>
      <td><%= -1 * oshoporder.FItemList(i).Fitemno %></td>
    </tr>
<% next %>
    </form>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<script>
function SubmitInsert() {
        var itemgubunarr, itemidarr, itemoptionarr, itemnoarr, sellcasharr, suplycasharr, buycasharr, designerarr;

        if (document.frm.yyyymmdd.value == "") {
                alert("�԰��û���� �����ϼ���.");
                return;
        }
        if (document.frm.suplyer.selectedIndex == 0) {
                alert("����ó�� �����ϼ���.");
                return;
        }

        itemgubunarr = "";
        itemidarr = "";
        itemoptionarr = "";
        itemnoarr = "";
        sellcasharr = "";
        suplycasharr = "";
        buycasharr = "";
        designerarr = "";
	for (var i=0;i<document.frm.elements.length;i++){
		if (document.frm.elements[i].name=="itemgubun"){
        	        itemgubunarr = itemgubunarr + document.frm.elements[i+0].value + "|";
        	        itemidarr = itemidarr + document.frm.elements[i+1].value + "|";
        	        itemoptionarr = itemoptionarr + document.frm.elements[i+2].value + "|";
        	        itemnoarr = itemnoarr + document.frm.elements[i+3].value + "|";
        	        sellcasharr = sellcasharr + document.frm.elements[i+4].value + "|";
        	        suplycasharr = suplycasharr + document.frm.elements[i+5].value + "|";
        	        buycasharr = buycasharr + document.frm.elements[i+6].value + "|";
        	        designerarr = designerarr + document.frm.elements[i+7].value + "|";
      	        }
	}

	if (itemgubunarr==""){
		alert('������ ��ǰ�� �����ϴ�.');
		return;
	}

        document.frmArrupdate.yyyymmdd.value = document.frm.yyyymmdd.value;
        document.frmArrupdate.baljuid.value = document.frm.shopid.value;
        document.frmArrupdate.targetid.value = document.frm.suplyer.value;
        document.frmArrupdate.reguser.value = document.frm.shopid.value;
        document.frmArrupdate.comment.value = document.frm.comment.value;

        document.frmArrupdate.itemgubunarr.value = itemgubunarr;
        document.frmArrupdate.itemidarr.value = itemidarr;
        document.frmArrupdate.itemoptionarr.value = itemoptionarr;
        document.frmArrupdate.itemnoarr.value = itemnoarr;
        document.frmArrupdate.sellcasharr.value = sellcasharr;
        document.frmArrupdate.suplycasharr.value = suplycasharr;
        document.frmArrupdate.buycasharr.value = buycasharr;
        document.frmArrupdate.designerarr.value = designerarr;

        document.frmArrupdate.submit();
}
</script>
<form name="frmArrupdate" method="post" action="batchjaegoinsert_process.asp">
<input type="hidden" name="mode" value="addshopjumun">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="yyyymmdd" value="">
<input type="hidden" name="baljuid" value="">
<input type="hidden" name="targetid" value="">
<input type="hidden" name="reguser" value="">
<input type="hidden" name="divcode" value="503">
<input type="hidden" name="vatinclude" value="Y">
<input type="hidden" name="comment" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="designerarr" value="">
</form>
<%
set oshoporder = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->