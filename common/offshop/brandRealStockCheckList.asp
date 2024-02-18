<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������θ��� �귣�庰 ��� �ľ�
' History : 2011.08.01 �̻� ����
'			2019.05.31 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<%
dim shopid, makerid, research
	shopid       = RequestCheckVar(request("shopid"),32)
	makerid      = RequestCheckVar(request("makerid"),32)
	research     = RequestCheckVar(request("research"),32)

dim usingyn, centermwdiv ,NoZeroStock, comm_cd
	usingyn      = RequestCheckVar(request("usingyn"),32)
	centermwdiv  = RequestCheckVar(request("centermwdiv"),32)
	NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
	comm_cd      = RequestCheckVar(request("comm_cd"),32)

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

if (research="") then NoZeroStock="on"

dim oOffStock
set oOffStock = new CShopItemSummary
	oOffStock.FRectShopID       = shopid
	oOffStock.FRectMakerID      = makerid
	oOffStock.FRectComm_cd      = comm_cd
	oOffStock.FRectIsUsing      = usingyn
	oOffStock.FRectNoZeroStock  = NoZeroStock

	if (shopid<>"") then
	    oOffStock.GetShopBrandRealCheckRequire
	end if

dim i
dim sumTotItemNo            : sumTotItemNo=0
dim sumStPLusStockItemCnt   : sumStPLusStockItemCnt=0
dim sumTotSellNo            : sumTotSellNo=0
dim sumTotRealStockNo       : sumTotRealStockNo=0
dim sumTotStockBuySum       : sumTotStockBuySum=0
dim sumTotOwnStockBuySum    : sumTotOwnStockBuySum=0

%>

<script language='javascript'>

function popBrandStock(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTaking(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock_takingWithList.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTakingInput(stIdx){
    var popUrl = "/common/offshop/shop_brandcurrentstock_byjobkey.asp?idx="+stIdx+"&sType=stTaking";
    var popwin = window.open(popUrl,'popBrandStockInput','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function frmsumbit(page){
	frm.page.value=page;
	frm.action="";
	frm.target = "";
	frm.submit();
}

function jsCurrStockDown(stockPlace,temp){
	if (stockPlace==""){
		alert('�����ġ�� �������� �ʾҽ��ϴ�.');
		return;
	}
	frm.stockPlace.value=stockPlace;
	frm.action="/admin/newreport/currentstock_excel.asp";
	frm.target = "view";
	frm.submit();
	frm.target = "";
	frm.action = ""
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<input type="hidden" name="stockPlace" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				<input type="hidden" name="shopid" value="<%= shopid %>">
				* ���� : <%= shopid %>
				&nbsp;
				* �귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			<% else %>
	        	* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
				&nbsp;
				* �귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<%
			end if
		else
			''��ü�ΰ��
			if (C_IS_Maker_Upche) then
		%>
				* ���� : <% drawSelectBoxOpenOffShop "shopid",shopid %>
				<input type="hidden" name="makerid" value="<%= makerid %>">
		<%
			else
				if (C_ADMIN_USER) then
		%>
					* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					&nbsp;
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<%
				end if
			end if
		end if
		%>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsumbit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��ǰ ��뱸�� : <% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
	    * ������Ա��� : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
		&nbsp;
		<input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > ���0�� �귣�� �˻� ����.
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
<tr>
	<td align="left">
		* ���̳ʽ� ���� ���ݾ� 0 ���� ������.
	</td>
	<td align="right">
		* ���ǻ�(�˻����ǿ� ������ �����ϸ� �ش���� ��� �ٿ�ε�˴ϴ�.) :
		<!--
		<br><br><input type="checkbox" name="day1after">�������ĺ���������
		<input type="button" class="button" value="���ǻ�ٿ�ε�(����)" onclick="jsstockDown('S','');">
		-->
		<input type="button" class="button" value="�������ٿ�ε�(<%= CHKIIF(shopid="", "streetshop011", shopid) %>)" onclick="jsCurrStockDown('S','');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		�˻���� �� <%= oOffStock.FTotalCount %> ��
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"></td>
    <td width="110">�귣��ID</td>
	<td width="100">���Ա���</td>
	<td width="70">��ǰǰ���</td>
	<td width="70">���>0��ǰ��</td>
	<td width="70">���Ǹ�</td>
	<td width="70">�� �ǻ����</td>
	<td width="90">����
    	<% IF (C_IS_FRN_SHOP) then %>
    		<br>(������԰�)
    	<% else %>
    		<br>(������԰�)
    	<% end if %>
	</td>
	<td width="90">���� �԰���</td>
	<td width="90">���� �԰���</td>
    <td >�ֱٽǻ���</td>
    <td >��Ʈ(����)<br>����ľ�</td>
    <td >���ڵ�<br>����ľ�</td>
    <td >����Է�</td>
</tr>
<% if (shopid="") then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="14">[���� ���� �� �����ϼ���.]</td>
</tr>
<% else %>
<% for i=0 to oOffStock.FResultCount-1 %>
<%
sumTotItemNo            =sumTotItemNo+oOffStock.FItemList(i).FtotItemNo
sumStPLusStockItemCnt   =sumStPLusStockItemCnt+oOffStock.FItemList(i).FstPLusStockItemCnt
sumTotSellNo            =sumTotSellNo+oOffStock.FItemList(i).FtotSellNo*-1
sumTotRealStockNo       =sumTotRealStockNo+oOffStock.FItemList(i).FtotRealStockNo
if not isNULL(oOffStock.FItemList(i).FtotStockBuySum) then

    IF (C_IS_FRN_SHOP) then
        sumTotStockBuySum       =sumTotStockBuySum+oOffStock.FItemList(i).FtotStockBuySum
    else
        sumTotOwnStockBuySum    =sumTotOwnStockBuySum+oOffStock.FItemList(i).FtotOwnStockBuySum
    end if
end if
%>
<tr align="center" bgcolor="#FFFFFF">
    <td></td>
    <td><%= oOffStock.FItemList(i).Fmakerid %></td>
    <td><%= oOffStock.FItemList(i).Fcomm_name %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FtotItemNo,0) %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FstPLusStockItemCnt,0) %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FtotSellNo*-1,0) %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FtotRealStockNo,0) %></td>
    <td align="right">
	    <% if isNULL(oOffStock.FItemList(i).FtotStockBuySum) then %>
			<% if (UBound(Split(oOffStock.FItemList(i).Fmakerid, "-")) = 2) then %>
				<font color=red>��ǰ���� ����</font>
			<% else %>
				<font color=red>��� ����</font>
			<% end if %>
	    <% else %>
	        <% IF (C_IS_FRN_SHOP) then %>
	        <%= FormatNumber(oOffStock.FItemList(i).FtotStockBuySum,0) %>
	        <% else %>
	        <%= FormatNumber(oOffStock.FItemList(i).FtotOwnStockBuySum,0) %>
	        <% end if %>
	    <% end if %>
    </td>
    <td><%= oOffStock.FItemList(i).Ffirstipgodate %></td>
    <td><%= oOffStock.FItemList(i).Flastipgodate %></td>
    <td><%= oOffStock.FItemList(i).FlastStockdate %></td>
    <td><input type="button" class="button" value="���� �Է�" onClick="popBrandStock('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');"></td>
    <td>
	    <% if oOffStock.FItemList(i).FstStatus=0 then %>
			<input type="button" class="button_ing" value="��� �ľ� ��" onClick="popBrandStockTaking('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">
	    <% elseif oOffStock.FItemList(i).FstStatus=3 then %>
			<input type="button" class="button" disabled value="��� �ľ�" onClick="popBrandStockTaking('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">
	    <% else %>
			<input type="button" class="button" value="��� �ľ�" onClick="popBrandStockTaking('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">
	    <% end if %>
    </td>
    <td>
	    <% if oOffStock.FItemList(i).FstStatus=3 then %>
			<input type="button" class="button_ing" value="��� �Է�" onClick="popBrandStockTakingInput(<%= oOffStock.FItemList(i).FstTakingIdx %>);">
	    <% else %>
			<input type="button" class="button" disabled value="��� �Է�" onClick="popBrandStockTakingInput(<%= oOffStock.FItemList(i).FstTakingIdx %>);">
	    <% end if %>
    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td></td>
    <td></td>
    <td></td>
    <td><%= formatNumber(sumTotItemNo,0) %></td>
    <td><%= formatNumber(sumStPLusStockItemCnt,0) %></td>
    <td><%= formatNumber(sumTotSellNo,0) %></td>
    <td><%= formatNumber(sumTotRealStockNo,0) %></td>

    <% IF (C_IS_FRN_SHOP) then %>
		<td align="right"><%= formatNumber(sumTotStockBuySum,0) %></td>
    <% else %>
		<td align="right"><%= formatNumber(sumTotOwnStockBuySum,0) %></td>
    <% end if %>

    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
</tr>
<% end if %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set oOffStock = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
