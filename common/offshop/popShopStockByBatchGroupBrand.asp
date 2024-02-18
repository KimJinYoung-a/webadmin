<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������θ��� �귣�庰 ��� �ľ� (��ġ)
' History : 2011.08
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->

<%
dim idx : idx = RequestCheckVar(request("idx"),32)
dim oshopBatch, shopid, jobkey, joborderno, StockDate
set oshopBatch = new CShopOrder
	oshopBatch.FRectidx=idx
	
	if (idx<>"") then
	    oshopBatch.GetOneShopBatchOrder
	end if

if (oshopBatch.FResultCount>0) then
    shopid = oshopBatch.FOneItem.Fjobshopid
    jobkey = oshopBatch.FOneItem.Fjobkey
    joborderno = oshopBatch.FOneItem.Forderno
    StockDate = Left(oshopBatch.FOneItem.FShopRegDate,10)
end if 

set oshopBatch= Nothing

'', research,NoZeroStock,centermwdiv
''dim makerid : makerid      = RequestCheckVar(request("makerid"),32)
''research     = RequestCheckVar(request("research"),32)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if
''if (research="") then NoZeroStock="on"

dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FRectShopID = shopid
oOffStock.FRectBatchIdx = idx
if (shopid<>"") then
    oOffStock.GetShopBrandBatchCheckList
end if

dim i
%>
<script language='javascript'>
function popBrandStock(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock_byJobKey.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&idx=<%= idx %>";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="idx" value="<%= idx %>">
	
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    ���� : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- ���� ��ü -->
    		    ���� : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        ���� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		    <% end if %>
		    
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    �۾���ȣ : <%= jobkey %> (<%= joborderno %>) 
		    ����Ͻ� : <%= StockDate %>
		    
			
		</td>
	</tr>
	
	</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" > 
    <tr height="30">
        <td>
        * ���̳ʽ� ���� ���ݾ� 0 ���� ������.    
        </td>
    </tr>
	<tr height="30">
		<td align="left">
			�˻���� �� <%= oOffStock.FTotalCount %> ��
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="20"></td>
        <td width="120">�귣��ID</td>
    	<td width="100">���Ա���</td>
    	<td width="100">�� �ý��ۻ�<br>�ǻ� ���</td>
    	<td width="100">��� �ľ� �ǻ����</td>
    	<td width="100">���� �԰���</td>
        <td >�ֱٽǻ���</td>
        <td >�귣��ǻ�</td>
    </tr>
    <% if (shopid="") then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <td colspan="10">[���� ���� �� �����ϼ���.]</td>
    </tr>
    <% else %>
    <% for i=0 to oOffStock.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td></td>
        <td><%= oOffStock.FItemList(i).Fmakerid %></td>
        <td><%= oOffStock.FItemList(i).Fcomm_name %></td>
        <td><%= FormatNumber(oOffStock.FItemList(i).FtotSysRealStockNo,0) %></td>
        <td><%= FormatNumber(oOffStock.FItemList(i).FtotRealStockNo,0) %></td>
        
        <td><%= oOffStock.FItemList(i).Ffirstipgodate %></td>
        <td><%= oOffStock.FItemList(i).FlastStockdate %></td>
        <td>
        <input type="button" class="button" value="�귣�� �ǻ� �Է�" onClick="popBrandStock('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">    
        </td>
    </tr>
    <% next %>
    <% end if %>
</table>
<%
set oOffStock = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" --> 