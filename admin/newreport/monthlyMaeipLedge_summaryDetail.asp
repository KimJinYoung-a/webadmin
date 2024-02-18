<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls.asp"-->

<%
Dim CCADMIN : CCADMIN = C_ADMIN_AUTH
dim showItemDetailPopup : showItemDetailPopup = (left(now, 10) <= "2014-08-09") or CCADMIN


dim research, i, page
dim yyyy1,mm1, yyyymm1, makerid, showsuply, meaipTp, showShopid, showDiff
dim stockPlace, shopid
dim targetGbn, itemgubun
dim stype       '' S:���, J:����
dim bPriceGbn
dim showUpbae
dim PurchaseType
dim showPoint

page        = requestCheckvar(request("page"),10)
stype       = requestCheckvar(request("stype"),10)
shopid    	= requestCheckvar(request("shopid"),32)
research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1       	= requestCheckvar(request("mm1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsuply   = requestCheckvar(request("showsuply"),10)
showShopid  = requestCheckvar(request("showShopid"),10)
meaipTp     = requestCheckvar(request("meaipTp"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
showDiff   	= requestCheckvar(request("showDiff"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
showUpbae   = requestCheckvar(request("showUpbae"),10)
PurchaseType   = requestCheckvar(request("PurchaseType"),10)
showPoint   = requestCheckvar(request("showPoint"),10)

if (page="") then page=1
if (stockPlace="") then stockPlace="L"

if yyyy1 = "" then
	yyyy1 = "2013"
	mm1 = "12"
end if

if (research="") and (bPriceGbn = "") then
    bPriceGbn="P"
end if

yyyymm1 = yyyy1 + "-" + mm1

dim oCMonthlyMaeipLedge
set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge

oCMonthlyMaeipLedge.FRectYYYYMM = yyyymm1
oCMonthlyMaeipLedge.FRectStockPlace = stockPlace
oCMonthlyMaeipLedge.FRectShopid = shopid
oCMonthlyMaeipLedge.FRectMakerid = makerid
oCMonthlyMaeipLedge.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
oCMonthlyMaeipLedge.FRectMeaipTp = meaipTp
oCMonthlyMaeipLedge.FRectItemgubun = itemgubun
oCMonthlyMaeipLedge.FRectTargetGbn = targetGbn
oCMonthlyMaeipLedge.FRectShowShopid = showShopid
oCMonthlyMaeipLedge.FRectPriceGubun = bPriceGbn
oCMonthlyMaeipLedge.FRectShowUpbae = showUpbae

oCMonthlyMaeipLedge.FRectShowDiff = showDiff
oCMonthlyMaeipLedge.FRectShowPurchaseType = "Y"
oCMonthlyMaeipLedge.FRectPurchaseType = PurchaseType
oCMonthlyMaeipLedge.FRectShowPoint = showPoint

oCMonthlyMaeipLedge.FPageSize = 4000
oCMonthlyMaeipLedge.FCurrPage = page

if (stype="S") then
    oCMonthlyMaeipLedge.GetMaeipLedgeSUMSubDetail
else
    oCMonthlyMaeipLedge.GetMaeipJungsanSumSubDetail
end if


dim totprevSysStockNo, totprevSysStockSum, totIpgoNo, totIpgoSum, totSellNo, totSellSum, totOffChulNo, totOffChulSum, totEtcChulNo, totEtcChulSum
dim totCsNo, totCsSum, totLossChulNo, totLossChulSum, totcurSysStockNo, totcurSysStockSum, totcurErrRealCheckNo, totcurErrRealCheckSum
dim diff, totdiff
dim diffSum, totdiffSum
dim totMoveNo, totMoveSum
dim totErrNo, totErrSum

%>
<script language='javascript'>

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function GotoBrand(makerid){
	document.frm.makerid.value = makerid;
	document.frm.submit();
}

<% if (showItemDetailPopup = True) then %>
function PopItemStock(itemgubun, itemid, itemoption) {
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"PopItemStock","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
function PopItemStockShop(shopid, itemgubun, itemid, itemoption) {
	var barcode, formatLength;
	if (itemid*1 >= 1000000) {
		formatLength = 8;
	} else {
		formatLength = 6;
	}

	while (itemid.length < formatLength) {
		itemid = "0" + itemid;
	}

	barcode = itemgubun + itemid + itemoption;

	var popwin = window.open("/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=" + shopid + "&barcode=" + barcode,"PopItemStockShop","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
<% end if %>

function popAccStockModiOne(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/newreport/pop_item_stock_Accsummary_edit.asp?yyyy1=2015&mm1=03&shopid=&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popAccStockModiOne","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" target="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="stype" value="<%=stype%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %>
			&nbsp;&nbsp;
			<font color="#CC3333">�귣��:</font> <%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;&nbsp;
			���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >���ް��� ǥ��
			&nbsp;&nbsp;
			<input type="checkbox" name="showShopid" value="on" <%= CHKIIF(showShopid="on","checked","") %> >���� ǥ��
			<% if (CCADMIN) then %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showDiff" value="on" <%= CHKIIF(showDiff="on","checked","") %> > ���������� ǥ��
		    <% end if %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showUpbae" value="on" <%= CHKIIF(showUpbae="on","checked","") %> >�����ǰ�� ǥ��
            &nbsp;&nbsp;
			<input type="checkbox" name="showPoint" value="on" <%= CHKIIF(showPoint="on","checked","") %> >�Ҽ�����ǰ ǥ��
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.target='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <font color="#CC3333">�����ġ:</font>
		    <select name="stockPlace">
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >����</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >����</option>
				<option value="T" <%= CHKIIF(stockPlace="T","selected" ,"") %> >���</option>
				<option value="O" <%= CHKIIF(stockPlace="O","selected" ,"") %> >�¶�������</option>
				<option value="N" <%= CHKIIF(stockPlace="N","selected" ,"") %> >�¶��� ��������(�����Ұ�)</option>
				<option value="F" <%= CHKIIF(stockPlace="F","selected" ,"") %> >��������</option>
				<option value="A" <%= CHKIIF(stockPlace="A","selected" ,"") %> >�ΰŽ�����</option>
				<option value="E" <%= CHKIIF(stockPlace="E","selected" ,"") %> >����</option>
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">���Ա���:</font>
        	<select name="meaipTp">
        	<option value="">��ü
        	<option value="M" <%= CHKIIF(meaipTp="M","selected" ,"") %> >�԰�и���
        	<option value="S" <%= CHKIIF(meaipTp="S","selected" ,"") %> >�Ǹźи���
        	<option value="C" <%= CHKIIF(meaipTp="C","selected" ,"") %> >���и���
        	<option value="E" <%= CHKIIF(meaipTp="E","selected" ,"") %> >��Ÿ����
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">�μ�����:</font>
        	<input type="text" name="targetGbn" value="<%=targetGbn%>" size="2" maxlength="2">

        	&nbsp;&nbsp;
        	<font color="#CC3333">�ڵ屸��:</font>
        	<input type="text" name="itemgubun" value="<%=itemgubun%>" size="2" maxlength="2">
        	&nbsp;&nbsp;
			<font color="#CC3333">���԰�����:</font>
			<input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >�ۼ��ø��԰�
			<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >��ո��԰�
			&nbsp;&nbsp;
			<font color="#CC3333">��������:</font>
			<% drawPartnerCommCodeBox true,"purchasetype","PurchaseType",PurchaseType,"" %>
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

* ����޿� ����� �� ��� ������, �����޿� ����� �Ǵ� ��� �ִ� ��� ǥ�õ˴ϴ�.

<p />

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
<td>�� <%=oCMonthlyMaeipLedge.FTotalCount%> �� <%=page%>/<%=oCMonthlyMaeipLedge.FtotalPage%> page</td>
</tr>
</table>
<p>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="6">��ǰ����</td>
        <td colspan="2">�������(��������)</td>
        <td colspan="2">�������(��)</td>
        <td colspan="2">����̵�(��)</td>
        <td colspan="2">����Ǹ�(��)</td>
        <td colspan="2">������1(��)</td>
        <td colspan="2">������2(��)</td>
        <td colspan="2">�����Ÿ���(��)</td>
        <td colspan="2">���CS���(��)</td>
        <td colspan="2">����(��)</td>
		<td colspan="2"><b>�⸻���(��)</b></td>
		<td rowspan="2">������</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >����ID</td>
		<td >�귣��ID</td>
	    <td >��ǰ�ڵ�</td>
		<td >��������</td>
	    <td >���Ա���</td>
	    <td >�����ġ</td>
    	<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
		<td>����</td>
		<td>�ݾ�</td>
    </tr>
    <% for i=0 to oCMonthlyMaeipLedge.FResultCount-1 %>
    <%

	totprevSysStockNo       	= totprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo
	totprevSysStockSum       	= totprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum

	totIpgoNo       			= totIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo
	totIpgoSum       			= totIpgoSum + Round(oCMonthlyMaeipLedge.FItemList(i).getIpgoSum,0)

	totMoveNo       			= totMoveNo + oCMonthlyMaeipLedge.FItemList(i).getMoveNo
	totMoveSum       			= totMoveSum + oCMonthlyMaeipLedge.FItemList(i).getMoveSum

	totSellNo       			= totSellNo + oCMonthlyMaeipLedge.FItemList(i).FSellNo
	totSellSum       			= totSellSum + oCMonthlyMaeipLedge.FItemList(i).FSellSum

	totOffChulNo       			= totOffChulNo + oCMonthlyMaeipLedge.FItemList(i).FOffChulNo
	totOffChulSum       		= totOffChulSum + oCMonthlyMaeipLedge.FItemList(i).FOffChulSum

	totEtcChulNo       			= totEtcChulNo + oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo
	totEtcChulSum       		= totEtcChulSum + oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum

	totLossChulNo       		= totLossChulNo + oCMonthlyMaeipLedge.FItemList(i).FLossChulNo
	totLossChulSum       		= totLossChulSum + oCMonthlyMaeipLedge.FItemList(i).FLossChulSum

	totCsNo       				= totCsNo + oCMonthlyMaeipLedge.FItemList(i).FCsNo
	totCsSum       				= totCsSum + oCMonthlyMaeipLedge.FItemList(i).FCsSum

	totcurSysStockNo       		= totcurSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo
	totcurSysStockSum       	= totcurSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum
	totcurErrRealCheckNo       	= totcurErrRealCheckNo + oCMonthlyMaeipLedge.FItemList(i).FcurErrRealCheckNo
	totcurErrRealCheckSum       = totcurErrRealCheckSum + oCMonthlyMaeipLedge.FItemList(i).FcurErrRealCheckSum

	'diff = oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getMoveNo + oCMonthlyMaeipLedge.FItemList(i).FSellNo + oCMonthlyMaeipLedge.FItemList(i).FOffChulNo + oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo + oCMonthlyMaeipLedge.FItemList(i).FCsNo + oCMonthlyMaeipLedge.FItemList(i).FLossChulNo - oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo
	diff = oCMonthlyMaeipLedge.FItemList(i).getDiffNo
	diffSum = oCMonthlyMaeipLedge.FItemList(i).getDiffSum
	totdiff = totdiff + diff
	totdiffSum = totdiffSum + diffSum

	totErrNo = totErrNo + oCMonthlyMaeipLedge.FItemList(i).getTotErrNo
    totErrSum = totErrSum + oCMonthlyMaeipLedge.FItemList(i).getTotErrSum
    %>
    <tr align="right" bgcolor="#FFFFFF" onmouseover="this.style.background='F1F1F1'" onmouseout="this.style.background='FFFFFF'" >
		<td align="center">
		    <% if (showShopid<>"") then %>
		    	<%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %>
			<% elseif (shopid <> "") then %>
				<%= shopid %>
		    <% end if %>
		</td>
		<td align="center"><a href="javascript:GotoBrand('<%= oCMonthlyMaeipLedge.FItemList(i).FMakerid%>')"><%= oCMonthlyMaeipLedge.FItemList(i).FMakerid%></a></td>
		<td align="center">
		    <% if (makerid<>"") then %>
				<% if (showItemDetailPopup = True) then %>
					<% if (oCMonthlyMaeipLedge.FItemList(i).Fshopid <> "") then %>
						<a href="javascript:PopItemStockShop('<%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>');">
					<% else %>
						<a href="javascript:PopItemStock('<%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>');">
					<% end if %>
				<% end if %>
		        <%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>-<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>-<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>
		    <% else %>

		    <% end if %>
		</td>
		<td align="center"><%= getBrandPurchaseType(oCMonthlyMaeipLedge.FItemList(i).FpurchaseType) %></td>
        <td align="center"><%= oCMonthlyMaeipLedge.FItemList(i).getMeaipTypeName %></td>
        <td align="center"><%= oCMonthlyMaeipLedge.FItemList(i).FstockPlace%></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getIpgoNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getIpgoSum,0) %></td>

        <td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getMoveNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getMoveSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FSellNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FSellSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FOffChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FOffChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum,0) %></td>

		<td align="center"><img src="/images/icon_arrow_link.gif" style="cursor:pointer" onClick="popAccStockModiOne('<%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>')"></td>


    </tr>
    <% if FIx(i / 1000)=(i / 1000) then response.flush %>
	<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
		<td></td>
		<td></td>
		<td></td>
    	<td></td>
        <td></td>
    	<td align="right" ><%= FormatNumber(totprevSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(totprevSysStockSum,0) %></td>

		<td align="right" ><%= FormatNumber(totIpgoNo,0) %></td>
		<td align="right" ><%= FormatNumber(totIpgoSum,0) %></td>

		<td align="right" ><%= FormatNumber(totMoveNo,0) %></td>
		<td align="right" ><%= FormatNumber(totMoveSum,0) %></td>

		<td align="right" ><%= FormatNumber(totSellNo,0) %></td>
		<td align="right" ><%= FormatNumber(totSellSum,0) %></td>

		<td align="right" ><%= FormatNumber(totOffChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totOffChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totEtcChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totEtcChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totLossChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totLossChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totCsNo,0) %></td>
		<td align="right" ><%= FormatNumber(totCsSum,0) %></td>

		<td align="right" ><%= FormatNumber(totErrNo,0) %></td>
		<td align="right" ><%= FormatNumber(totErrSum,0) %></td>

		<td align="right" ><%= FormatNumber(totcurSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(totcurSysStockSum,0) %></td>

		<td></td>

    </tr>

	<tr height="25" bgcolor="FFFFFF">
	    <td><%=i%></td>
		<td colspan="26" align="center">
			<% if oCMonthlyMaeipLedge.HasPreScroll then %>
        		<a href="javascript:NextPage('<%= oCMonthlyMaeipLedge.StarScrollPage-1 %>')">[pre]</a>
        	<% else %>
        		[pre]
        	<% end if %>

        	<% for i=0 + oCMonthlyMaeipLedge.StarScrollPage to oCMonthlyMaeipLedge.FScrollCount + oCMonthlyMaeipLedge.StarScrollPage - 1 %>
        		<% if i>oCMonthlyMaeipLedge.FTotalpage then Exit for %>
        		<% if CStr(page)=CStr(i) then %>
        		<font color="red">[<%= i %>]</font>
        		<% else %>
        		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
        		<% end if %>
        	<% next %>

        	<% if oCMonthlyMaeipLedge.HasNextScroll then %>
        		<a href="javascript:NextPage('<%= i %>')">[next]</a>
        	<% else %>
        		[next]
        	<% end if %>
		</td>
	</tr>

</table>



<%
set oCMonthlyMaeipLedge = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
