<%@ language=vbscript %>
<% option explicit %>
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
dim yyyy1, yyyymm1, makerid, showsuply, meaipTp, showShopid, showDiff
dim stockPlace, shopid
dim targetGbn, itemgubun
dim bPriceGbn, showItem
dim stype       '' S:���, J:����

page        = requestCheckvar(request("page"),10)
stype       = requestCheckvar(request("stype"),10)
shopid    	= requestCheckvar(request("shopid"),32)
research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsuply   = requestCheckvar(request("showsuply"),10)
showShopid  = requestCheckvar(request("showShopid"),10)
meaipTp     = requestCheckvar(request("meaipTp"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
showDiff   	= requestCheckvar(request("showDiff"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
showItem	= requestCheckvar(request("showItem"),2)

if (page="") then page=1
if (stockPlace="") then stockPlace="L"

if (research="") and (bPriceGbn = "") then
    bPriceGbn = "V"
	showsuply = "on"
end if

if yyyy1 = "" then
	yyyy1 = "2012"
end if

dim oCMonthlyMaeipLedge
set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge

oCMonthlyMaeipLedge.FRectYYYY = yyyy1
oCMonthlyMaeipLedge.FRectStockPlace = stockPlace
oCMonthlyMaeipLedge.FRectShopid = shopid
oCMonthlyMaeipLedge.FRectMakerid = makerid
oCMonthlyMaeipLedge.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
oCMonthlyMaeipLedge.FRectMeaipTp = meaipTp
oCMonthlyMaeipLedge.FRectItemgubun = itemgubun
oCMonthlyMaeipLedge.FRectTargetGbn = targetGbn
oCMonthlyMaeipLedge.FRectShowShopid = showShopid
oCMonthlyMaeipLedge.FRectShowItem = showItem
oCMonthlyMaeipLedge.FRectPriceGubun = bPriceGbn

oCMonthlyMaeipLedge.FRectShowDiff = showDiff

oCMonthlyMaeipLedge.FPageSize = 1000
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
	var popwin = window.open("http://webadmin.10x10.co.kr/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"PopItemStock","width=1000 height=600 scrollbars=yes resizable=yes");
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

	var popwin = window.open("http://webadmin.10x10.co.kr/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=" + shopid + "&barcode=" + barcode,"PopItemStockShop","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
<% end if %>

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" target="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="stype" value="<%=stype%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">�⵵ :</font> 
			<% Call DrawyearBoxdynamic("yyyy1",yyyy1,"") %>
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
			&nbsp;&nbsp;
			<input type="checkbox" name="showItem" value="on" <%= CHKIIF(showItem="on","checked","") %> > ��ǰ�� ǥ��
		    <% end if %>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.target='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <font color="#CC3333">�����ġ:</font>
		    <select name="stockPlace">
		        <option value="" <%= CHKIIF(stockPlace="","selected" ,"") %> >��ü</option>
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >����</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >����</option>
				<option value="T" <%= CHKIIF(stockPlace="T","selected" ,"") %> >���</option>
				<option value=""">---------</option>
				<option value="O" <%= CHKIIF(stockPlace="O","selected" ,"") %> >�¶��� ��������</option>
				<option value="N" <%= CHKIIF(stockPlace="N","selected" ,"") %> >�¶��� ��������(�����Ұ�)</option>
				<option value="F" <%= CHKIIF(stockPlace="F","selected" ,"") %> >���� ��������</option>
				<option value="A" <%= CHKIIF(stockPlace="A","selected" ,"") %> >�ΰŽ� ��������</option>
				<option value="R" <%= CHKIIF(stockPlace="R","selected" ,"") %> >��Ż ��������</option>
				<option value=""">---------</option>
				<option value="E" <%= CHKIIF(stockPlace="E","selected" ,"") %> >����</option>
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">���Ա���:</font>
        	<select name="meaipTp">
        	<option value="">��ü
        	<option value="M" <%= CHKIIF(meaipTp="M","selected" ,"") %> >�԰��и���
        	<option value="S" <%= CHKIIF(meaipTp="S","selected" ,"") %> >�Ǹźи���
        	<option value="C" <%= CHKIIF(meaipTp="C","selected" ,"") %> >����и���
        	<option value="E" <%= CHKIIF(meaipTp="E","selected" ,"") %> >��Ÿ����
        	</select>
        	<% if (FALSE) then %>
        	&nbsp;&nbsp;
        	<font color="#CC3333">�μ�����:</font>
        	<input type="text" name="targetGbn" value="<%=targetGbn%>" size="2" maxlength="2">
            <% end if %>
        	&nbsp;&nbsp;
        	<font color="#CC3333">�ڵ屸��:</font>
        	<input type="text" name="itemgubun" value="<%=itemgubun%>" size="2" maxlength="2">
			&nbsp;&nbsp;
			<font color="#CC3333">���԰�����:</font>
			<input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >�ۼ��ø��԰�
			<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >��ո��԰�
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
	<td>�� <%=oCMonthlyMaeipLedge.FTotalCount%> �� <%=page%>/<%=oCMonthlyMaeipLedge.FtotalPage%> page</td>
</tr>
</table>
<p>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="5">��ǰ����</td>
        <td colspan="2">�������(�⸻����)</td>
        <td colspan="2">����(��)</td>
        <td colspan="2">�̵�(��)</td>
        <td colspan="2">�Ǹ�(��)</td>
        <td colspan="2">�������(��)</td>
        <td colspan="2">��Ÿ���(��)</td>
        <td colspan="2">�ν����(��)</td>
        <td colspan="2">CS���(��)</td>
        <td colspan="2">����</td>
		<td colspan="2"><b>�⸻���(�⸻)</b></td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >����ID</td>
		<td >�귣��ID</td>
	    <td >��ǰ�ڵ�</td>
	    <td >���Ա���</td>
	    <td >�����ġ</td>
    	<td width="50">����</td>
    	<td width="80">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="70">�ݾ�</td>
    	<td width="50">����</td>
    	<td width="80">�ݾ�</td>
    </tr>
    <% for i=0 to oCMonthlyMaeipLedge.FResultCount-1 %>
    <%

	totprevSysStockNo       	= totprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo
	totprevSysStockSum       	= totprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum

	totIpgoNo       			= totIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo
	totIpgoSum       			= totIpgoSum + oCMonthlyMaeipLedge.FItemList(i).getIpgoSum
	
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

	diff    = oCMonthlyMaeipLedge.FItemList(i).getDiffNo
    diffSum = oCMonthlyMaeipLedge.FItemList(i).getDiffSum
    
    totErrNo = totErrNo + oCMonthlyMaeipLedge.FItemList(i).getTotErrNo
    totErrSum = totErrSum + oCMonthlyMaeipLedge.FItemList(i).getTotErrSum
    
	totdiff = totdiff + diff
	totdiffSum = totdiffSum + diffSum
    %>
    <tr align="right" bgcolor="#FFFFFF" onmouseover="this.style.background='F1F1F1'" onmouseout="this.style.background='FFFFFF'" >
		<td align="center">
		    <% if (showShopid<>"") then %>
		        <%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %>
		    <% end if %>
		</td>
		<td align="center"><a href="javascript:GotoBrand('<%= oCMonthlyMaeipLedge.FItemList(i).FMakerid%>')"><%= oCMonthlyMaeipLedge.FItemList(i).FMakerid%></a></td>
		<td align="center">
		    <% if (makerid<>"") or (showItem<>"") then %>
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
    </tr>
	<% next %>
    <tr align="center" bgcolor="#FFFFFF">
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
    </tr>

	<tr height="25" bgcolor="FFFFFF">
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