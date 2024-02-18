<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

Dim isViewUser : isViewUser = FALSE ''(session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal, research, newitem, minusinc, bPriceGbn, vatyn, designer, monthGubun
dim mwgubun, buseo, itemgubun, mygubun
dim purchasetype, stplace, shopid, swSppPrc, etcjungsantype
Dim byall, page, ordTp
dim dispCate
dim showUpbae
dim startMon, endMon

yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
isusing     = requestCheckvar(request("isusing"),10)
sysorreal   = requestCheckvar(request("sysorreal"),10)
research    = requestCheckvar(request("research"),10)
newitem     = requestCheckvar(request("newitem"),10)
mwgubun     = requestCheckvar(request("mwgubun"),10)
mygubun     = requestCheckvar(request("mygubun"),10)
minusinc   	= requestCheckvar(request("minusinc"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
buseo       = requestCheckvar(request("buseo"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
purchasetype   	= requestCheckvar(request("purchasetype"),10)
vatyn       	= requestCheckvar(request("vatyn"),10)
designer       	= requestCheckvar(request("designer"),32)
monthGubun      = requestCheckvar(request("monthGubun"),32)
stplace       	= requestCheckvar(request("stplace"),10)
shopid       	= requestCheckvar(request("shopid"),32)
swSppPrc	= requestCheckvar(request("swSppPrc"),32)
byall       = requestCheckvar(request("byall"),32)
page        = requestCheckvar(request("page"),10)
ordTp   	= RequestCheckVar(request("ordTp"),10)
etcjungsantype  = requestCheckvar(request("etcjungsantype"),10)
dispCate 		= requestCheckvar(request("disp"),16)
showUpbae 		= requestCheckvar(request("showUpbae"),16)
startMon     	= RequestCheckVar(request("startMon"),32)
endMon     		= RequestCheckVar(request("endMon"),32)

if (sysorreal="") then sysorreal="sys"  ''real
if (isViewUser="") then sysorreal="sys"
if (isViewUser="") then bPriceGbn="P"
if (isViewUser="") then isusing=""
''if (monthGubun="") then monthGubun="1"
if (research="") and (etcjungsantype="") then etcjungsantype="41" ''����+�Ǹź�

if (page="") then page=1

if (bPriceGbn="") then
    bPriceGbn="P"
end if

if (mygubun = "") then
	mygubun = "M"
end if

if (stplace="") then
    stplace="L"
end if

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FCurrPage = page
ojaego.FPageSize = 1000
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectGubun = sysorreal
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectItemGubun = itemgubun
ojaego.FRectPurchaseType = purchasetype
ojaego.FRectTargetGbn = buseo
ojaego.FRectVatYn    = vatyn
ojaego.FRectMakerid = designer
ojaego.FRectMonthGubun = monthGubun
ojaego.FRectShopID    = shopid
ojaego.FRectShopSuplyPrice    = swSppPrc
ojaego.FRectOrdTp = ordTp
ojaego.FRectetcjungsantype = etcjungsantype
ojaego.FRectPriceGubun = bPriceGbn
ojaego.FRectDispCate		= dispCate

if IsNumeric(startMon) then
	ojaego.FRectStartDate = startMon
elseif (startMon <> "") then
	response.write "<script>alert('������ ���ڸ� �����մϴ�. " & startMon & "')</script>"
end if
if IsNumeric(endMon) then
	ojaego.FRectEndDate = endMon
elseif (endMon <> "") then
	response.write "<script>alert('������ ���ڸ� �����մϴ�. " & endMon & "')</script>"
end if

if (byall<>"") then
    ojaego.FRectShowItemList = "on"
end if

if (stplace = "L") then
	ojaego.FRectShowUpbae		= showUpbae
	ojaego.GetJeagoOverValueDetailSum
else
	''ojaego.GetJeagoOverValueSum_Shop
	ojaego.FRectLastIpgoGBN = stplace
	ojaego.GetJeagoOverValueDetailSum_Shop
end if



''response.end

dim i

dim subTotBuySum1, subTotBuySum2, subTotBuySum3, subTotBuySum4, subTotBuySum5, subTotBuySum6, subTotBuySum7, subTotBuySum8, subTotBuySum11, subTotBuySum12, subTotBuySum13, subTotBuySum14, subTotBuySum
dim totBuySum1, totBuySum2, totBuySum3, totBuySum4, totBuySum5, totBuySum6, totBuySum7, totBuySum8, totBuySum11, totBuySum12, totBuySum13, totBuySum14, totBuySum
dim totStockSum

dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy

dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy


dim iURL
dim nBusiName

Dim isItemListType : isItemListType=(designer<>"")or(byall<>"")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function changeScroll(comp){
    document.frm.page.value=comp.value;
    document.frm.submit();
}

function jsSearchByList(){
    document.frm.byall.value="on";
    document.frm.target="_blank";
    document.frm.submit();
    document.frm.byall.value="";
    document.frm.target="";
}

function jsSearchBrand(makerid, monthGubun) {
	document.frm.designer.value = makerid;
	document.frm.monthGubun.value = monthGubun;
	document.frm.submit();
}

function jsSearchItemStock(shopid,itemgubun,itemid,itemoption) {
<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
	if (shopid==''){
	alert("���� ������ �����ϼ���.");
	return;
	}
<% end if %>
	<% if (stplace = "S") or (stplace = "T") or (stplace = "M")  then %>
	var popwin = window.open("/common/offshop/shop_itemcurrentstock.asp?shopid="+shopid+"&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"jsSearchItemStock","width=1000 height=600 scrollbars=yes resizable=yes");
	<% else %>
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"jsSearchItemStock","width=1200 height=600 scrollbars=yes resizable=yes");
	<% end if %>

	popwin.focus();
}

function jsSearchItemStockLgs(itemgubun,itemid,itemoption) {
    var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"jsSearchItemStock","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//���Ա��� ������ ����
function TnPopItemStockModifyMW(itemgubun,itemid,itemoption) {
	var popwin = window.open("pop_item_stock_edit.asp?yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"TnPopItemStockModifyMW","width=600 height=300 scrollbars=yes resizable=yes");
	popwin.focus();
}

function TnPopItemStockModifyNull(itemgubun,itemid,itemoption) {
	var popwin = window.open("pop_item_stock_edit_lastIpgo.asp?yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"TnPopItemStockModifyNull","width=600 height=300 scrollbars=yes resizable=yes");
	popwin.focus();
}

function TnPopItemStockModifyLastIpgo(yyyymm, stockPlace, shopid, itemgubun,itemid,itemoption) {
	var popwin = window.open("popAssignMonthlyAccLastIpgo.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"TnPopItemStockModifyLastIpgo","width=500 height=230 scrollbars=yes resizable=yes");

	if (((stockPlace == "S") || (stockPlace == "T")) && (shopid == "")) {
		alert("���� �˻� �� �Է°����մϴ�.");
		return;
	}

	popwin.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="byall" value="<%=byall%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
	        &nbsp;&nbsp;|&nbsp;&nbsp;
			��������
			<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >��ü
			<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >����
			<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >�鼼
			&nbsp;&nbsp;
			<input type="checkbox" name="swSppPrc" value="Y" <%= CHKIIF(swSppPrc="Y","checked","") %> >���ް��� ǥ��
    		�귣�� : <% drawSelectBoxDesignerwithName "designer", designer %>
			&nbsp;&nbsp;
			����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			&nbsp;&nbsp;
			<input type="checkbox" name="showUpbae" value="on" <%= CHKIIF((showUpbae="on" and stplace="L"),"checked","") %> <%= CHKIIF(stplace<>"L","disabled","") %> >�����ǰ�� ǥ��
		</td>

		<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<% if (Not isViewUser) then %>
			<font color="#CC3333">�����:</font>
			<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >�ý������
			<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >�ǻ����
			&nbsp;&nbsp;
			<% end if %>

        	<font color="#CC3333">���Ա���:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >����
			<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >��Ź
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >������

        	<input type="radio" name="mwgubun" value="B031" <% if mwgubun="B031" then response.write "checked" %> >������
        	<input type="radio" name="mwgubun" value="B021" <% if mwgubun="B021" then response.write "checked" %> >��������
        	<input type="radio" name="mwgubun" value="B022" <% if mwgubun="B022" then response.write "checked" %> >�������
        	<input type="radio" name="mwgubun" value="B013" <% if mwgubun="B013" then response.write "checked" %> >�����Ź
        	<input type="radio" name="mwgubun" value="B012" <% if mwgubun="B012" then response.write "checked" %> >��ü��Ź


		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		<font color="#CC3333">���̳ʽ�����:</font>
		<input type="radio" name="minusinc" value="" <%= CHKIIF(minusinc="","checked","") %> >���̳ʽ���� ����(��ü)
	    &nbsp;&nbsp;
	    <% if (Not isViewUser) then %>
	    <font color="#CC3333">���԰�����:</font>
	    <input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >�ۼ��ø��԰�
		<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >��ո��԰�
		&nbsp;
	    <font color="#CC3333">�����Ⱓ:</font>
	    <input type="radio" name="mygubun" value="M" <%= CHKIIF(mygubun="M","checked","") %>  >����
		<input type="radio" name="mygubun" value="Y" <%= CHKIIF(mygubun="Y","checked","") %>  >������
	    <% end if %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">�����ġ:</font>
	        <select name="stplace">
        		<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >����</option>
				<option value="M" <%= CHKIIF(stplace="M","selected" ,"") %> >����(���Ա��к�)</option>
				<option value="">---------</option>
				<option value="T" <%= CHKIIF(stplace="T","selected" ,"") %> >����(�����԰���)</option>
				<option value="S" <%= CHKIIF(stplace="S","selected" ,"") %> >����(�����԰���)</option>
        	</select>
			&nbsp;
	    	<font color="#CC3333">�μ�����:</font>
	        <select name="buseo">
        	<option value="" <%= CHKIIF(buseo="","selected" ,"") %> >��ü
        	<option value="ON" <%= CHKIIF(buseo="ON","selected" ,"") %> >�¶���
        	<option value="OF" <%= CHKIIF(buseo="OF","selected" ,"") %> >��������
        	<option value="IT" <%= CHKIIF(buseo="IT","selected" ,"") %> >���̶��(��)
        	<option value="ET" <%= CHKIIF(buseo="ET","selected" ,"") %> >���
        	<option value="EG" <%= CHKIIF(buseo="EG","selected" ,"") %> >���׷���
        	</select>
			&nbsp;
	    	<font color="#CC3333">��ǰ����:</font>
        	<select name="itemgubun">
        	<option value="" <%= CHKIIF(itemgubun="","selected" ,"") %> >��ü
        	<option value="10" <%= CHKIIF(itemgubun="10","selected" ,"") %> >�Ϲ�(10)
        	<option value="70" <%= CHKIIF(itemgubun="70","selected" ,"") %> >����ǰ(70)
        	<option value="85" <%= CHKIIF(itemgubun="85","selected" ,"") %> >����ǰ(85)
        	<option value="80" <%= CHKIIF(itemgubun="80","selected" ,"") %> >����ǰ(80)
        	<option value="90" <%= CHKIIF(itemgubun="90","selected" ,"") %> >��������(90)
			<option value="00" <%= CHKIIF(itemgubun="00","selected" ,"") %> >ERR(00)
        	</select>
			&nbsp;
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
			&nbsp;
			������:
			<!-- �ϴ��� ���ܵд�. skyer9, 2017-08-30
        	<select name="monthGubun">
				<option value="" <%= CHKIIF(monthGubun="","selected" ,"") %> >��ü</option>
				<% if (mygubun = "Y") then %>
				<option value="11" <%= CHKIIF(monthGubun="11","selected" ,"") %> ><%= (yyyy1 - 0) %></option>
				<option value="12" <%= CHKIIF(monthGubun="12","selected" ,"") %> ><%= (yyyy1 - 1) %></option>
				<option value="13" <%= CHKIIF(monthGubun="13","selected" ,"") %> ><%= (yyyy1 - 2) %></option>
				<option value="14" <%= CHKIIF(monthGubun="14","selected" ,"") %> >~ <%= (yyyy1 - 3) %></option>
				<% else %>
        		<option value="1" <%= CHKIIF(monthGubun="1","selected" ,"") %> >1����~3����</option>
				<option value="2" <%= CHKIIF(monthGubun="2","selected" ,"") %> >4����~6����</option>
				<option value="3" <%= CHKIIF(monthGubun="3","selected" ,"") %> >7����~12����</option>
				<option value="4" <%= CHKIIF(monthGubun="4","selected" ,"") %> >1��~2��</option>
				<option value="7" <%= CHKIIF(monthGubun="7","selected" ,"") %> >13����~18����</option>
				<option value="8" <%= CHKIIF(monthGubun="8","selected" ,"") %> >19����~24����</option>
				<option value="5" <%= CHKIIF(monthGubun="5","selected" ,"") %> >2���ʰ�</option>
				<% end if %>
				<option value="6" <%= CHKIIF(monthGubun="6","selected" ,"") %> >NULL</option>
        	</select>
			or
			-->
			<input type="text" class="text" name="startMon" size="2" value="<%= startMon %>">
			~
			<input type="text" class="text" name="endMon" size="2" value="<%= endMon %>"> ����

			<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
				&nbsp;
				����(������� �˻���) : <% Call drawSelectBoxAccShop(yyyy1 + "-" + mm1, designer, "shopid", shopid) %>

				������:
				<select class="select" name="etcjungsantype"  >
                <option value="">-����-</option>
                <option value="1" <%=CHKIIF(etcjungsantype="1","selected","")%> >�Ǹź�����</option>
                <option value="2" <%=CHKIIF(etcjungsantype="2","selected","")%> >��������</option>
                <option value="3" <%=CHKIIF(etcjungsantype="3","selected","")%> >����������</option>
                <option value="4" <%=CHKIIF(etcjungsantype="4","selected","")%> >����������</option>
                <option value="41" <%=CHKIIF(etcjungsantype="41","selected","")%> >������+�Ǹź�����</option>
                </select>
			<% end if %>
			<% if (byall="") then %>
			<input type="button" class="button" value="��ǰ����Ʈ����" onClick='jsSearchByList();'">
			<% end if %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	        <font color="#CC3333">���ı���:</font>
	        <% if designer<>"" then %>
	        <input type="radio" name="ordTp" value="" <%= CHKIIF(ordTp="","checked","") %> >��ǰ�ڵ�
	        <input type="radio" name="ordTp" value="S" <%= CHKIIF(ordTp="S","checked","") %> >������
	        <% else %>
	        <input type="radio" name="ordTp" value="" <%= CHKIIF(ordTp="","checked","") %> >�⺻
	        <% end if %>
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<% if (ojaego.FResultcount>0) then %>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td>
        <% if (isItemListType) then %>
        �� <%=formatNumber(ojaego.FTotalCount,0)%>��, �˻���� <%=formatNumber(ojaego.FResultcount,0)%>��,
        <select name="npage" onChange="changeScroll(this);">
        <% for i=0 to ojaego.FTotalPage-1 %>
        <option value="<%=i+1%>" <%=CHKIIF(i+1=CLNG(page),"selected","")%>><%=i+1%>
        <% next %>
        </select>
        /<%=ojaego.FTotalPage%> Page
        <% else %>
        �˻���� <%=formatNumber(ojaego.FResultcount,0)%>��
        <% end if %>
    </td>
</tr>
</table>
<p>
<% end if %>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="4">��ǰ����</td>
		<td rowspan="2">��������</td>
		<td rowspan="2">�귣��</td>
		<% if (IsItemListType) then %>
		<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
		<td rowspan="2">����</td>
		<% end if %>
		<td rowspan="2">��ǰ�ڵ�</td>
		<td rowspan="2">�ɼ�</td>
		<td rowspan="2">��ǰ��</td>
		<td rowspan="2">�ɼǸ�</td>
		<td rowspan="2">�ܰ�</td>
		<td rowspan="2">�Ѽ���</td>
		<% else %>
		<td rowspan="2">�Ѽ���</td>
		<% end if %>

		<% if (mygubun = "Y") then %>
		<td rowspan="2" width="100"><%= yyyy1 %></td>
		<td rowspan="2" width="100"><%= (yyyy1 - 1) %></td>
		<td rowspan="2" width="100"><%= (yyyy1 - 2) %></td>
		<td rowspan="2" width="100">~ <%= (yyyy1 - 3) %></td>
		<% else %>
		<td rowspan="2" width="80">1����~3����</td>
		<td rowspan="2" width="80">4����~6����</td>
		<td rowspan="2" width="80">7����~12����</td>
		<td rowspan="2" width="80">13����~18����</td>
		<td rowspan="2" width="80">19����~24����</td>
		<td rowspan="2" width="80">2���ʰ�</td>
		<% end if %>

		<td rowspan="2" width="80">NULL</td>
		<td rowspan="2" width="100">�Ѱ�</td>
		<td rowspan="2" width="80">�����԰�</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td >�μ�</td>
	    <td >����</td>
	    <td >�ڵ�<br>����</td>
    	<td >����<br>����</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
	<%
    totStockSum = totStockSum + ojaego.FItemList(i).FtotStockNo
	totBuySum1 = totBuySum1 + ojaego.FItemList(i).FTotBuySum1
	totBuySum2 = totBuySum2 + ojaego.FItemList(i).FTotBuySum2
	totBuySum3 = totBuySum3 + ojaego.FItemList(i).FTotBuySum3
	totBuySum4 = totBuySum4 + ojaego.FItemList(i).FTotBuySum4
	totBuySum5 = totBuySum5 + ojaego.FItemList(i).FTotBuySum5
	totBuySum6 = totBuySum6 + ojaego.FItemList(i).FTotBuySum6
	totBuySum7 = totBuySum7 + ojaego.FItemList(i).FTotBuySum7
	totBuySum8 = totBuySum8 + ojaego.FItemList(i).FTotBuySum8

	totBuySum11 = totBuySum11 + ojaego.FItemList(i).FTotBuySum11
	totBuySum12 = totBuySum12 + ojaego.FItemList(i).FTotBuySum12
	totBuySum13 = totBuySum13 + ojaego.FItemList(i).FTotBuySum13
	totBuySum14 = totBuySum14 + ojaego.FItemList(i).FTotBuySum14

	totBuySum = totBuySum + ojaego.FItemList(i).FTotBuySum

	%>
<% if (ojaego.FResultcount>1000) or (byall<>"") then %>
<tr align="center" bgcolor="#FFFFFF">
<td><%= ojaego.FItemList(i).getBusiName %></td><td><%= ojaego.FItemList(i).getITemGubunName %></td><td><%= ojaego.FItemList(i).Fitemgubun %></td>
<td><% if (IsItemListType) and ((stplace = "S") or (stplace = "T") or (stplace = "M")) then %><%= ojaego.FItemList(i).getLastCommCD %><% else %><%= ojaego.FItemList(i).getMaeipGubunName %><% end if %></td>
<td><%= ojaego.FItemList(i).FpurchasetypeStr %></td><td><%= ojaego.FItemList(i).Fmakerid %></td>
<% if (IsItemListType) then %>
<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %><td><%= ojaego.FItemList(i).Fshopid %></td><% end if %>
<td><%= ojaego.FItemList(i).Fitemid %></td>
<td><%= ojaego.FItemList(i).Fitemoption %></td>
<td><%= ojaego.FItemList(i).Fitemname %></td>
<td><%= ojaego.FItemList(i).Fitemoptionname %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FbuyPrice, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FtotStockNo, 0) %></td>
<% end if %>
<% if (mygubun = "Y") then %>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum11, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum12, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum13, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum14, 0) %></td>
<% else %>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum1, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum2, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum3, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum7, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum8, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum5, 0) %></td>
<% end if %>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum6, 0) %></td>
<td><%= FormatNumber(ojaego.FItemList(i).FTotBuySum, 0) %></td>
<td><%= ojaego.FItemList(i).GetlastIpgoDate %></td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= ojaego.FItemList(i).getBusiName %></td>
    <td><%= ojaego.FItemList(i).getITemGubunName %></td>
    <td><%= ojaego.FItemList(i).Fitemgubun %></td>
	<td>
		<% if (IsItemListType) and ((stplace = "S") or (stplace = "T") or (stplace = "M")) then %>
			<%= ojaego.FItemList(i).getLastCommCD %>
		<% else %>
			<%= ojaego.FItemList(i).getMaeipGubunName %>
		<% end if %>
	</td>
	<td align="left"><%= ojaego.FItemList(i).FpurchasetypeStr %></td>
	<% if (IsItemListType) then %>
	<td align="left"><%= ojaego.FItemList(i).Fmakerid %></td>
	<% else %>
	<td align="left"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '<%= monthGubun %>')"><%= ojaego.FItemList(i).Fmakerid %></a></td>
	<% end if %>
	<% if (IsItemListType) then %>
	<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
	<td><%= ojaego.FItemList(i).Fshopid %></td>
	<% end if %>
	<td><a href="javascript:jsSearchItemStock('<%= ojaego.FItemList(i).Fshopid %>', '<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')"><%= ojaego.FItemList(i).Fitemid %></a></td>
	<td><a href="javascript:jsSearchItemStock('<%= ojaego.FItemList(i).Fshopid %>', '<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')"><%= ojaego.FItemList(i).Fitemoption %></a></td>
	<td align="left"><a href="javascript:jsSearchItemStock('<%= ojaego.FItemList(i).Fshopid %>', '<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')"><%= ojaego.FItemList(i).Fitemname %></a></td>
	<td align="left"><%= ojaego.FItemList(i).Fitemoptionname %></td>
	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FbuyPrice, 0) %></td>
	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FtotStockNo, 0) %></td>
	<% else %>
	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FtotStockNo, 0) %></td>
	<% end if %>

	<% if (mygubun = "Y") then %>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '11')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum11, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '12')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum12, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '13')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum13, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '14')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum14, 0) %></a></td>
	<% else %>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '1')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum1, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '2')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum2, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '3')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum3, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '7')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum7, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '8')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum8, 0) %></a></td>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '5')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum5, 0) %></a></td>
	<% end if %>
	<td align="right"><a href="javascript:jsSearchBrand('<%= ojaego.FItemList(i).Fmakerid %>', '6')"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum6, 0) %></a></td>
	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum, 0) %></td>

	<% if (byall="") and ((stplace = "T") or (stplace = "M")) then %>
	<td align="center">
		<a href="javascript:TnPopItemStockModifyLastIpgo('<%= yyyy1 & "-" & mm1 %>', '<%= stplace %>', '<%= ojaego.FItemList(i).Fshopid %>', '<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')">
			<% if (ojaego.FItemList(i).FlastIpgoDate >= "11") and (ojaego.FItemList(i).FlastIpgoDate <= "14") then %>
			<%
			Select Case ojaego.FItemList(i).FlastIpgoDate
				Case "11"
					response.write (yyyy1)
				Case "12"
					response.write (yyyy1 - 1)
				Case "13"
					response.write (yyyy1 - 2)
				Case "14"
					response.write "~ " & (yyyy1 - 3)
				Case Else
					ojaego.FItemList(i).FlastIpgoDate
			End Select
			%>
			<% elseif isNULL(ojaego.FItemList(i).FlastIpgoDate) or ojaego.FItemList(i).FlastIpgoDate = "" then %>
			NULL
			<% else %>
			<%= ojaego.FItemList(i).GetlastIpgoDate %>
			<% end if %>
		</a>
	</td>
	<% else %>
	<td align="center">
	<% if (ojaego.FItemList(i).FlastIpgoDate >= "11") and (ojaego.FItemList(i).FlastIpgoDate <= "14") then %>
		<%
		Select Case ojaego.FItemList(i).FlastIpgoDate
			Case "11"
				response.write (yyyy1)
			Case "12"
				response.write (yyyy1 - 1)
			Case "13"
				response.write (yyyy1 - 2)
			Case "14"
				response.write "~ " & (yyyy1 - 3)
			Case Else
				ojaego.FItemList(i).FlastIpgoDate
		End Select
		%>
	<% elseif isNULL(ojaego.FItemList(i).GetlastIpgoDate) or (ojaego.FItemList(i).GetlastIpgoDate="") or isNULL(ojaego.FItemList(i).FlastIpgoDate) then %>
	<!--
	<a href="javascript:TnPopItemStockModifyNull('<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')">NULL</a>
	-->
	<a href="javascript:TnPopItemStockModifyLastIpgo('<%= yyyy1 & "-" & mm1 %>', '<%= stplace %>', '<%= ojaego.FItemList(i).Fshopid %>', '<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')">NULL</a>
	<% else %>
	<a href="javascript:TnPopItemStockModifyLastIpgo('<%= yyyy1 & "-" & mm1 %>', '<%= stplace %>', '<%= ojaego.FItemList(i).Fshopid %>', '<%= ojaego.FItemList(i).Fitemgubun %>', '<%= ojaego.FItemList(i).Fitemid %>', '<%= ojaego.FItemList(i).Fitemoption %>')">
	<%= ojaego.FItemList(i).GetlastIpgoDate %>
	</a>
	<% end if %>
	</td>
	<% end if %>
</tr>
<% end if %>
	<% next %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
    	<td>�Ѱ�</td>
    	<td></td>
    	<td></td>
		<td></td>
    	<td align="right" ></td>
		<% if (IsItemListType) then %>
		<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
		<td align="right" ></td>
		<% end if %>
		<td align="right" ></td>
		<td align="right" ></td>
		<td align="right" ></td>
		<td align="right" ></td>
		<td align="right" ></td>
		<td align="right" ><%= FormatNumber(totStockSum,0) %></td>
		<% else %>
		<td align="right" ><%= FormatNumber(totStockSum,0) %></td>
		<% end if %>

		<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(totBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum14,0) %></td>
		<% else %>
		<td align="right" ><%= FormatNumber(totBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum5,0) %></td>
		<% end if %>

		<td align="right" ><%= FormatNumber(totBuySum6,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum,0) %></td>
		<td align="right" ></td>
    </tr>
</table>

<%

set ojaego = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
