<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǰ�˻�
' History : 2009.04.07 ������ ����
'			2012.08.29 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/offitemstock_cls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption ,BasicMonth ,restockdate, makerid
dim sum_ipgono,sum_reipgono,sum_sellno,sum_resellno ,i
dim sum_offchulgono, sum_offrechulgono, sum_etcchulgono, sum_etcrechulgono
dim sum_totsysstock, sum_availsysstock, sum_realstock
dim sum_errbaditemno, sum_errrealcheckno, sum_errcsno
dim mm_ipgono,mm_reipgono,mm_sellno,mm_resellno ,sysstock, sysavailstock, realstock, maystock ,ErrMsg, realstockWithBad
dim mm_offchulgono, mm_offrechulgono, mm_etcchulgono, mm_etcrechulgono ,mm_errbaditemno, mm_errrealcheckno, mm_errcsno
dim barcode, srcBarcode, sqlStr
dim useoff
	itemgubun   = request("itemgubun")
	itemid      = request("itemid")
	itemoption  = request("itemoption")
	useoff  	= request("useoff")
	srcBarcode	= request("barcode")

dim isShopreturnItem : isShopreturnItem = (itemgubun="90") and (itemid="1385")

itemid = Replace(itemid, ",", "")
if Not IsNumeric(itemid) then itemid=""

BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

'���ڵ� �˻�
srcBarcode = replace(replace(Trim(srcBarcode)," ",""),"-","")
if srcBarcode <> "" then
	if Len(srcBarcode) >= "12" then
	    sqlStr = "select top 1 b.itemgubun ,b.itemid ,b.itemoption"
	    sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b"
	    sqlStr = sqlStr + " where b.barcode='" & srcBarcode & "'"
	    rsget.Open sqlStr,dbget,1

	    if Not rsget.Eof then
	    	itemgubun = rsget("itemgubun")
	    	itemid = rsget("itemid")
	    	itemoption = rsget("itemoption")
	    end if

	    rsget.Close

	    if itemid = "" then
			sqlStr = "select top 1 i.itemgubun, i.shopitemid , i.itemoption"
			sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item i"
			sqlStr = sqlStr + " where i.extbarcode='" & srcBarcode & "'"
		    rsget.Open sqlStr,dbget,1

		    if Not rsget.Eof then
		    	itemgubun = rsget("itemgubun")
		    	itemid = rsget("shopitemid")
		    	itemoption = rsget("itemoption")
		    end if

		    rsget.Close

		    if itemid = "" then
	            IF (Len(srcBarcode)=12) and ((Left(srcBarcode,2)="10") or (Left(srcBarcode,2)="90") or (Left(srcBarcode,2)="70") or (Left(srcBarcode,2)="80") or (Left(srcBarcode,2)="85")) then
	                itemgubun = Left(srcBarcode,2)
	                itemid = CLng(Mid(srcBarcode,3,6))
	                itemoption = Right(srcBarcode,4)
	            end if

	            IF (Len(srcBarcode)=14) and ((Left(srcBarcode,2)="10") or (Left(srcBarcode,2)="90") or (Left(srcBarcode,2)="70") or (Left(srcBarcode,2)="80") or (Left(srcBarcode,2)="85")) then
	                itemgubun = Left(srcBarcode,2)
	                itemid = CLng(Mid(srcBarcode,3,8))
	                itemoption = Right(srcBarcode,4)
	            end if
		    end if
	    end if

	else
		response.write "<script>alert('���ڵ� ���̰� ª���ϴ�. 12�ڸ� �̻����� �Է��ϼ���.');history.go(-1);</script>"
		response.end	:	dbget.close()
	end if
end if

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"
if itemgubun<>"10" and (Not isShopreturnItem) then itemoption="0000"

dim oitem
if itemgubun = "10" then
	set oitem = new CItemInfo
		oitem.FRectItemID = itemid

		if itemid<>"" then
			oitem.GetOneItemInfo
			if oitem.FResultCount > 0 then
				makerid = oitem.foneitem.FMakerid
			end if
		end if
else
	set oitem = new CoffstockItemlist	'//�¶��� ��ũ������� Ŭ������ �浹, �������� ���� ����
		oitem.frectitemgubun = itemgubun
		oitem.FRectItemID = itemid
		oitem.frectitemoption = itemoption

		if itemid<>"" then
			oitem.GetoffItemDefaultData
			if oitem.FResultCount > 0 then
				makerid = oitem.foneitem.FMakerid
			end if
		end if
end if


dim oitemoption, oitemoptionOff
set oitemoption = new CItemOption
set oitemoptionOff = new CItemOption
	oitemoption.FRectItemID = itemid

	if itemid<>"" and itemgubun="10" then
		if (useoff = "Y") then
			oitemoption.FRectItemGubun = itemgubun
			oitemoption.GetItemOptionInfoByOffItemTable
		else
			oitemoption.GetItemOptionInfo
		end if
	end if

	if itemid<>"" and itemgubun="10" and (oitemoption.FResultCount<1) then
		oitemoptionOff.FRectItemGubun = itemgubun
		oitemoptionOff.FRectItemID = itemid
		oitemoptionOff.GetItemOptionInfoByOffItemTable
	end if

if (oitemoption.FResultCount<1) then
    if (Not isShopreturnItem) then
    	itemoption = "0000"
    end if
end if

if (oitem.FResultCount > 0) and (itemgubun = "10") and (itemgubun <> "") then
	barcode = itemgubun & CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) & itemoption
end if

dim offstock, offtotalstock
offtotalstock = 0
set offstock = new COffShopDailyStock
	offstock.FRectItemGubun = itemgubun
	offstock.FRectItemid = cStr(itemid)
	offstock.FRectItemoption = itemoption

	if itemid<>"" then
		if oitem.FResultCount>0 then
			offstock.FRectMakerid = oitem.FOneItem.FMakerid
		end if

		offstock.GetCurrentAllShopItemStockNEW

		for i=0 to offstock.FResultcount-1
			if not IsNULL(offstock.FItemList(i).Fcurrno) then
				offtotalstock = offtotalstock + offstock.FItemList(i).Fcurrno
			end if
		next
	end if

dim osummaryMonthstock
set osummaryMonthstock = new CSummaryItemStock
	osummaryMonthstock.FRectYYYYMM = BasicMonth
	osummaryMonthstock.FRectItemGubun = itemgubun
	osummaryMonthstock.FRectItemID =  itemid
	osummaryMonthstock.FRectItemOption =  itemoption

	if itemid<>"" then
		osummaryMonthstock.GetMonthly_Logisstock_Summary
	end if

dim osummarystock, isCurrStockExists
set osummarystock = new CSummaryItemStock
	osummarystock.FRectStartDate = BasicMonth + "-01"
	osummarystock.FRectItemGubun = itemgubun
	osummarystock.FRectItemID =  itemid
	osummarystock.FRectItemOption =  itemoption

	if itemid<>"" then
		osummarystock.GetCurrentItemStock
		isCurrStockExists= (osummarystock.FResultCount>0)
		osummarystock.GetDaily_Logisstock_Summary
	end if

dim osummaryagvstock, isCurrStockAgvExists
isCurrStockAgvExists = False
set osummaryagvstock = new CSummaryItemStock
	osummaryagvstock.FRectItemGubun = itemgubun
	osummaryagvstock.FRectItemID =  itemid
	osummaryagvstock.FRectItemOption =  itemoption

	if itemid<>"" then
		osummaryagvstock.GetCurrentAgvItemStock
		isCurrStockAgvExists = (osummaryagvstock.FResultCount>0)
	end if

dim oLastMonthstock
set oLastMonthstock = new CSummaryItemStock
	oLastMonthstock.FRectItemGubun = itemgubun
	oLastMonthstock.FRectItemID =  itemid
	oLastMonthstock.FRectItemOption =  itemoption

	if itemid<>"" then
	   oLastMonthstock.getLastMonthStock
	end if

if (itemid = "") then
elseif (oitem.FResultCount < 1) then
elseif (oitemoption.FResultCount>0) and (itemoption="0000") then
elseif (oitemoption.FResultCount<1) and (itemoption<>"0000") then
else
    '�԰�����
    if ((oitem.FOneItem.Fdanjongyn="S") and (itemoption="0000")) then
    	restockdate = oitem.GetReStockDate
    end if
end if

''''���Էµ� ����ΰ�� - ��� ���� Key
''dim IsInvalidOption
''if (oitemoption.FResultCount>0)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function barcodesearch(){
	frm.itemid.value="";
	frm.submit();
}

function frmsearch(){
	frm.barcode.value="";
	frm.submit();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'simpleitemedit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function RefreshIpchulStock(){
	if (confirm('����� ���� ��ü ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="ipchulallrefreshbyitemid";
		frmrefresh.submit();
	}
}

function RefreshOldTotalSellStock(){
	if (confirm('���� ���� ��ü ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemsellrefresholdall";
		frmrefresh.submit();
	}
}


function RefreshRecentStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('�ֱ� ������Ʈ �� ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemrecentipchulrefresh";
		frmrefresh.submit();
	}
}

function RefreshRecentStockV2(){
	if (confirm('�ֱ� ������Ʈ �� ������ ���ΰ�ħ �Ͻðڽ��ϱ�? V2')){
		frmrefresh.mode.value="itemrecentipchulrefreshv2";
		frmrefresh.submit();
	}
}

function RefreshAgvStock() {
    var url;
    var brandArray;
    var skuCdArray;
    var barcode = '<%= barcode %>';

    <% if (oitemoption.FResultCount > 0) and (itemoption = "0000") then %>
    alert('���� �ɼ��� �����ϼ���(' + barcode + ')');
    return;
    <% end if %>

    if ((barcode.length != 12) && (barcode.length != 14)) {
        alert('�߸��� ���ڵ��Դϴ�.(' + barcode + ')');
        return;
    }

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'https://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=currstockList&skuCdArray=' + barcode;

    if (confirm('AGV���(��ǰ) ���ΰ�ħ �Ͻðڽ��ϱ�?') != true) { return; }

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                alert('������Ʈ�Ǿ����ϴ�.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
}

function RefreshAgvStockByBrand(brand) {
    var url;
    var brandArray;
    var skuCdArray;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'https://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=currstockList&brandArray=' + brand;

    if (confirm('AGV���(�귣��) ���ΰ�ħ �Ͻðڽ��ϱ�?') != true) { return; }

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                alert('������Ʈ�Ǿ����ϴ�.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
}

function refreshAccStock(comp,yyyymm){
	var frm =document.frmrefresh;
	frm.mode.value = "itemAccStock";
	frm.yyyymm.value = yyyymm;


	if (confirm(yyyymm+'�� �⸻��� ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		comp.disabled=true;
		frm.submit();
	}
}

function RefreshTodayStock(itemgubun,itemid,itemoption){
	if (confirm('���� ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemtodayipchulrefresh";
		frmrefresh.submit();
	}
}


function RefreshALLStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('��ü ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemallipchulrefresh";
		frmrefresh.submit();
	}
}

function PopStockBaditem(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popbaditemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popRealErrList(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=900,height=460,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popBuyItemList(itemstate){
    var popwin = window.open('/admin/ordermaster/oneitembuylist.asp?itemgubun=<%= itemgubun %>&itemid=<%= itemid %>&itemoption=<%= itemoption %>&itemstate=' + itemstate + '&menupos=77','popBuyItemList','width=980,height=460,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popOffItemList(itemstate) {
    var popwin = window.open('/admin/fran/jumunlist.asp?menupos=520&statecd=' + itemstate + '&designer=<%= makerid %>&itemgubun=<%= itemgubun %>&itemid=<%= itemid %>&itemoption=<%= itemoption %>','popOffItemList','width=1180,height=460,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popBuyItemListChulgo(ostr){
    if (ostr.length==7){
        var yyyy1   =   ostr.substr(0,4);
        var mm1     =   ostr.substr(5,2);
        var dd1     =   '01';

        var lastdate = new Date(yyyy1,mm1*1+1,0);
        var lastdate2 = new Date(yyyy1,mm1,0);

        var yyyy2   =   lastdate.getFullYear().toString(); //lastdate.getYear().toString();
        var mm2     =   lastdate.getMonth().toString();
        var dd2     =   lastdate2.getDate().toString();

        if (mm2.length<2) { mm2 = '0' + mm2 };
        if (dd2.length<2) { dd2 = '0' + dd2 };

    }else{
        var yyyy1   =   ostr.substr(0,4);
        var mm1     =   ostr.substr(5,2);
        var dd1     =   ostr.substr(8,2);

        var yyyy2   =   yyyy1;
        var mm2     =   mm1;
        var dd2     =   dd1;
    }

    var rectStr = '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2;

	var popwin;
	<% if (itemgubun = "85") then %>
		popwin = window.open('/admin/ordermaster/onegiftitembuylist.asp?itemgubun=<%= itemgubun %>&itemid=<%= itemid %>&itemoption=<%= itemoption %>&itemstate=8&menupos=1527&datetype=beasong' + rectStr ,'popBuyItemList','width=980,height=460,scrollbars=yes,resizable=yes')
	<% else %>
		popwin = window.open('/admin/ordermaster/oneitembuylist.asp?itemgubun=<%= itemgubun %>&itemid=<%= itemid %>&itemoption=<%= itemoption %>&itemstate=8&menupos=77&datetype=beasong' + rectStr ,'popBuyItemList','width=980,height=460,scrollbars=yes,resizable=yes')
	<% end if %>
	popwin.focus();
}

function popCSItemListChulgo(ostr){
    if (ostr.length==7){
        var yyyy1   =   ostr.substr(0,4);
        var mm1     =   ostr.substr(5,2);
        var dd1     =   '01';

        var lastdate = new Date(yyyy1,mm1*1+1,0);
        var lastdate2 = new Date(yyyy1,mm1,0);

        var yyyy2   =   lastdate.getFullYear().toString(); //lastdate.getYear().toString();
        var mm2     =   lastdate.getMonth().toString();
        var dd2     =   lastdate2.getDate().toString();

        if (mm2.length<2) { mm2 = '0' + mm2 };
        if (dd2.length<2) { dd2 = '0' + dd2 };

    }else{
        var yyyy1   =   ostr.substr(0,4);
        var mm1     =   ostr.substr(5,2);
        var dd1     =   ostr.substr(8,2);

        var yyyy2   =   yyyy1;
        var mm2     =   mm1;
        var dd2     =   dd1;
    }

    var rectStr = '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2;

//alert(rectStr);
//return;
    var popwin = window.open('/cscenter/action/oneitemcslist.asp?itemgubun=<%= itemgubun %>&itemid=<%= itemid %>&itemoption=<%= itemoption %>&currstate=finish&menupos=1457&datetype=finish' + rectStr ,'popCSItemListChulgo','width=980,height=460,scrollbars=yes,resizable=yes')
	popwin.focus();
}

//����
function pop_itemedit_off_edit(ibarcode){

	var pop_itemedit_off_edit = window.open('/common/offshop/item/pop_itemedit_off_edit.asp?barcode=' + ibarcode,'pop_itemedit_off_edit','width=1024,height=768,resizable=yes,scrollbars=yes');
	pop_itemedit_off_edit.focus();
}

//���ڵ����
function barcodeManage(itemcode)
{
	var popbarcodemanage = window.open('/admin/stock/popBarcodeManage.asp?itemcode=' + itemcode,'popbarcodemanage','width=550,height=400,resizable=yes,scrollbars=yes');
	popbarcodemanage.focus();
}

//���ڵ����
function upcheManageCode(itemcode)
{
	var popupcheManageCode = window.open('/admin/stock/popUpcheManageCode.asp?itemcode=' + itemcode,'popupcheManageCode','width=550,height=400,resizable=yes,scrollbars=yes');
	popupcheManageCode.focus();
}

function jsPopIpgoList(itembarcode) {
    var pop = window.open('/admin/newstorage/orderlist.asp?menupos=537&barcode=' + itembarcode,'jsPopIpgoList','width=1600,height=500,resizable=yes,scrollbars=yes');
    pop.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td align="left">
					��ǰ�ڵ�:
					<% drawSelectBoxItemGubun "itemgubun", itemgubun %>

		        	<input type="text" class="text" name=itemid value="<%= itemid %>" size=8 maxlength=8  onKeyPress="if (event.keyCode == 13){ frmsearch(); return false;}">

		        	<input type="text" class="text_ro" value="<%= itemoption %>" size=4 maxlength=4 readonly>

					<% if oitemoption.FResultCount>0 then %>

					<select class="select" name="itemoption">
					<option value="0000">----
					<% for i=0 to oitemoption.FResultCount-1 %>
					<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> >[<%= oitemoption.FITemList(i).FItemOption %>]<%= oitemoption.FITemList(i).FOptionName %>
					<% next %>
					</select>
					<% end if %>

		        	<input type="button" class="button" value="�˻�" onclick="frmsearch();">


					<% if (oitemoption.FResultCount<1) and (oitemoptionOff.FResultCount>0) then %>
					<font color=red>���� : ��-���� �ɼǴٸ�</font>
					<% end if %>
					<input type="checkbox" name="useoff" value="Y" <%= CHKIIF(useoff="Y", "checked", "") %>> OFF��ǰ���� ���
		        </td>
			    <td align="right">
					���ڵ�: <input type="text" class="text" name="barcode" value="<%= srcBarcode %>" size="16" maxlength="16" onKeyPress="if (event.keyCode == 13){barcodesearch(); return false;}">
    				<input type="button" class="button" value="�˻�" onclick="barcodesearch();">
				</td>
			</tr>
			<% if oitem.FResultCount>0 or (isCurrStockExists) then %>
			<tr bgcolor="#FFFFFF">
		        <td colspan="2" align="right">
		            <% if itemid<>"" then %>
			        		����������Ʈ : <b><%= osummarystock.FOneItem.Flastupdate %></b>
				    <% end if %>
		        	<% if (C_ADMIN_AUTH=true) or (session("ssBctId")="josin222") then %>
						&nbsp;
			            <input type="button" class="button" value="����� ��ü ���ΰ�ħ" onclick="RefreshIpchulStock();">
			        <% end if %>
					&nbsp;
			        <input type="button" class="button" value="���ΰ�ħ" onclick="RefreshRecentStock();">
					&nbsp;
					<input type="button" class="button" value="���ΰ�ħ(V2)" onclick="RefreshRecentStockV2();">
                    &nbsp;
					<input type="button" class="button" value="AGV ���ΰ�ħ" onclick="RefreshAgvStock();">
		    	</td>
		    </tr>
			<% end if %>
		</table>
	</td>
</tr>
</form>
</table>

<p>

<% if (oitem.FResultCount>0) or (isCurrStockExists) then %>

<% if itemgubun="10" then %>
	<% if (oitem.FResultCount>0) then %>
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
		  	<td width="60">��ǰ�ڵ�</td>
		  	<td width="300">
		  		<%= itemgubun %> <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
		  		&nbsp;
		  		<% if itemgubun="10" then %>
		  		<input type="button" class="button" value="����" onclick="PopItemSellEdit('<%= itemid %>');">
		  		<% end if %>
		  	</td>
		  	<td colspan="5">��չ�ۼҿ��� :
			<% if (oitem.FOneItem.FavgDLvDate>-1) then %>
			    <a href="javascript:popItemAvgDlvList('<%= itemid %>');">D+<%= oitem.FOneItem.FavgDLvDate+1 %></a>
			<% else %>
			    <a href="javascript:popItemAvgDlvList('<%= itemid %>');">������ ����</a>
			<% end if %>
			</td>

		</tr>
		<tr bgcolor="#FFFFFF">
		  	<td>�귣��ID</td>
		  	<td><a href="javascript:RefreshAgvStockByBrand('<%= oitem.FOneItem.FMakerid %>')"><%= oitem.FOneItem.FMakerid %></a></td>
		  	<td>�Ǹſ���</td>
		  	<td colspan=4><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  	<td>��ǰ��</td>
		  	<td><%= oitem.FOneItem.FItemName %></td>
		  	<td>��뿩��</td>
		  	<td colspan=4><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  	<td>�ǸŰ�</td>
		  	<td>
		  		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
		  		&nbsp;&nbsp;
		  		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
		  	    <% if oitem.FOneItem.FSellcash<>0 then %>
				<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
				<% end if %>
				&nbsp;&nbsp;
				<!-- ���ο���/�������뿩�� -->
				<% if (oitem.FOneItem.FSailYn="Y") then %>
				    <font color=red>
				    <% if (oitem.FOneItem.Forgprice<>0) then %>
				        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
				    <% end if %>
				     ����
				    </font>
				<% end if %>

				<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

				    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> ����
				    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
				<% end if %>

		  	</td>
		  	<td>��������</td>
		  	<td colspan="2">
		  		<%= fncolor(oitem.FOneItem.Fdanjongyn,"dj") %>
		  		<% if oitem.FOneItem.Fdanjongyn="N" then %>
				������
				<% end if %>
			</td>
			<td align="center"><input type="button" class="button" value="���ڵ����" onClick="barcodeManage('<%= BF_MakeTenBarcode(itemgubun, itemid, itemoption) %>');"></td>
			<td align="center"><input type="button" class="button" value="��ü�ڵ����" onClick="upcheManageCode('<%= BF_MakeTenBarcode(itemgubun, itemid, itemoption) %>');"></td>
		</tr>

		<% if oitemoption.FResultCount>1 then %>
		    <!-- �ɼ��� �ִ°�� -->
		    <% for i=0 to oitemoption.FResultCount -1 %>
			    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
			    <tr bgcolor="#FFFFFF">
			      	<td><font color="#AAAAAA">�ɼǸ� :</font></td>
			      	<td><font color="#AAAAAA"><%
			      		Response.Write "[" & oitemoption.FITemList(i).Fitemoption & "]" & oitemoption.FITemList(i).FOptionName & "&nbsp;"
			      		Response.Write CHKIIF(oitemoption.FITemList(i).Foptaddprice <> "0","(+"&FormatNumber(oitemoption.FITemList(i).Foptaddprice,0)&")","")
			      	%></font></td>
			      	<td><font color="#AAAAAA">�������� : </font></td>
			      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
			      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
					<td align="center"><%= oitemoption.FITemList(i).Fbarcode %></td>
					<td align="center"><%= oitemoption.FITemList(i).Fupchemanagecode %></td>
			    </tr>
			    <% else %>

			    <% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
			    <tr bgcolor="#EEEEEE">
			    <% else %>
			    <tr bgcolor="#FFFFFF">
			    <% end if %>
			      	<td>�ɼǸ�</td>
			      	<td><%
			      		Response.Write "[" & oitemoption.FITemList(i).Fitemoption & "]" & oitemoption.FITemList(i).FOptionName & "&nbsp;"
			      		Response.Write CHKIIF(oitemoption.FITemList(i).Foptaddprice <> "0","(+"&FormatNumber(oitemoption.FITemList(i).Foptaddprice,0)&")","")
			      	%></td>
			      	<td>��������</td>
			      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
			      	<td>
			      	  ���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)
				      <% if (oitem.FOneItem.Fdanjongyn = "S") then %>
				      (���԰����� : <%= oitemoption.FITemList(i).Frestockdate %>)
				      <% end if %>
			      	</td>
					<td align="center"><%= oitemoption.FITemList(i).Fbarcode %></td>
					<td align="center"><%= oitemoption.FITemList(i).Fupchemanagecode %></td>
			    </tr>
			    <% end if %>
		    <% next %>
		<% else %>
			<tr bgcolor="#FFFFFF">
		      	<td>�ɼǸ�</td>
		      	<td>-</td>
		      	<td>��������</td>
		      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
		      	<td>
		      		���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)
				<% if ((oitem.FOneItem.Fdanjongyn="S") and (oitemoption.FResultCount<1)) then %>
				(���԰����� : <%= restockdate %>)
				<% end if %>
		      	</td>
				<td align="center"><%= oitem.FOneItem.Fbarcode %></td>
				<td align="center"><%= oitem.FOneItem.Fupchemanagecode %></td>
		    </tr>
		<% end if %>
		</table>
	<% end if %>
<%
'//�¶��� ���� ������
else
%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
		<td rowspan=<%= 5 + oitem.FResultCount -1 %> width="110" valign="top" align="center">
			<img src="<%= oitem.foneitem.FImageList %>" width="100" height="100">
		</td>
	  	<td width="60"><b>*��ǰ����</b></td>
	  	<td width="300">
	  		<!--<input type="button" value="����" onclick="pop_itemedit_off_edit('<%'= oitem.foneitem.Fitemgubun %><%'=  Format00(6,oitem.foneitem.Fitemid) %><%'= oitem.foneitem.Fitemoption %>');" class="button">-->
	  	</td>
	  	<td width="60">�귣��ID :</td>
	  	<td colspan=2><%= oitem.foneitem.FMakerid %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  	<td>��ǰ�ڵ� :</td>
	  	<td><%= oitem.foneitem.fitemgubun %> <b><%= CHKIIF(oitem.foneitem.FItemID>=1000000,Format00(8,oitem.foneitem.FItemID),Format00(6,oitem.foneitem.FItemID)) %></b> <%= oitem.foneitem.fitemoption %></td>
	  	<td>��뿩�� : </td>
	  	<td colspan=2><%= oitem.foneitem.FIsUsing %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>�ǸŰ� :</td>
		<td>
			<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
		</td>
	  	<td>��ǰ�� :</td>
	  	<td><%= oitem.foneitem.FItemName %></td>
	</tr>
    <tr bgcolor="#FFFFFF">
      	<td><font color="#AAAAAA">�ɼǸ� :</font></td>
      	<td><font color="#AAAAAA"><%= oitem.foneitem.FItemOptionName %></font></td>
      	<td><font color="#AAAAAA">������� : </font></td>
      	<td>
      		<%= oitem.foneitem.GetCheckStockNo %> : (NEW)
      	</td>
    </tr>
	</table>
<% end if %>

<!-- ǥ �߰��� ����
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	<br>�ý��� ����� = �԰�/��ǰ�� + ON�Ǹ�/��ǰ�� + OFF���/��ǰ�� + ��Ÿ���/��ǰ��
	        <br>�ý��� ��ȿ��� = �ý��� ����� - �Ѻҷ�(+)
	        <br>�ǻ� ��� = �ý��� ��ȿ��� - �ѽǻ����(+)
	        <br>����ľ� ��� = �ǻ� ��� + ON��ǰ�غ� + OFF��ǰ�غ�
			<br>������� = ����ľ� ��� + ON�����Ϸ� + ON�ֹ����� + OFF�ֹ�����
			<br>
        </td>
        <td valign="top">
        	<br><font color="blue">ON 7�ϰ��Ǹ� = ON 7�ϰ� ��� + ON��ǰ�غ� + ON�����Ϸ� + ON�ֹ�����</font>
        	<br><font color="blue">OFF 7�ϰ��Ǹ� = OFF 7�ϰ� �Ǹ�(SHOP�Ǹŷ�)</font>
			<br><font color="blue">���� ����� = ����ľ� ��� + ON�����Ϸ� + ON�ֹ�����</font>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<p>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			<b>*�������</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">��<br>�԰�/��ǰ</td>
    	<td width="50">ON��<br>�Ǹ�/��ǰ</td>
		<td width="50">OFF��<br>���/��ǰ</td>
		<td width="50">��Ÿ<br>���/��ǰ</td>
		<td width="50">CS<br>���/��ǰ</td>
		<td width="50" bgcolor="F4F4F4">�ý���<br>�����</td>
		<td width="50">�ѽǻ�<br>����</td>
		<td width="50" bgcolor="F4F4F4">�ǻ�<br>���</td>
		<td width="50">�Ѻҷ�</td>
      	<td width="50" bgcolor="F4F4F4">�ǻ�<br>��ȿ���</td>
		<td width="50">ON<br>��ǰ�غ�</td>
		<td width="50">OFF<br>��ǰ�غ�</td>
		<td width="50" bgcolor="F4F4F4">����ľ�<br>���</td>
		<td width="50">ON<br>�����Ϸ�</td>
		<td width="50">ON<br>�ֹ�����</td>
		<td width="50">OFF<br>�ֹ�����</td>
		<td bgcolor="F4F4F4">����<br>���</td>
		<td bgcolor="F4F4F4">��<br>���</td>
        <td bgcolor="F4F4F4">���<br />����</td>
        <td bgcolor="F4F4F4">AGV<br />���</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= osummarystock.FOneItem.Ftotipgono %></td>
    	<td><%= -1*osummarystock.FOneItem.Ftotsellno %></td>
    	<td><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
    	<td><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
    	<td><%= osummarystock.FOneItem.Ferrcsno %></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.Ftotsysstock %></b></td>
    	<td><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
    	<td><%= osummarystock.FOneItem.getErrAssignStock %></td>
    	<td><%= osummarystock.FOneItem.Ferrbaditemno %></td>
    	<!--td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.Favailsysstock %></b></td-->
    	<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.Frealstock %></td>
    	<td><a href="javascript:popBuyItemList('5');"><%= osummarystock.FOneItem.Fipkumdiv5 %></a></td>
    	<td><a href="javascript:popOffItemList('1')"><%= osummarystock.FOneItem.Foffconfirmno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.GetCheckStockNo %></b></td>
    	<td><a href="javascript:popBuyItemList('4');"><%= osummarystock.FOneItem.Fipkumdiv4 %></a></td>
    	<td><a href="javascript:popBuyItemList('2');"><%= osummarystock.FOneItem.Fipkumdiv2 %></a></td>
    	<td><a href="javascript:popOffItemList('before1')"><%= osummarystock.FOneItem.Foffjupno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.GetMaystock %></b></td>
		<td bgcolor="F4F4F4"><b><%= offtotalstock %></b></td>
        <td bgcolor="F4F4F4"><%= CHKIIF(isCurrStockAgvExists, osummaryagvstock.FOneItem.FwarehouseCd, "BLK") %></td>
        <td bgcolor="F4F4F4"><%= CHKIIF(isCurrStockAgvExists, osummaryagvstock.FOneItem.Fagvstock, 0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="10" align="right"><input type="button" class="button" value="�ǻ�����Է�" onclick="popRealErrInput('<%= itemgubun %>','<%= itemid %>','<%= itemoption %>');"></td>
    	<td colspan="2"><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
    	<td></td>
    	<td colspan="3"><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>
    	<td colspan="4"></td>
    </tr>
</table>

<!--
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">

        	<br><b>Fsell7days</b> --&gt; ����7��(��������)����ֹ��Ǽ� &nbsp;&nbsp;&nbsp;  45(50) --&gt; �ֹ��Ǽ�(�ֹ���ǰ��)
        	<br><b>Foffchulgo7days</b> --&gt; ����7�� ��������(��üƯ������) �ǸŰǼ� &nbsp;&nbsp;&nbsp;  8(10) --&gt; �ֹ��Ǽ�(�ֹ���ǰ��)
	        <br><b>Frequireno</b> --&gt; (Fsell7days + Foffchulgo7days) / Fmaxsellday * C_STOCK_DAY
	        <br><b>Fshortageno</b> --&gt;  Frealstock - Frequireno - Fipkumdiv5 - Foffconfirmno - Fipkumdiv4 - Fipkumdiv2 - Foffjupno
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
-->

<p>

* P = �����Ǹż��� = (A+B)/<%= osummarystock.FOneItem.Fmaxsellday %>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>*�������</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">ON<br><%= osummarystock.FOneItem.Fmaxsellday %>�ϰ� ���<br>(A)</td>
		<td width="80">OFF<br><%= osummarystock.FOneItem.Fmaxsellday %>�ϰ� �Ǹ�<br>(B)</td>
		<td width="80">��<br><%= osummarystock.FOneItem.Fmaxsellday %>�ϰ� �Ǹ�<br>(A+B)</td>
		<td width="80" bgcolor="F4F4F4">
			(<%= osummarystock.FOneItem.FDayForSafeStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)�ϰ�<br>�ʿ����<br>
			C=(<%= osummarystock.FOneItem.FDayForSafeStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)*P
		<td width="80" bgcolor="F4F4F4">
			(<%= osummarystock.FOneItem.FDayForMaxStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)�ϰ�<br>�ʿ����<br>
			D=(<%= osummarystock.FOneItem.FDayForMaxStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)*P
		</td>
		<td width="80" bgcolor="F4F4F4">�������<br>�ʿ����<br>(R)</td>
		<td width="80">�ǻ�<br>��ȿ���<br>(S)</td>
		<td width="100" bgcolor="F4F4F4">(<%= osummarystock.FOneItem.FDayForSafeStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)����<br>�ʰ�(����)����<br>(S-C-R)</td>
		<td width="100" bgcolor="F4F4F4">(<%= osummarystock.FOneItem.FDayForMaxStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)����<br>�ʰ�(����)����<br>(S-D-R)</td>
		<td>����ּ���</td>
        <td>��ǰ���</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25" align=center>
		<td><%= osummarystock.FOneItem.Fsell7days*-1 %><br><!-- (--) --></td>
		<td><%= osummarystock.FOneItem.Foffchulgo7days*-1 %><br><!-- (--) --></td>
		<td><b><%= osummarystock.FOneItem.Fsell7days*-1 + osummarystock.FOneItem.Foffchulgo7days*-1 %><br><!-- (--) --></b></td>
		<td><%= osummarystock.FOneItem.Frequireno*-1 %></td>
		<td><%= osummarystock.FOneItem.FrequireMaxno*-1 %></td>

		<td><%= (osummarystock.FOneItem.GetReqNotChulgoNo)*-1 %></td>
		<td><b><%= osummarystock.FOneItem.Frealstock %></b></td>
		<td><b><%= osummarystock.FOneItem.Fshortageno %></b>
		    <% if osummarystock.FOneItem.Fshortageno<>osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+osummarystock.FOneItem.GetReqNotChulgoNo then %>
		    <br><font color="red"><%= osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+osummarystock.FOneItem.GetReqNotChulgoNo %></font>
		    <% end if %>
		</td>
		<td>
		    <b><%= osummarystock.FOneItem.GetShortageMaxNo %></b>
		    <% if osummarystock.FOneItem.GetShortageMaxNo <> osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+(osummarystock.FOneItem.FrequireMaxno - osummarystock.FOneItem.Frequireno)+osummarystock.FOneItem.GetReqNotChulgoNo then %>
		    <br><font color="red"><%= osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+(osummarystock.FOneItem.FrequireMaxno - osummarystock.FOneItem.Frequireno)+osummarystock.FOneItem.GetReqNotChulgoNo %></font>
		    <% end if %>
		</td>
		<td>
            <a href="javascript:jsPopIpgoList('<%= BF_MakeTenBarcode(itemgubun, itemid, itemoption) %>')">
			<% if osummarystock.FOneItem.Fpreorderno<>osummarystock.FOneItem.Fpreordernofix then %>
			<%= osummarystock.FOneItem.Fpreorderno %>-&gt;<%= osummarystock.FOneItem.Fpreordernofix %>
			<% else %>
			<%= osummarystock.FOneItem.Fpreordernofix %>
			<% end if %>
            </a>
		</td>
		<td>
			<%= osummarystock.FOneItem.Fitemgrade %>
		</td>
	</tr>
</table>

	<% if (C_ADMIN_AUTH=true) or C_ADMIN_AUTH then %>
	<p>

	* P = �����Ǹż��� = (A+B)/<%= osummarystock.FOneItem.Fmaxsellday %>

	<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15">
				<b>*�������<font color=red>NEW</font></b>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="80">ON<br><%= osummarystock.FOneItem.Fmaxsellday %>�ϰ� ���<br>(A)</td>
			<td width="80">OFF<br><%= osummarystock.FOneItem.Fmaxsellday %>�ϰ� �Ǹ�<br>(B)</td>
			<td width="80">��<br><%= osummarystock.FOneItem.Fmaxsellday %>�ϰ� �Ǹ�<br>(A+B)</td>
			<td width="80" bgcolor="F4F4F4">
				(<%= osummarystock.FOneItem.FDayForSafeStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)�ϰ�<br>�ʿ����<br>
				C=(<%= osummarystock.FOneItem.FDayForSafeStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)*P
			</td>
			<td width="80" bgcolor="F4F4F4">
				(<%= osummarystock.FOneItem.FDayForMaxStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)�ϰ�<br>�ʿ����<br>
				D=(<%= osummarystock.FOneItem.FDayForMaxStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)*P
			</td>
			<td width="80" bgcolor="F4F4F4">�������<br>�ʿ����<br>(R)</td>
			<td width="80">�ǻ�<br>��ȿ���<br>(S)</td>
			<td width="100" bgcolor="F4F4F4">(<%= osummarystock.FOneItem.FDayForSafeStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)����<br>�ʰ�(����)����<br>(S-C-R)</td>
			<td width="100" bgcolor="F4F4F4">(<%= osummarystock.FOneItem.FDayForMaxStock %>+<%= osummarystock.FOneItem.FDayForLeadTime %>)����<br>�ʰ�(����)����<br>(S-D-R)</td>
			<td>����ּ���</td>
            <td>��ǰ���</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="25" align=center>
			<td><%= osummarystock.FOneItem.Fsell7days*-1 %><br><!-- (--) --></td><!-- 7���� �ƴҼ� �ִ�. -->
			<td><%= osummarystock.FOneItem.Foffchulgo7days*-1 %><br><!-- (--) --></td>
			<td><b><%= osummarystock.FOneItem.Fsell7days*-1 + osummarystock.FOneItem.Foffchulgo7days*-1 %><br><!-- (--) --></b></td>
			<td><%= osummarystock.FOneItem.Frequireno*-1 %></td>
			<td><%= osummarystock.FOneItem.FrequireMaxno*-1 %></td>

			<td><%= (osummarystock.FOneItem.GetReqNotChulgoNo)*-1 %></td>
			<td><b><%= osummarystock.FOneItem.Frealstock %></b></td>
			<td><b><%= osummarystock.FOneItem.Fshortageno %></b>
			    <% if osummarystock.FOneItem.Fshortageno<>osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+osummarystock.FOneItem.GetReqNotChulgoNo then %>
			    <br><font color="red"><%= osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+osummarystock.FOneItem.GetReqNotChulgoNo %></font>
			    <% end if %>
			</td>
			<td>
			    <b><%= osummarystock.FOneItem.GetShortageMaxNo %></b>
			    <% if osummarystock.FOneItem.GetShortageMaxNo <> osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+(osummarystock.FOneItem.FrequireMaxno - osummarystock.FOneItem.Frequireno)+osummarystock.FOneItem.GetReqNotChulgoNo then %>
			    <br><font color="red"><%= osummarystock.FOneItem.Frealstock+osummarystock.FOneItem.Frequireno+(osummarystock.FOneItem.FrequireMaxno - osummarystock.FOneItem.Frequireno)+osummarystock.FOneItem.GetReqNotChulgoNo %></font>
			    <% end if %>
			</td>
			<td>
				<% if osummarystock.FOneItem.Fpreorderno<>osummarystock.FOneItem.Fpreordernofix then %>
				<%= osummarystock.FOneItem.Fpreorderno %>-&gt;<%= osummarystock.FOneItem.Fpreordernofix %>
				<% else %>
				<%= osummarystock.FOneItem.Fpreordernofix %>
				<% end if %>
			</td>
		    <td>
			    <%= osummarystock.FOneItem.Fitemgrade %>
		    </td>
		</tr>
	</table>
	<% end if %>

<% end if %>

<p>

<% if (oitem.FResultCount>0) or (itemgubun<>"10" and osummaryMonthstock.FResultCount>0)  then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>*�Ϻ� ���⳻��</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80">�Ͻ�</td>
      	<td width="55">�԰�</td>
      	<td width="55">��ǰ</td>
      	<td width="55">ON<br>���</td>
      	<td width="55">ON<br>��ǰ</td>
      	<td width="55">OFF<br>���</td>
      	<td width="55">OFF<br>��ǰ</td>

      	<td width="55">��Ÿ<br>���/��ǰ</td>
      	<td width="55">CS<br>���/��ǰ</td>
        <td width="60">�ý���<br>�����</td>
        <td width="55">(�ǻ�)<br>����</td>
        <td width="60">�ǻ�<br>���</td>
      	<td width="55">�ҷ�</td>
      	<!-- td width="60">�ý���<br>��ȿ���</td -->
      	<td width="60">�ǻ�<br>��ȿ���</td>
      	<td>���</td>
    </tr>
<!-- �����α� -->
<% if osummaryMonthstock.FResultCount>0 then %>
<% for i=0 to osummaryMonthstock.FResultCount-1 %>
<%
sum_ipgono = sum_ipgono + osummaryMonthstock.FItemList(i).Fipgono
sum_reipgono = sum_reipgono + osummaryMonthstock.FItemList(i).Freipgono
sum_sellno = sum_sellno + osummaryMonthstock.FItemList(i).Fsellno
sum_resellno = sum_resellno + osummaryMonthstock.FItemList(i).Fresellno
sum_offchulgono = sum_offchulgono + osummaryMonthstock.FItemList(i).Foffchulgono
sum_offrechulgono = sum_offrechulgono + osummaryMonthstock.FItemList(i).Foffrechulgono
sum_etcchulgono = sum_etcchulgono + osummaryMonthstock.FItemList(i).Fetcchulgono
sum_etcrechulgono = sum_etcrechulgono + osummaryMonthstock.FItemList(i).Fetcrechulgono
sum_errbaditemno	= sum_errbaditemno + osummaryMonthstock.FItemList(i).Ferrbaditemno
sum_errrealcheckno	= sum_errrealcheckno + osummaryMonthstock.FItemList(i).Ferrrealcheckno
sum_errcsno         = sum_errcsno + osummaryMonthstock.FItemList(i).Ferrcsno

sum_totsysstock = sum_totsysstock + osummaryMonthstock.FItemList(i).Ftotsysstock
sum_availsysstock = sum_availsysstock + osummaryMonthstock.FItemList(i).Favailsysstock
sum_realstock = sum_realstock + osummaryMonthstock.FItemList(i).Frealstock


sysstock = sysstock + osummaryMonthstock.FItemList(i).Ftotsysstock
sysavailstock = sysavailstock + osummaryMonthstock.FItemList(i).Favailsysstock
realstock = realstock + osummaryMonthstock.FItemList(i).Frealstock
maystock = maystock + osummaryMonthstock.FItemList(i).Frealstock

realstockWithBad = sysstock+sum_errrealcheckno ''2013/11/22�߰�

'sum_offsell = sum_offsell + osummaryMonthstock.FItemList(i).Foffsellno
'offstockno = offstockno + osummaryMonthstock.FItemList(i).Foffchulgono*-1 + osummaryMonthstock.FItemList(i).Foffrechulgono*-1 - osummaryMonthstock.FItemList(i).Foffsellno

''rw DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0)
%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummaryMonthstock.FItemList(i).Fyyyymm %></td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>','I');"><%= osummaryMonthstock.FItemList(i).Fipgono %></a></td>
      	<td><%= osummaryMonthstock.FItemList(i).Freipgono %></td>
      	<td><a href="javascript:popBuyItemListChulgo('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>')"><%= osummaryMonthstock.FItemList(i).Fsellno %></a></td>
      	<td><%= osummaryMonthstock.FItemList(i).Fresellno %></td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>','S');"><%= osummaryMonthstock.FItemList(i).Foffchulgono %></a></td>
      	<td><%= osummaryMonthstock.FItemList(i).Foffrechulgono %></td>

      	<td><a href="javascript:PopItemIpChulList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>','E');"><%= osummaryMonthstock.FItemList(i).Fetcchulgono + osummaryMonthstock.FItemList(i).Fetcrechulgono %></a></td>
    	<td><a href="javascript:popCSItemListChulgo('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>')"><%= osummaryMonthstock.FItemList(i).Ferrcsno %></a></td>
        <td><%= sysstock %></td>
        <td><a href="javascript:popRealErrList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>')"><%= osummaryMonthstock.FItemList(i).Ferrrealcheckno %></a></td>
      	<td><%= realstockWithBad %></td>
      	<td><a href="javascript:PopStockBaditem('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>')"><%= osummaryMonthstock.FItemList(i).Ferrbaditemno %></a></td>
      	<!-- td><%= sysavailstock %></td -->
      	<td><%= realstock %></td>
      	<td>
      	    <% if realstock<>0 then %>
      	    <%= CLng((osummaryMonthstock.FItemList(i).Fsellno + osummaryMonthstock.FItemList(i).Foffchulgono)*-1/realstock*100)/100 %>
      	    <% end if %>

      	    <% if Not isNULL(osummaryMonthstock.FItemList(i).Flastmwdiv) then %>
      	    [<%= osummaryMonthstock.FItemList(i).Flastmwdiv %> / <%= osummaryMonthstock.FItemList(i).Flasttotsysstock %>]
      	    <% end if %>

			<% if (osummaryMonthstock.FItemList(i).Fyyyymm+"-01">=LEFT(dateadd("m",-1,LEFT(now(),7)+"-01"),10)) then %>
			<input type="button" value="�⸻���ۼ�" onClick="refreshAccStock(this,'<%=osummaryMonthstock.FItemList(i).Fyyyymm%>')">
			<% end if %>
      	</td>
    </tr>
<% next %>
	<tr align="center" bgcolor="#EEEEEE">
		<td>�����Ұ�</td>
		<td>
		    <%= sum_ipgono %>
		    <% if oLastMonthstock.FOneItem.Fipgono<>sum_ipgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fipgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_reipgono %>
		    <% if oLastMonthstock.FOneItem.Freipgono<>sum_reipgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Freipgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_sellno %>
		    <% if oLastMonthstock.FOneItem.Fsellno<>sum_sellno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fsellno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_resellno %>
		    <% if oLastMonthstock.FOneItem.Fresellno<>sum_resellno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fresellno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_offchulgono %>
		    <% if oLastMonthstock.FOneItem.Foffchulgono<>sum_offchulgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Foffchulgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_offrechulgono %>
		    <% if oLastMonthstock.FOneItem.Foffrechulgono<>sum_offrechulgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Foffrechulgono %>)</font>
		    <% end if %>
		</td>

		<td>
		    <%= sum_etcchulgono + sum_etcrechulgono %>
		    <% if (oLastMonthstock.FOneItem.Fetcchulgono+oLastMonthstock.FOneItem.Fetcrechulgono)<>(sum_etcchulgono + sum_etcrechulgono) then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fetcchulgono+oLastMonthstock.FOneItem.Fetcrechulgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_errcsno %>
		    <% if oLastMonthstock.FOneItem.Ferrcsno<>sum_errcsno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ferrcsno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <b><%= sum_totsysstock %></b>
		    <% if oLastMonthstock.FOneItem.Ftotsysstock<>sum_totsysstock then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ftotsysstock %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_errrealcheckno %>
		    <% if oLastMonthstock.FOneItem.Ferrrealcheckno<>sum_errrealcheckno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ferrrealcheckno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_totsysstock+sum_errrealcheckno %>
		    <% if oLastMonthstock.FOneItem.Ftotsysstock+oLastMonthstock.FOneItem.Ferrrealcheckno<>sum_totsysstock+sum_errrealcheckno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ftotsysstock+oLastMonthstock.FOneItem.Ferrrealcheckno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_errbaditemno %>
		    <% if oLastMonthstock.FOneItem.Ferrbaditemno<>sum_errbaditemno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ferrbaditemno %>)</font>
		    <% end if %>
		</td>
		<!--
		<td>
		    <b><%= sum_availsysstock %></b>
		    <% if oLastMonthstock.FOneItem.Favailsysstock<>sum_availsysstock then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Favailsysstock %>)</font>
		    <% end if %>
		</td>
		-->
		<td>
		    <b><%= sum_realstock %></b>
		    <% if oLastMonthstock.FOneItem.Frealstock<>sum_realstock then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Frealstock %>)</font>
		    <% end if %>
		</td>
		<td>

		</td>
	</tr>
<% end if %>
<!-- �Ϻ� �α� -->
<%
dim ismidSubtotalShow
%>
<% for i=0 to osummarystock.FResultCount-1 %>
<%
sum_ipgono = sum_ipgono + osummarystock.FItemList(i).Fipgono
sum_reipgono = sum_reipgono + osummarystock.FItemList(i).Freipgono
sum_sellno = sum_sellno + osummarystock.FItemList(i).Fsellno
sum_resellno = sum_resellno + osummarystock.FItemList(i).Fresellno
sum_offchulgono = sum_offchulgono + osummarystock.FItemList(i).Foffchulgono
sum_offrechulgono = sum_offrechulgono + osummarystock.FItemList(i).Foffrechulgono
sum_etcchulgono = sum_etcchulgono + osummarystock.FItemList(i).Fetcchulgono
sum_etcrechulgono = sum_etcrechulgono + osummarystock.FItemList(i).Fetcrechulgono
sum_errbaditemno	= sum_errbaditemno + osummarystock.FItemList(i).Ferrbaditemno
sum_errrealcheckno	= sum_errrealcheckno + osummarystock.FItemList(i).Ferrrealcheckno
sum_errcsno = sum_errcsno + osummarystock.FItemList(i).Ferrcsno
sum_totsysstock = sum_totsysstock + osummarystock.FItemList(i).Ftotsysstock
sum_availsysstock = sum_availsysstock + osummarystock.FItemList(i).Favailsysstock
sum_realstock = sum_realstock + osummarystock.FItemList(i).Frealstock

sysstock = sysstock + osummarystock.FItemList(i).Ftotsysstock
sysavailstock = sysavailstock + osummarystock.FItemList(i).Favailsysstock
realstock = realstock + osummarystock.FItemList(i).Frealstock
maystock = maystock + osummarystock.FItemList(i).Frealstock


mm_ipgono = mm_ipgono + osummarystock.FItemList(i).Fipgono
mm_reipgono = mm_reipgono + osummarystock.FItemList(i).Freipgono
mm_sellno = mm_sellno + osummarystock.FItemList(i).Fsellno
mm_resellno = mm_resellno + osummarystock.FItemList(i).Fresellno
mm_offchulgono = mm_offchulgono + osummarystock.FItemList(i).Foffchulgono
mm_offrechulgono = mm_offrechulgono + osummarystock.FItemList(i).Foffrechulgono
mm_etcchulgono = mm_etcchulgono + osummarystock.FItemList(i).Fetcchulgono
mm_etcrechulgono = mm_etcrechulgono + osummarystock.FItemList(i).Fetcrechulgono
mm_errbaditemno	= mm_errbaditemno + osummarystock.FItemList(i).Ferrbaditemno
mm_errrealcheckno	= mm_errrealcheckno + osummarystock.FItemList(i).Ferrrealcheckno
mm_errcsno  = mm_errcsno + osummarystock.FItemList(i).Ferrcsno

'sum_offsell = sum_offsell + osummarystock.FItemList(i).Foffsellno
'offstockno = offstockno + osummarystock.FItemList(i).Foffchulgono*-1 + osummarystock.FItemList(i).Foffrechulgono*-1 - osummarystock.FItemList(i).Foffsellno
%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummarystock.FItemList(i).Fyyyymmdd %>(<%= osummarystock.FItemList(i).GetDpartName %>)</td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>','I');"><%= osummarystock.FItemList(i).Fipgono %></a></td>
      	<td><%= osummarystock.FItemList(i).Freipgono %></td>
      	<td><a href="javascript:popBuyItemListChulgo('<%= osummarystock.FItemList(i).Fyyyymmdd %>');"><%= osummarystock.FItemList(i).Fsellno %></a></td>
      	<td><%= osummarystock.FItemList(i).Fresellno %></td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>','S');"><%= osummarystock.FItemList(i).Foffchulgono %></a></td>
      	<td><%= osummarystock.FItemList(i).Foffrechulgono %></td>

      	<td><a href="javascript:PopItemIpChulList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>','E');"><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></a></td>
    	<td><a href="javascript:popCSItemListChulgo('<%= osummarystock.FItemList(i).Fyyyymmdd %>')"><%= osummarystock.FItemList(i).Ferrcsno %></a></td>
        <td><%= sysstock %></td>
        <td><a href="javascript:popRealErrList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>')"><%= osummarystock.FItemList(i).Ferrrealcheckno %></a></td>
        <td><%= sysstock+sum_errrealcheckno %></td>
      	<td><a href="javascript:PopStockBaditem('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>')"><%= osummarystock.FItemList(i).Ferrbaditemno %></a></td>

      	<!-- td><%= sysavailstock %></td -->
      	<td><%= realstock %></td>
      	<td></td>

    </tr>
    <%
        ismidSubtotalShow = false

        if (i>=osummarystock.FResultCount-1) then
            ismidSubtotalShow = true
        elseif Left(osummarystock.FItemList(i).Fyyyymmdd,7)<>Left(osummarystock.FItemList(i+1).Fyyyymmdd,7) then
            ismidSubtotalShow = true
        end if

    %>
    <% if (ismidSubtotalShow) then %>
    <!-- ���� �հ� �߰� -->
    <tr align="center" bgcolor="#EEEEEE">
		<td><%= Left(osummarystock.FItemList(i).Fyyyymmdd,7) %></td>
		<td><%= mm_ipgono %></td>
		<td><%= mm_reipgono %></td>
		<td><%= mm_sellno %></td>
		<td><%= mm_resellno %></td>
		<td><%= mm_offchulgono %></td>
		<td><%= mm_offrechulgono %></td>

		<td><%= mm_etcchulgono + mm_etcrechulgono%></td>
		<td><%= mm_errcsno %></td>
        <td><b><%= sum_totsysstock %></b></td>
        <td><%= mm_errrealcheckno %></td>
        <td><%= sum_totsysstock+sum_errrealcheckno %></td>
		<td><%= mm_errbaditemno %></td>
		<!-- td><b><%= sum_availsysstock %></b></td -->
		<td><b><%= sum_realstock %></b></td>
        <td>
            <% if sum_realstock<>0 then %>
      	    <b><%= CLng((mm_sellno + mm_offchulgono)*-1/sum_realstock*100)/100 %></b>
      	    <% end if %>
			<% if (C_ADMIN_AUTH) then %>
			<input type="button" value="�⸻���ۼ� <%=Left(osummarystock.FItemList(i).Fyyyymmdd,7)%>" onClick="refreshAccStock(this,'<%=Left(osummarystock.FItemList(i).Fyyyymmdd,7)%>')">
			<% end if %>
        </td>
	</tr>
	<%
	mm_ipgono = 0
    mm_reipgono = 0
    mm_sellno = 0
    mm_resellno = 0
    mm_offchulgono = 0
    mm_offrechulgono = 0
    mm_etcchulgono = 0
    mm_etcrechulgono = 0
    mm_errbaditemno	= 0
    mm_errrealcheckno = 0
    mm_errcsno = 0
	%>
    <% end if %>
<% next %>
	<tr align="center" bgcolor="#EEEEEE">
		<td>ToTal</td>
		<td><%= sum_ipgono %></td>
		<td><%= sum_reipgono %></td>
		<td><%= sum_sellno %></td>
		<td><%= sum_resellno %></td>
		<td><%= sum_offchulgono %></td>
		<td><%= sum_offrechulgono %></td>

		<td><%= sum_etcchulgono + sum_etcrechulgono%></td>
		<td><%= sum_errcsno %></td>
        <td><b><%= sum_totsysstock %></b></td>
        <td><%= sum_errrealcheckno %></td>
        <td><%= sum_totsysstock+sum_errrealcheckno %></td>
		<td><%= sum_errbaditemno %></td>
		<!-- td><b><%= sum_availsysstock %></b></td -->
		<td><b><%= sum_realstock %></b></td>
        <td>
			<% if (C_ADMIN_AUTH) then %>
			<% if (osummarystock.FResultCount<1) then %>
				<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-1,now()),7)%>" onClick="refreshAccStock(this,'<%=LEFT(dateadd("m",-1,now()),7)%>')">
				<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-0,now()),7)%>" onClick="refreshAccStock(this,'<%=LEFT(dateadd("m",-0,now()),7)%>')">
			<% end if %>
			<% end if %>
		</td>

	</tr>
</table>

<p>

<%

dim colcount

colcount = 0
for i = 0 to offstock.FResultcount - 1
	if not IsNULL(offstock.FItemList(i).Fcurrno) then
		colcount = colcount + 1
	end if
next
''colcount = offstock.FResultCount
%>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="<%= colcount+1 %>">
			<b>*���� ���� ���</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   	<% for i=0 to offstock.FResultcount-1 %>
	    	<% if not IsNULL(offstock.FItemList(i).Fcurrno) then %>
	    		<td ><acronym title="<%= offstock.FItemList(i).Fshopname %>"><%= Right(offstock.FItemList(i).FShopid,3) %></acronym></td>
	    	<% end if %>
    	<% next %>
    	<td>�����</td>
    </tr>
	<tr align="center" bgcolor="#FFFFFF">
    	<% for i=0 to offstock.FResultcount-1 %>
    		<% if not IsNULL(offstock.FItemList(i).Fcurrno) then %>
    		<td>
				<% if (itemgubun = "10") then %>
				<a href="/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=<%= offstock.FItemList(i).FShopid %>&barcode=<%= barcode %>" target="_blank"><%= offstock.FItemList(i).Fcurrno %></a>
				<% else %>
				<%= offstock.FItemList(i).Fcurrno %>
				<% end if %>
			</td>
    		<% end if %>
    	<% next %>
    	<td><%= offtotalstock %></td>
    </tr>
</table>

<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">�˻� ����� �����ϴ�.</td>
    </tr>
</table>
<% end if %>

<%
if (oitemoption.FResultCount>0) and (itemoption="0000") then
    ErrMsg = "�ɼ� ���� �� �� �˻��ϼ���."
elseif (oitemoption.FResultCount<1) and (itemoption<>"0000") then
    ErrMsg = "�� �˻��ϼ���."
end if
%>

<form name=frmrefresh method=post action="stockrefresh_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
	<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="itemoption" value="<%= itemoption %>">
	<input type="hidden" name="yyyymm" value="">
</form>

<% if ErrMsg<>"" then %>
	<script language='javascript'>
		alert('<%= ErrMsg %>');
	</script>
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing
set osummaryMonthstock = Nothing
set osummarystock = Nothing
set offstock = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
