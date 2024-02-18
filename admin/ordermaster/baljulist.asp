<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ü�����Ʈ
' History : �̻� ����
'			2023.07.11 �ѿ�� ����(�˻����� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, showItemKind, refer, showOldData, showOldOrderData, nowdate,date1,date2,Edate
dim workgroup, baljuid
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	showItemKind = requestCheckVar(request("showItemKind"),1)
	showOldData = requestCheckVar(request("showOldData"),1)
	showOldOrderData = requestCheckVar(request("showOldOrderData"),1)
	workgroup = requestCheckVar(request("workgroup"),1)
	baljuid = requestCheckVar(getNumeric(request("baljuid")),10)

refer = request.ServerVariables("HTTP_REFERER")

if instr(refer,"baljulist.asp")<1 then
	autoChulgodateSet()
end if

nowdate = now

if (yyyy1="") then
	date1 = dateAdd("d",-4,nowdate)
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(dateserial(yyyy2, mm2 , dd2)+1),10)
end if

dim baljumaster
baljumaster = request("baljumaster")

dim obalju,i
set obalju = New CBalju
obalju.FStartdate = yyyy1 + "-" + mm1 + "-" + dd1
obalju.FEndDate = Left(CStr(Edate),10)
obalju.FRectShowItemKind = showItemKind
obalju.FRectOldDate = showOldData
obalju.FRectOldOrderData = showOldOrderData
obalju.FRectworkgroup = workgroup
obalju.FRectbaljuid = baljuid
obalju.getBaljumaster

''������û󼼳���
'if baljumaster<>"" then
'	obalju.getBaljuDetailList baljumaster
'end if

dim ppdate, ppcnt, sscnt, sumppcnt, sumsscnt, itemCnt, sumItemCnt, itemSortCnt, sumitemSortCnt, itemPickSortCnt, sumitemPickSortCnt, itemPickOptionSortCnt, sumitemPickOptionSortCnt, ItemnoBulk, sumItemnoBulk
dim itemSkuNo, itemSkuAgvNo, itemSkuBulkNo, giftSkuNo, sumitemSkuNo, sumitemSkuAgvNo, sumitemSkuBulkNo, sumgiftSkuNo
	sumppcnt = 0
	sumsscnt = 0
	sumItemCnt = 0
	sumitemSortCnt = 0
    sumitemPickSortCnt = 0
    sumitemPickOptionSortCnt = 0
    sumItemnoBulk = 0
    sumitemSkuNo = 0
    sumitemSkuAgvNo = 0
    sumitemSkuBulkNo = 0
    sumgiftSkuNo = 0

dim SubChulgoCount, Subdelay0chulgocnt, Subdelay1chulgocnt, Subdelay2chulgocnt, Subdelay3chulgocnt, SubCancelCnt, SubMichulgoCnt
dim SumChulgoCount, Sumdelay0chulgocnt, Sumdelay1chulgocnt, Sumdelay2chulgocnt, Sumdelay3chulgocnt, SumCancelCnt, SumMichulgoCnt

%>

<script type='text/javascript'>

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function AnViewUpcheList(iid){
	var popwin = window.open('/admin/pop/viewupchelist.asp?iid=' + iid,'viewupchelist','width=780,height=700,scrollbars=yes');
	popwin.focus();
}

function ViewAddDetailList(){
    var popwin = window.open('/admin/ordermaster/pop_makeonorder_list_UTF8.asp?research=on&menupos=44','pop_makeonorder_list','width=1100,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ViewQuickOrderList() {
	var popwin = window.open('/admin/ordermaster/pop_QuickOrder_list.asp?research=on&menupos=44','ViewQuickOrderList','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsPopLogisticsBaljuList(sitebaljukey) {
    var popwin = window.open("pop_logistics_baljuitemlist.asp?sitebaljukey=" + sitebaljukey,"jsPopLogisticsBaljuList","width=800,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsPopLogisticsBaljuitem(sitebaljukey) {
    var popwin = window.open("/admin/ordermaster/pop_logistics_baljuitem.asp?sitebaljukey=" + sitebaljukey,"jsPopLogisticsBaljuitem","width=800,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function ViewBaljuArr(){
	var frm;
	var pass = false;
	var idxarr = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �ֹ��� �����ϴ�.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (idxarr==""){
					idxarr = frm.iidx.value;
				}else{
					idxarr = idxarr + "," + frm.iidx.value;
				}
			}
		}
	}

	window.open('/admin/pop/viewitemlist.asp?idxarr=' + idxarr,'','');
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" bgcolor="#ffffff">
		* ��ȸ�Ⱓ : <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		* �۾��׷� : <% DrawWorkgroup "workgroup", workgroup, "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center">
	<td align="left" bgcolor="#ffffff">
		* �������ID <input type="text" name="baljuid" value="<%= baljuid %>" maxlength=10 size=8>
		&nbsp;
		<input type="checkbox" name="showItemKind" value="Y" <%= CHKIIF(showItemKind="Y", "checked", "") %>> ��ǰ ������ ǥ��
		&nbsp;
		<input type="checkbox" name="showOldData" value="Y" <%= CHKIIF(showOldData="Y", "checked", "") %>> ���ų��� ǥ��(�������)
		&nbsp;
		<input type="checkbox" name="showOldOrderData" value="Y" <%= CHKIIF(showOldOrderData="Y", "checked", "") %>> ���ų��� ǥ��(��ǰ����)
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
    <tr height="25">
        <td align="left">
			<!--
        	<input type="button" class="button" value="����������þ����۸�Ϻ���" onclick="ViewBaljuArr()">
			-->
        </td>
        <td align="right">
            <input type="button" class="button" value="����� �ֹ����" onclick="ViewQuickOrderList()">
			&nbsp;
			<input type="button" class="button" value="�ֹ�����List" onclick="ViewAddDetailList()">
        </td>
    </tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		�˻���� : <b><%= obalju.FTotalCount %></b>
	</td>
</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=20></td>
		<td width="80">�������ID</td>
		<td>��������Ͻ�</td>
		<td>�ֹ�<br>����Ʈ</td>
		<td>����</td>
		<td>�׷�</td>
		<td>�������<br>Ÿ��</td>
		<td>�ù��</td>
		<td>��������ð�</td>
		<td>(%)</td>
		<td>��ü���<br>�Ǽ�</td>
		<td>��ǰ��<br>(���X)</td>
		<td>SKU<br />(��ü)</td>
        <td>SKU<br />(AGV)</td>
        <td>SKU<br />(BULK)</td>
        <td>SKU<br />(����ǰ)</td>
		<td>���</td>
		<td>��(�ٹ�)<br>���Ǽ�</td>
		<td>����<br>���Ǽ�</td>
		<td>1��<br>�������</td>
		<td>2��<br>�������</td>
		<td>3��<br>�������</td>
		<td>�����</td>
		<td>�ֹ�<br>����Ʈ</td>
		<td>��ǰ<br>����Ʈ</td>
		<!--
		<td>��ǰ<br>���</td>
		<td>��������</td>
		-->
		<td>����ǰ<br>����Ʈ</td>
		<td>����ǰ<br>�հ�</td>
		<td>�ٹ�<br>�ֹ�����</td>
		<td>������ü����</td>
	</tr>
<% if (obalju.resultBaljucount<1) then %>
	<tr bgcolor="#FFFFFF" height="31"><td colspan="29" align="center">�ش�Ⱓ�� ������ü�����</td></tr>
<% else %>

<% for i=0 to obalju.resultBaljucount-1 %>
<%
if ppdate<>Left(obalju.FBaljumasterList(i).FBaljudate,10) then
	ppdate = Left(obalju.FBaljumasterList(i).FBaljudate,10)
	ppcnt = 0
	sscnt = 0
	itemCnt = 0
	itemSortCnt = 0
    itemPickSortCnt = 0
    itemPickOptionSortCnt = 0
    ItemnoBulk = 0
    itemSkuNo = 0
    itemSkuAgvNo = 0
    itemSkuBulkNo = 0
    giftSkuNo = 0
    SubChulgoCount      = 0
    Subdelay0chulgocnt  = 0
    Subdelay1chulgocnt  = 0
    Subdelay2chulgocnt  = 0
    Subdelay3chulgocnt  = 0
    SubMichulgoCnt      = 0
    SubCancelCnt        = 0
end if

ppcnt = ppcnt + obalju.FBaljumasterList(i).FCount
sscnt = sscnt + obalju.FBaljumasterList(i).Fsongjangcnt
itemCnt = itemCnt + obalju.FBaljumasterList(i).Fitemno
itemSortCnt = itemSortCnt + obalju.FBaljumasterList(i).FitemSortNo
itemPickSortCnt = itemPickSortCnt + obalju.FBaljumasterList(i).FitemPickSortNo
itemPickOptionSortCnt = itemPickOptionSortCnt + obalju.FBaljumasterList(i).FitemPickOptionSortNo
ItemnoBulk = ItemnoBulk + obalju.FBaljumasterList(i).FitemnoBulk
itemSkuNo = itemSkuNo + obalju.FBaljumasterList(i).FitemSkuNo
itemSkuAgvNo = itemSkuAgvNo + obalju.FBaljumasterList(i).FitemSkuAgvNo
itemSkuBulkNo = itemSkuBulkNo + obalju.FBaljumasterList(i).FitemSkuBulkNo
giftSkuNo = giftSkuNo + obalju.FBaljumasterList(i).FgiftSkuNo

SubChulgoCount      = SubChulgoCount +  obalju.FBaljumasterList(i).GetTotalChulgoCount
Subdelay0chulgocnt  = Subdelay0chulgocnt + obalju.FBaljumasterList(i).Fdelay0chulgocnt
Subdelay1chulgocnt  = Subdelay1chulgocnt + obalju.FBaljumasterList(i).Fdelay1chulgocnt
Subdelay2chulgocnt  = Subdelay2chulgocnt + obalju.FBaljumasterList(i).Fdelay2chulgocnt
Subdelay3chulgocnt  = Subdelay3chulgocnt + obalju.FBaljumasterList(i).Fdelay3chulgocnt
SubMichulgoCnt      = SubMichulgoCnt + obalju.FBaljumasterList(i).GetTenMiChulgoCount
SubCancelCnt        = SubCancelCnt + obalju.FBaljumasterList(i).FCancelCnt
%>


	<form name="frmBuyPrc_<%= obalju.FBaljumasterList(i).FBaljuID %>" method="post" >
	<input type="hidden" name="iidx" value="<%= obalju.FBaljumasterList(i).FBaljuID %>">
	<tr bgcolor="#FFFFFF" align="center">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= obalju.FBaljumasterList(i).FBaljuID %></td>
		<td align="left"><%= obalju.FBaljumasterList(i).FBaljudate %></td>
		<td><%= obalju.FBaljumasterList(i).GetExtSiteName %></td>
		<td><%= obalju.FBaljumasterList(i).Fdifferencekey %></td>
		<td><%= obalju.FBaljumasterList(i).Fworkgroup %></td>
		<td><%= obalju.FBaljumasterList(i).getBaljuTypeName %></td>
		<td><%= obalju.FBaljumasterList(i).getDeliverName %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).FCount,0) %></td>
		<td>
		<% if obalju.FBaljumasterList(i).FCount<>0 then %>
			<%= CLng(obalju.FBaljumasterList(i).Fsongjangcnt/obalju.FBaljumasterList(i).FCount*100) %> %
		<% end if %>
        </td>
        <td><%= FormatNumber(obalju.FBaljumasterList(i).Fsongjangcnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fitemno,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).FitemSkuNo,0) %></td>
        <td><%= FormatNumber(obalju.FBaljumasterList(i).FitemSkuAgvNo,0) %></td>
        <td><%= FormatNumber(obalju.FBaljumasterList(i).FitemSkuBulkNo,0) %></td>
        <td><%= FormatNumber(obalju.FBaljumasterList(i).FgiftSkuNo,0) %></td>
    	<td><%= FormatNumber(obalju.FBaljumasterList(i).Fcancelcnt,0) %></td>
	    <td><%= FormatNumber(obalju.FBaljumasterList(i).GetTotalChulgoCount,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay0chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay1chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay3chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).GetTenMiChulgoCount,0) %></td>
		<td><a href="popbaljulist.asp?idx=<%= obalju.FBaljumasterList(i).FBaljuID %>&songjangDiv=<%= obalju.FBaljumasterList(i).FsongjangDiv%>" target=_blank>����</a></td>
		<td>
			<a href="#" onClick="jsPopLogisticsBaljuitem(<%= obalju.FBaljumasterList(i).FBaljuID %>); return false;">����</a>
		</td>
		<!--
		<td><a href="/admin/pop/viewupchelist.asp?iid=<%= obalju.FBaljumasterList(i).FBaljuID %>" target=_blank>����</a></td>
		<td><a href="/admin/ordermaster/popsongjangmaker.asp?iid=<%= obalju.FBaljumasterList(i).FBaljuID %>" target=_blank>����</a></td>
	    -->
	    <td><a href="/admin/ordermaster/poporder_gift.asp?baljuid=<%= obalju.FBaljumasterList(i).FBaljuID %>" target=_blank>����</a></td>
	    <td><a href="/admin/ordermaster/poporder_gift_summary.asp?menupos=1011&research=on&evt_code=&balju_code=<%= obalju.FBaljumasterList(i).FBaljuID %>&viewType=summary&isupchebeasong=N&dateview1=yes&date_display=on" target=_blank>����</a></td>
		<td><a href="/admin/ordermaster/pop_makeonorder_list_UTF8.asp?menupos=1568&balju_code=<%= obalju.FBaljumasterList(i).FBaljuID %>" target="_blank">����</a></td>
		<td>
			<input type="button" class="button" value="���" onClick="jsPopLogisticsBaljuList(<%= obalju.FBaljumasterList(i).FBaljuID %>)">
		</td>
	</tr>
	</form>

<% if i+1<obalju.resultBaljucount then %>
	<% if (ppdate<>Left(obalju.FBaljumasterList(i+1).FBaljudate,10)) then %>
<%
sumppcnt = sumppcnt + ppcnt
sumsscnt = sumsscnt + sscnt
sumItemCnt = sumItemCnt + itemCnt
sumitemSortCnt = sumitemSortCnt + itemSortCnt
sumitemPickSortCnt = sumitemPickSortCnt + itemPickSortCnt
sumitemPickOptionSortCnt = sumitemPickOptionSortCnt + itemPickOptionSortCnt
sumItemnoBulk = sumItemnoBulk + ItemnoBulk
sumitemSkuNo = sumitemSkuNo + itemSkuNo
sumitemSkuAgvNo = sumitemSkuAgvNo + itemSkuAgvNo
sumitemSkuBulkNo = sumitemSkuBulkNo + itemSkuBulkNo
sumgiftSkuNo = sumgiftSkuNo + giftSkuNo

SumChulgoCount      = SumChulgoCount + SubChulgoCount
Sumdelay0chulgocnt  = Sumdelay0chulgocnt + Subdelay0chulgocnt
Sumdelay1chulgocnt  = Sumdelay1chulgocnt + Subdelay1chulgocnt
Sumdelay2chulgocnt  = Sumdelay2chulgocnt + Subdelay2chulgocnt
Sumdelay3chulgocnt  = Sumdelay3chulgocnt + Subdelay3chulgocnt
SumMichulgoCnt      = SumMichulgoCnt + SubMichulgoCnt
SumCancelCnt        = SumCancelCnt + SubCancelCnt
%>
	<tr align="center" bgcolor="#EEEEEE">
		<td></td>
		<td>�Ұ�</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber(ppcnt,0) %></td>
		<td>
		<% if ppcnt<>0 then %>
			<%= CLng(sscnt/ppcnt*100) %> %
		<% end if %>
        </td>
        <td><%= FormatNumber(sscnt,0) %></td>
		<td><%= FormatNumber(itemCnt,0) %></td>
        <td><%= FormatNumber(itemSkuNo,0) %></td>
        <td><%= FormatNumber(itemSkuAgvNo,0) %></td>
        <td><%= FormatNumber(itemSkuBulkNo,0) %></td>
        <td><%= FormatNumber(giftSkuNo,0) %></td>
        <td><%= FormatNumber(SubCancelCnt,0) %></td>
        <td><font color="<%= ChkIIF(sscnt-SubCancelCnt-SubMichulgoCnt<>SubChulgoCount,"#FF0000","#000000") %>"><%= FormatNumber(SubChulgoCount,0) %></font></td>
    	<td><%= FormatNumber(Subdelay0chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay1chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(Subdelay3chulgocnt,0) %></td>

		<td>
		    <% if SubMichulgoCnt<>0 then %>
		    <b><font color="red"><%= FormatNumber(SubMichulgoCnt,0) %></font></b>
		    <% else %>
		    <%= FormatNumber(SubMichulgoCnt,0) %>
		    <% end if %>
		</td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<% end if %>
<% end if %>

<% next %>

<%
sumppcnt = sumppcnt + ppcnt
sumsscnt = sumsscnt + sscnt
sumItemCnt = sumItemCnt + itemCnt
sumitemSortCnt = sumitemSortCnt + itemSortCnt
sumitemPickSortCnt = sumitemPickSortCnt + itemPickSortCnt
sumitemPickOptionSortCnt = sumitemPickOptionSortCnt + itemPickOptionSortCnt
sumItemnoBulk = sumItemnoBulk + ItemnoBulk
sumitemSkuNo = sumitemSkuNo + itemSkuNo
sumitemSkuAgvNo = sumitemSkuAgvNo + itemSkuAgvNo
sumitemSkuBulkNo = sumitemSkuBulkNo + itemSkuBulkNo
sumgiftSkuNo = sumgiftSkuNo + giftSkuNo

SumChulgoCount      = SumChulgoCount + SubChulgoCount
Sumdelay0chulgocnt  = Sumdelay0chulgocnt + Subdelay0chulgocnt
Sumdelay1chulgocnt  = Sumdelay1chulgocnt + Subdelay1chulgocnt
Sumdelay2chulgocnt  = Sumdelay2chulgocnt + Subdelay2chulgocnt
Sumdelay3chulgocnt  = Sumdelay3chulgocnt + Subdelay3chulgocnt
SumMichulgoCnt      = SumMichulgoCnt + SubMichulgoCnt
SumCancelCnt        = SumCancelCnt + SubCancelCnt
%>

<% end if %>
	<tr align="center" bgcolor="#EEEEEE">
		<td></td>
		<td>�Ұ�</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber(ppcnt,0) %></td>
		<td>
		<% if ppcnt<>0 then %>
			<%= CLng(sscnt/ppcnt*100) %> %
		<% end if %>
    	</td>
    	<td><%= FormatNumber(sscnt,0) %></td>
		<td><%= FormatNumber(itemCnt,0) %></td>
		<td><%= FormatNumber(itemSkuNo,0) %></td>
        <td><%= FormatNumber(itemSkuAgvNo,0) %></td>
        <td><%= FormatNumber(itemSkuBulkNo,0) %></td>
        <td><%= FormatNumber(giftSkuNo,0) %></td>

		<td><%= FormatNumber(SubCancelCnt,0) %></td>
		<td><%= FormatNumber(SubChulgoCount,0) %></td>
    	<td><%= FormatNumber(Subdelay0chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay1chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(Subdelay3chulgocnt,0) %></td>
		<td>
		    <% if SubMichulgoCnt<>0 then %>
		    <b><font color="red"><%= FormatNumber(SubMichulgoCnt,0) %></font></b>
		    <% else %>
		    <%= FormatNumber(SubMichulgoCnt,0) %>
		    <% end if %>
		</td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr align="center" bgcolor="#EEEEEE">
		<td></td>
		<td>������ �հ�</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber(sumppcnt,0) %></td>
		<td>
		<% if sumppcnt<>0 then %>
			<%= CLng(sumsscnt/sumppcnt*100) %> %
		<% end if %>
        </td>
        <td><%= FormatNumber(sumsscnt,0) %></td>
		<td><%= FormatNumber(sumItemCnt,0) %></td>
		<td><%= FormatNumber(sumitemSkuNo,0) %></td>
        <td><%= FormatNumber(sumitemSkuAgvNo,0) %></td>
        <td><%= FormatNumber(sumitemSkuBulkNo,0) %></td>
        <td><%= FormatNumber(sumgiftSkuNo,0) %></td>
        <td><%= FormatNumber(SumCancelCnt,0) %></td>

    	<td><%= FormatNumber(SumChulgoCount,0) %></td>
    	<td><%= FormatNumber(Sumdelay0chulgocnt,0) %></td>
    	<td><%= FormatNumber(Sumdelay1chulgocnt,0) %></td>
    	<td><%= FormatNumber(Sumdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(Sumdelay3chulgocnt,0) %></td>
		<td><%= FormatNumber(SumMichulgoCnt,0) %></td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
</table>

<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
