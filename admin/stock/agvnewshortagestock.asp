<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newshortagestockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
const C_STOCK_DAY=7

''�Ʒ� �� �������� �˻������� �����ϰ� ����� �Ѵ�.
''/admin/stock/newshortagestock.asp
''/admin/newstorage/popjumunitemNew.asp

dim page, mode, makerid, shopid,itemid, research
dim onlynotupchebeasong, onlyusingitem, onlyusingitemoption, onlynotdanjong, soldoutover7days, onlysell, onlynottempdanjong
dim onlynotmddanjong, includepreorder, skiplimitsoldout
dim onoffgubun, idx, shortagetype, onlystockminus
dim changemakerid
dim purchasetype
dim itemgubun, itemname
dim chkMinusStockGubun, minusStockGubun
dim mwdiv, excmkr, onlyOn, pagesize, onlyrealup, ordBy

shopid = requestCheckVar(("shopid"),32)
page = requestCheckVar(request("page"),32)
mode = requestCheckVar(request("mode"),32)
itemid = requestCheckVar(request("itemid"),32)
research = requestCheckVar(request("research"),32)
onlynotupchebeasong = requestCheckVar(request("onlynotupchebeasong"),32)
onlyusingitem = requestCheckVar(request("onlyusingitem"),32)
onlyusingitemoption = requestCheckVar(request("onlyusingitemoption"),32)
onlynotdanjong = requestCheckVar(request("onlynotdanjong"),32)
soldoutover7days = requestCheckVar(request("soldoutover7days"),32)
onoffgubun = requestCheckVar(request("onoffgubun"),32)
idx = requestCheckVar(request("idx"),32)
shortagetype = requestCheckVar(request("shortagetype"),32)
onlysell = requestCheckVar(request("onlysell"),32)
onlynottempdanjong = requestCheckVar(request("onlynottempdanjong"),32)
onlynotmddanjong = requestCheckVar(request("onlynotmddanjong"),32)
includepreorder = requestCheckVar(request("includepreorder"),32)
skiplimitsoldout = requestCheckVar(request("skiplimitsoldout"),32)
onlystockminus = requestCheckVar(request("onlystockminus"),32)
purchasetype = requestCheckVar(request("purchasetype"),32)
itemgubun = requestCheckVar(request("itemgubun"),32)
itemname = requestCheckVar(request("itemname"),128)
chkMinusStockGubun = requestCheckVar(request("chkMinusStockGubun"),32)
minusStockGubun = requestCheckVar(request("minusStockGubun"),32)
mwdiv = requestCheckVar(request("mwdiv"),32)
excmkr = requestCheckVar(request("excmkr"),32)
onlyOn = requestCheckVar(request("onlyOn"),32)
pagesize = requestCheckVar(request("pagesize"),32)
onlyrealup = requestCheckVar(request("onlyrealup"),32)
ordBy = requestCheckVar(request("ordBy"),32)

changemakerid = "Y"
if (changemakerid = "") then
	changemakerid = requestCheckVar(request("changemakerid"),32)
end if

makerid = request("makerid")
if (makerid = "") then
	makerid = requestCheckVar(request("suplyer"),32)
end if


if (research<>"on") then
	excmkr = "Y"
    chkMinusStockGubun="Y"
    minusStockGubun = "agv"
end if

if (research<>"on") and (onlynotupchebeasong = "") then
	onlynotupchebeasong = "on"
end if

if (research<>"on") and (onlyusingitem = "") then
	onlyusingitem = "on"
end if

if (research<>"on") and (onlyusingitemoption="") then
	onlyusingitemoption = "on"
end if

if (research<>"on") and (onlynotdanjong = "") then
	onlynotdanjong = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if (research<>"on") and (itemgubun="") then
	itemgubun = "10"
end if

if (research<>"on") and (shortagetype="") then
	shortagetype = "7day"
end if

if (research<>"on") and (includepreorder="") then
	includepreorder = "on"
end if

if (pagesize="") then
	pagesize = 100
end if

if (research<>"on") and (onlyrealup="") then
	onlyrealup = "on"
end if



if page="" then page=1
if mode="" then mode="bybrand"
'��ǰ�ڵ� ��ȿ�� �˻�(2008.07.31;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim oshortagestock
set oshortagestock  = new CShortageStock
oshortagestock.FPageSize = pagesize
oshortagestock.FCurrPage = page

oshortagestock.FRectOnlySell			= onlysell
oshortagestock.FRectOnlyUsingItem		= onlyusingitem
oshortagestock.FRectOnlyUsingItemOption	= onlyusingitemoption
oshortagestock.FRectOnlyNotUpcheBeasong	= onlynotupchebeasong

oshortagestock.FRectOnlyNotDanjong		= onlynotdanjong
oshortagestock.FRectOnlyNotTempDanjong	= onlynottempdanjong
oshortagestock.FRectOnlyNotMDDanjong	= onlynotmddanjong
oshortagestock.FRectSkipLimitSoldOut	= skiplimitsoldout

oshortagestock.FRectPurchaseType		= purchasetype

oshortagestock.FRectMakerid				= makerid
oshortagestock.FRectItemId				= itemid
'oshortagestock.FRectItemOption			= makerid

oshortagestock.FRectItemGubun			= itemgubun

if (chkMinusStockGubun = "Y") then
	oshortagestock.FRectMinusStockGubun			= minusStockGubun
end if

if (itemname <> "") then
	if (makerid <> "") then
		oshortagestock.FRectItemName			= itemname
	else
		response.write "<script>alert('���� �귣�带 �����ϼ���.');</script>"
	end if
end if

oshortagestock.FRectMWDiv				= mwdiv
oshortagestock.FRectExcMkr				= excmkr
oshortagestock.FRectOnlyOn				= onlyOn
oshortagestock.FRectOnlyRealUp			= onlyrealup
oshortagestock.FRectOrderBy				= ordBy
oshortagestock.FRectAGVCheck			= "Y"
if (itemgubun = "10") then
	oshortagestock.GetShortageItemListOnline
else
	oshortagestock.GetShortageItemListOffline
end if



dim i, shopsuplycash, buycash
dim IsAvailDelete



'==============================================================================
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, iStartDate, iEndDate

'���԰�����
'���ñ��� +- �������� ������ ǥ�� / �� �̿� ȸ��ǥ��
if (yyyy1="") then
    nowdate = Left(CStr(DateAdd("d",now(),-7)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

    nowdate = Left(CStr(DateAdd("d",now(),+7)),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

iStartDate  = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
iEndDate    = Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ChangeReqDay(frm){
	if (!(IsDigit(frm.maxsellday.value))){
		alert('���ڸ� �����մϴ�.');
		return;
	}

	if (confirm('�ʿ� ��� �������� �����Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function Research(page){
	document.frm.page.value= page;
	document.frm.action= "";
	document.frm.target= "";
	document.frm.submit();
}

function DeleteStockLog(itemgubun,itemid,itemoption){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frmdelstock.target="_blank";
        frmdelstock.itemgubun.value = itemgubun;
        frmdelstock.itemid.value = itemid;
        frmdelstock.itemoption.value = itemoption;
        frmdelstock.submit();
    }
}

function search(frm){
	/*
	if ((frm.makerid.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('�귣�带 ���� �ϼ���.');
			frm.designer.focus();
			return;
		}
	}
	*/
	document.frm.page.value = 1;
	document.frm.action= "";
	document.frm.target= "";
	frm.submit();
}

function jsUpdateAgvStockInfo() {
    var url;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'http://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=currstockall';

    if (confirm('AGV��� ���ΰ�ħ �Ͻðڽ��ϱ�?') != true) { return; }

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

function RefreshAgvStock(barcode) {
    var url;
    var brandArray;
    var skuCdArray;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'http://wapi.10x10.co.kr';
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
    url = 'http://wapi.10x10.co.kr';
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

function jsPopChgStockGubun() {
	var v = "popChgStockGubun.asp";
	var popwin = window.open(v,"jsPopChgStockGubun","width=250,height=150,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function exelagvstock() {
	alert('�����;��� ���� ���� �ɸ��ϴ�. ��ٷ��ּ���.')
	document.frm.action= "/admin/stock/agvnewshortagestock_excel.asp";
	document.frm.target= "view";
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<% if (changemakerid <> "Y") then %>
	<input type="hidden" name="makerid" value="<%= makerid %>" >
	<% else %>
	<input type="hidden" name="changemakerid" value="Y" >
	<% end if %>
	<input type="hidden" name="shopid" value="<%= shopid %>" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<% if (changemakerid <> "Y") then %>
			�귣�� : <b><%= makerid %></b>
			<% else %>
			�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			<% end if %>
			&nbsp;
			|
			&nbsp;
			���� :
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			<!--
			<select class="select" name="itemgubun">
				<option value="10" <% if (itemgubun = "10") then %>selected<% end if %> >�¶���(10)</option>
				<option value="90" <% if (itemgubun = "90") then %>selected<% end if %> >��������(90)</option>
				<option value="70" <% if (itemgubun = "70") then %>selected<% end if %> >����ǰ ��(70)</option>
				<option value="80" <% if (itemgubun = "80") then %>selected<% end if %> >����ǰ ��(80)</option>
				<option value="XX" <% if (itemgubun = "XX") then %>selected<% end if %> >��Ÿ</option>
			</select>
			-->
			&nbsp;
			|
			&nbsp;
			<input type=checkbox name="onlyusingitem" <% if onlyusingitem = "on" then response.write "checked" %> >����ǰ��
			<input type=checkbox name="onlyusingitemoption" <% if onlyusingitemoption = "on" then response.write "checked" %> >���ɼǸ�
			<input type=checkbox name="onlysell" <% if onlysell = "on" then response.write "checked" %> >�ǸŻ�ǰ��
			<input type=checkbox name="onlynotupchebeasong" <% if onlynotupchebeasong = "on" then response.write "checked" %> >��ü�������
			&nbsp;
			|
			&nbsp;
			<input type=checkbox name="onlynotdanjong" <% if onlynotdanjong = "on" then response.write "checked" %> >��������
			<input type=checkbox name="onlynottempdanjong" <% if onlynottempdanjong = "on" then response.write "checked" %> >�Ͻ�ǰ������
			<input type=checkbox name="onlynotmddanjong" <% if onlynotmddanjong = "on" then response.write "checked" %> >MD��������
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:search(frm);">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=8 maxlength=7>
			&nbsp;
			��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size=16 maxlength=16>
			&nbsp;
			|
			&nbsp;
			�ŷ����� :
			<select class="select" name="mwdiv">
				<option value="">-����-</option>
				<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >����</option>
				<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >Ư��</option>
				<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >��ü</option>
				<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >������</option>
			</select>
			&nbsp;
			<% if (FALSE) then %>
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", CHKIIF(purchasetype="", "1", purchasetype), "" %> <!-- ������. by eastone -->
			<% else %>
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		    <% end if %>
			&nbsp;
			|
			&nbsp;
			<input type="checkbox" name="chkMinusStockGubun" value="Y" <%if (chkMinusStockGubun = "Y") then %>checked<% end if %> >
			����� :
			<select class="select" name="minusStockGubun">
                <option value="agv" <%if (minusStockGubun = "agv") then %>selected<% end if %> >AGV ���</option>
				<option value="real" <%if (minusStockGubun = "real") then %>selected<% end if %> >�ǻ���ȿ���</option>
				<option value="check" <%if (minusStockGubun = "check") then %>selected<% end if %> >����ľ����</option>
				<option value="may" <%if (minusStockGubun = "may") then %>selected<% end if %> >�������</option>
			</select>
			���̳ʽ���
			&nbsp;
			|
			&nbsp;
			<input type="checkbox" class="checkbox" name="excmkr" value="Y" <%= CHKIIF(excmkr="Y", "checked", "")%> > ���̶������
			&nbsp;
			|
			&nbsp;
			<input type="checkbox" class="checkbox" name="onlyOn" value="Y" <%= CHKIIF(onlyOn="Y", "checked", "")%> > 7���Ǹ�1�̻�
            &nbsp;
			|
			&nbsp;
			<input type="checkbox" class="checkbox" name="onlyrealup" <%= CHKIIF(onlyrealup="on", "checked", "")%> > �ǻ����1�̻�
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			������������ :
			<select class="select" name="pagesize">
				<option value="50" <%= CHKIIF(pagesize="50", "selected", "") %>>50</option>
                <option value="100" <%= CHKIIF(pagesize="100", "selected", "") %>>100</option>
                <option value="500" <%= CHKIIF(pagesize="500", "selected", "") %>>500</option>
                <option value="1000" <%= CHKIIF(pagesize="1000", "selected", "") %>>1000</option>
			</select>

            <!-- ���� ������
            �������� :
			<select class="select" name="ordBy">
				<option value="makerid" <%= CHKIIF(ordBy="makerid", "selected", "") %>>�귣��</option>
                <option value="subrackcode" <%= CHKIIF(ordBy="subrackcode", "selected", "") %>>�������ڵ�</option>
			</select>
            -->
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<br>

<input type="button" class="button" value="����� ��ũ ��ȯ" onClick="jsPopChgStockGubun()">
<input type="button" class="button" value="�����ٿ�ε�" onClick="exelagvstock();">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frmshortage" method=post action="doshortagestock.asp">
	<input type="hidden" name="mode" value="maxsellday">
	<tr>
		<td align="left">
			<!--
			<input type="text" class="text" name="maxsellday" size="2" value="" maxlength=2>�� ��������
			<input type="button" class="button" value="����" onClick="ChangeReqDay(frmshortage);">
			-->
		</td>
		<td align="right">

		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="18" bgcolor="FFFFFF">
		<td colspan="25">
			�˻���� : <b><%= oshortagestock.FResultCount %></b>
			&nbsp;
			(�ִ�˻��Ǽ� : <%= oshortagestock.FTotalCount %>)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��ID</td>
        <td width="80">���ڵ�</td>
        <td width="80">������</td>
		<td width="30">����</td>
		<td width="40">��ǰ<br>�ڵ�</td>
		<td width="40">�ɼ�<br>�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="35" bgcolor="#F3F3FF"><b>�ǻ�<br>��ȿ<br>���<br>(V)</b></td>
		<td width="35" bgcolor="#F3F3FF"><b>���<br>�ľ�<br>���</b></td>
		<td width="35" bgcolor="#F3F3FF"><b>����<br>���</b></td>
        <td width="35" bgcolor="#F3F3FF"><b>AGV<br>���</b></td>
		<td width="40">ON<br>����<br>�Ϸ�</td>
        <td width="40">ON<br>������</td>
		<td width="40">OFF<br>������</td>

		<td width="50" bgcolor="#F3F3FF"><b>��(<%= C_STOCK_DAY %>��)<br>�ʿ�<br>����</b></td>
		<td width="30">�������<br>�ʿ���� <!-- OFF<br>�ֹ� --></td>
		<td width="30" bgcolor="#F3F3FF"><b>AGV����<br>����</b></td>
		<td width="40">ON<br>(7��)<br>�Ǹ�</td>
		<td width="40">OFF<br>(7��)<br>�Ǹ�</td>
	</tr>
<% for i=0 to oshortagestock.FResultCount -1 %>
<%
    IsAvailDelete = (oshortagestock.FItemList(i).Ftotipgono=0) and (oshortagestock.FItemList(i).FtotSellNo=0) and (oshortagestock.FItemList(i).Fshortageno=0) and (oshortagestock.FItemList(i).Frealstock=0) and (oshortagestock.FItemList(i).Fpreorderno=0)
%>

	<% if oshortagestock.FItemList(i).IsInvalidOption then %>
	<tr align="center" bgcolor="#CCCCCC">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="javascript:RefreshAgvStockByBrand('<%= oshortagestock.FItemList(i).FMakerID %>')"><%= oshortagestock.FItemList(i).FMakerID %></a></td>
        <td><%= oshortagestock.FItemList(i).FrackcodeByOption %></td>
        <td><%= oshortagestock.FItemList(i).FsubRackcodeByOption %></td>
		<td><%= oshortagestock.FItemList(i).Fitemgubun %></td>
		<td><a href="javascript:RefreshAgvStock('<%= BF_MakeTenBarcode(oshortagestock.FItemList(i).Fitemgubun, oshortagestock.FItemList(i).Fitemid, oshortagestock.FItemList(i).Fitemoption) %>');"><%= oshortagestock.FItemList(i).FItemID %></a></td>
		<td><%= oshortagestock.FItemList(i).Fitemoption %></td>
    	<td width="50" align=center><img src="<%= oshortagestock.FItemList(i).FimageSmall %>" width=50 height=50></td>
		<td align="left">
			<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= oshortagestock.FItemList(i).FItemID %>&itemoption=<%= oshortagestock.FItemList(i).FItemOption %>" target=_blank ><%= oshortagestock.FItemList(i).FItemName %></a>
			<% if oshortagestock.FItemList(i).FItemOption <> "0000" then %>
				<% if oshortagestock.FItemList(i).Foptionusing="Y" then %>
					<br><font color="blue">[<%= oshortagestock.FItemList(i).FItemOptionName %>]</font>
				<% else %>
					<br><font color="#AAAAAA">[<%= oshortagestock.FItemList(i).FItemOptionName %>]</font>
				<% end if %>
			<% end if %>
		</td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).Frealstock %></b></td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).GetCheckStockNo %></b></td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).GetMaystock %></b></td>
        <td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).FAGVStock %></b></td>

		<td><%= oshortagestock.FItemList(i).FIpkumdiv4 %></td>
        <td><%= oshortagestock.FItemList(i).FIpkumdiv5 %></td>
		<td><%= oshortagestock.FItemList(i).Foffconfirmno %></td>

		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).Frequireno %></b></td>
		<td>
		    <!-- ������� �ʿ���� -->
		    <%= (oshortagestock.FItemList(i).Fipkumdiv5 + oshortagestock.FItemList(i).Foffconfirmno+oshortagestock.FItemList(i).Fipkumdiv4 + oshortagestock.FItemList(i).Fipkumdiv2 + oshortagestock.FItemList(i).Foffjupno)*-1 %>
		</td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).GetAGVShortageNo %></b></td>
		<td><%= oshortagestock.FItemList(i).Fsell7days %></td>
		<td><%= oshortagestock.FItemList(i).Foffchulgo7days %></td>
	</tr>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if oshortagestock.HasPreScroll then %>
		<a href="javascript:Research('<%= oshortagestock.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oshortagestock.StartScrollPage to oshortagestock.FScrollCount + oshortagestock.StartScrollPage - 1 %>
			<% if i>oshortagestock.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:Research('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oshortagestock.HasNextScroll then %>
			<a href="javascript:Research('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set oshortagestock = Nothing
%>
<form name="frmdelstock" method="post" action="doshortagestockrefresh.asp">

<input type="hidden" name="mode" value="dellog">
<input type="hidden" name="itemgubun">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
