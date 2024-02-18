<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1, yyyy2, mm2, stplace, targetGbn, itemgubun
dim ipgoMWdiv, itemMWdiv, itemid
dim startYYYYMMDD, endYYYYMMDD
dim addInfoType
dim lastmwdiv, lastmakerid, lastCenterMWDiv
dim tmpDate
dim shopid
dim IpgoType


page       		= requestCheckvar(request("page"),10)
research		= requestCheckvar(request("research"),10)
yyyy1       	= requestCheckvar(request("yyyy1"),10)
mm1         	= requestCheckvar(request("mm1"),10)
yyyy2       	= requestCheckvar(request("yyyy2"),10)
mm2         	= requestCheckvar(request("mm2"),10)
stplace     	= requestCheckvar(request("stplace"),10)
targetGbn   	= requestCheckvar(request("targetGbn"),10)
itemgubun   	= requestCheckvar(request("itemgubun"),10)
ipgoMWdiv   	= requestCheckvar(request("ipgoMWdiv"),10)
itemMWdiv   	= requestCheckvar(request("itemMWdiv"),10)
itemid   		= requestCheckvar(request("itemid"),10)
addInfoType		= requestCheckvar(request("addInfoType"),10)
lastmwdiv		= requestCheckvar(request("lastmwdiv"),10)
lastCenterMWDiv	= requestCheckvar(request("lastCenterMWDiv"),10)
lastmakerid		= requestCheckvar(request("lastmakerid"),32)
shopid			= requestCheckvar(request("shopid"),32)
IpgoType		= requestCheckvar(request("IpgoType"),32)


if (page="") then page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", -1, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
	yyyy2 = yyyy1
	mm2 = mm1
end if

if (addInfoType="") then addInfoType = "none"

'// ============================================================================
dim ojaego
set ojaego = new CMonthlyStock

ojaego.FPageSize = 100
ojaego.FCurrPage = page
ojaego.FRectStartYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectEndYYYYMM = yyyy2 + "-" + mm2
ojaego.FRectPlaceGubun = stplace
ojaego.FRectTargetGbn = targetGbn
ojaego.FRectItemGubun = itemgubun

ojaego.FRectIpgoMWdiv = ipgoMWdiv
ojaego.FRectItemMWdiv = itemMWdiv
ojaego.FRectItemid = itemid

ojaego.FRectMwDiv = lastmwdiv
ojaego.FRectLastCenterMWDiv = lastCenterMWDiv
ojaego.FRectMakerid = lastmakerid
ojaego.FRectShopid = shopid
ojaego.FRectIpgoType = IpgoType

ojaego.GetMonthlyIpgoList

'// ============================================================================
dim ojaegosum
set ojaegosum = new CMonthlyStock

ojaegosum.FRectStartYYYYMM = yyyy1 + "-" + mm1
ojaegosum.FRectEndYYYYMM = yyyy2 + "-" + mm2

if (addInfoType = "sum") or (addInfoType = "sumshop") then
	ojaegosum.FRectPlaceGubun = stplace
	ojaegosum.FRectTargetGbn = targetGbn
	ojaegosum.FRectItemGubun = itemgubun
	ojaegosum.FRectIpgoMWdiv = ipgoMWdiv
	ojaegosum.FRectItemMWdiv = itemMWdiv
	ojaegosum.FRectItemid = itemid
	ojaegosum.FRectMwDiv = lastmwdiv
	ojaegosum.FRectLastCenterMWDiv = lastCenterMWDiv
	ojaegosum.FRectMakerid = lastmakerid
	ojaegosum.FRectShopid = shopid
	ojaegosum.FRectShowShopid = shopid
	ojaegosum.FRectIpgoType = IpgoType

	if (addInfoType = "sumshop") then
		ojaegosum.FRectShowShopid = "Y"
	end if

	ojaegosum.GetMonthlyIpgoSum
elseif (addInfoType = "diff") then
	ojaegosum.FPageSize = 20
	ojaegosum.FCurrPage = 1
	ojaegosum.FRectPlaceGubun = stplace

	ojaegosum.FRectYYYYMM = yyyy2 + "-" + mm2

	ojaegosum.GetMonthlyIpgoDiff
end if

startYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
endYYYYMMDD = Left(DateAdd("d", -1, DateSerial(yyyy1, mm1 + 1, 1)), 10)

dim totItemNoSUM, totBuyCashSUM
totItemNoSUM = 0
totBuyCashSUM = 0

%>

<script language='javascript'>

function fnUpdateIpgoList(stockPlace) {
	var frm = document.frm;
	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;

	if (!confirm(yyyymm + " ���� �԰��� ���Ӹ��� ���ۼ� �Ͻðڽ��ϱ�?")) {
		return;
	}

	var popwin = window.open("monthlystock_ipgoSum_process.asp?mode=monthlystockipgo&yyyymm=" + yyyymm + "&stockPlace=" + stockPlace,"fnUpdateIpgoList","width=100,height=100");
	popwin.focus();
}

function fnUpdateavgIpgoPrice(stockPlace) {
	var frm = document.frm;
	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;

	if (!confirm(yyyymm + " ��ո��԰��� ���ۼ� �Ͻðڽ��ϱ�?")) {
		return;
	}

	var popwin = window.open("monthlystock_ipgoSum_process.asp?mode=monthlystockavgipgoprice&yyyymm=" + yyyymm + "&stockPlace=" + stockPlace,"fnUpdateavgIpgoPrice","width=100,height=100");
	popwin.focus();
}

function addDays(theDate, days) {
    return new Date(theDate.getTime() + days*24*60*60*1000);
}

function Format00(d) {
    return (d < 10) ? '0' + d.toString() : d.toString();
}

function popIpgoDetailList(yyyymm, ipgoType, itemgubun, itemid) {
	var parts = (yyyymm + "-01").split("-");
	var dt1, dt2;
	dt1 = new Date(parts[0], parts[1] - 1, parts[2]);
	dt2 = new Date(parts[0], parts[1], parts[2]);
	dt2 = addDays(dt2, -1);

	var yyyy1 = dt1.getFullYear(), mm1 = Format00(dt1.getMonth() + 1), dd1 = Format00(dt1.getDate());
	var yyyy2 = dt2.getFullYear(), mm2 = Format00(dt2.getMonth() + 1), dd2 = Format00(dt2.getDate());

	var gubun = "I";
	if (ipgoType != "normal")  {
		gubun = "S";
	}

	window.open("/admin/storage/itemipchullist.asp?menupos=168&gubun=" + gubun + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&yyyy2=" + yyyy2 + "&mm2=" + mm2 + "&dd2=" + dd2);
}

function jsSetNotAssignedMWDiv() {
	var frm = document.frm;
	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;

	if (!confirm(yyyymm + " ���Ա��� ���������� ���Լ��� �Ͻðڽ��ϱ�?")) {
		return;
	}

	var popwin = window.open("monthlystock_ipgoSum_process.asp?mode=setmwdiv2m&yyyymm=" + yyyymm,"jsSetNotAssignedMWDiv","width=100,height=100");
	popwin.focus();
}

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			&nbsp;
			<font color="#CC3333">��/�� :</font> <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %> �� �԰���
			&nbsp;
			<font color="#CC3333">�԰�ó:</font>
		    <select name="stplace" class="select">
        		<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >����</option>
        		<option value="S" <%= CHKIIF(stplace="S","selected" ,"") %> >����</option>
				<option value="O" <%= CHKIIF(stplace="O","selected" ,"") %> >�¶�������</option>
				<option value="F" <%= CHKIIF(stplace="F","selected" ,"") %> >��������</option>
				<option value="A" <%= CHKIIF(stplace="A","selected" ,"") %> >�ΰŽ�����</option>
        	</select>
			&nbsp;
	    	<font color="#CC3333">�μ�����:</font>
	        <select name="targetGbn" class="select">
				<option value="" <%= CHKIIF(targetGbn="","selected" ,"") %> >��ü
				<option value="ON" <%= CHKIIF(targetGbn="ON","selected" ,"") %> >�¶���
				<option value="OF" <%= CHKIIF(targetGbn="OF","selected" ,"") %> >��������
				<option value="IT" <%= CHKIIF(targetGbn="IT","selected" ,"") %> >���̶��(��)
				<option value="ET" <%= CHKIIF(targetGbn="ET","selected" ,"") %> >3PL(���̶��)
				<option value="EG" <%= CHKIIF(targetGbn="EG","selected" ,"") %> >3PL(���׷���)
        	</select>
			&nbsp;
	    	<font color="#CC3333">�԰���:</font>
	        <select name="ipgoMWdiv" class="select">
				<option value="" <%= CHKIIF(ipgoMWdiv="","selected" ,"") %> >��ü</option>
				<option value="M" <%= CHKIIF(ipgoMWdiv="M","selected" ,"") %> >����</option>
				<option value="W" <%= CHKIIF(ipgoMWdiv="W","selected" ,"") %> >��Ź</option>
				<option value="X" <%= CHKIIF(ipgoMWdiv="X","selected" ,"") %> >��Ÿ</option>
        	</select>
			&nbsp;
	    	<font color="#CC3333">���Ա���(����):</font>
	        <select name="lastCenterMWDiv" class="select">
				<option value="" <%= CHKIIF(lastCenterMWDiv="","selected" ,"") %> >��ü</option>
				<option value="M" <%= CHKIIF(lastCenterMWDiv="M","selected" ,"") %> >����</option>
				<option value="W" <%= CHKIIF(lastCenterMWDiv="W","selected" ,"") %> >��Ź</option>
				<option value="X" <%= CHKIIF(lastCenterMWDiv="X","selected" ,"") %> >��Ÿ</option>
        	</select>
			&nbsp;
	    	<font color="#CC3333">���Ա���(�԰�ó):</font>
	        <select name="lastmwdiv" class="select">
				<option value="" <%= CHKIIF(lastmwdiv="","selected" ,"") %> >��ü</option>
				<option value="M" <%= CHKIIF(lastmwdiv="M","selected" ,"") %> >����</option>
				<option value="W" <%= CHKIIF(lastmwdiv="W","selected" ,"") %> >��Ź</option>
				<option value="X" <%= CHKIIF(lastmwdiv="X","selected" ,"") %> >��Ÿ</option>
        	</select>
			&nbsp;
	    	<font color="#CC3333">�԰���:</font>
	        <select name="IpgoType" class="select">
				<option value="" <%= CHKIIF(IpgoType="","selected" ,"") %> >��ü</option>
				<option value="shopchulgo" <%= CHKIIF(IpgoType="shopchulgo","selected" ,"") %> >����-&gt;����</option>
				<option value="shopipgo" <%= CHKIIF(IpgoType="shopipgo","selected" ,"") %> >���� ���԰�</option>
        	</select>
			<!--
			&nbsp;
	    	<font color="#CC3333">��ǰ����(�ۼ���):</font>
	        <select name="itemMWdiv" class="select">
				<option value="" <%= CHKIIF(itemMWdiv="","selected" ,"") %> >��ü</option>
				<option value="M" <%= CHKIIF(itemMWdiv="M","selected" ,"") %> >����</option>
				<option value="W" <%= CHKIIF(itemMWdiv="W","selected" ,"") %> >��Ź</option>
        	</select>
			-->
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
	    	<font color="#CC3333">��ǰ����:</font>
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;
			<font color="#CC3333">��ǰ�ڵ�:</font>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="8">
			&nbsp;
			<font color="#CC3333">�귣��:</font>
			<input type="text" class="text" name="lastmakerid" value="<%= lastmakerid %>" size="20">
			&nbsp;
			<font color="#CC3333">����:</font>
			<input type="text" class="text" name="shopid" value="<%= shopid %>" size="20">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
	    	<font color="#CC3333">�߰�����:</font>
        	<input type="radio" name="addInfoType" value="none" <% if (addInfoType = "none") then %>checked<% end if %> > ����
			<input type="radio" name="addInfoType" value="sum" <% if (addInfoType = "sum") then %>checked<% end if %> > �հ�
			<input type="radio" name="addInfoType" value="sumshop" <% if (addInfoType = "sumshop") then %>checked<% end if %> > �հ�(����)
			<input type="radio" name="addInfoType" value="diff" <% if (addInfoType = "diff") then %>checked<% end if %> > �԰���-������� ���� ����ġ
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p>

	<h5>�۾���...</h5>
	* �����԰� ��ǰ�� ���� �� ������ �ý������ �ջ��Ͽ� ����մϴ�.(���Ա����� ������ ���)<br>
	* ���Ա����� �ٸ��ų� �����԰� ��ǰ�� �ƴ� ��� ���庰�� ��ո��԰��� ���˴ϴ�.<br>
	* ����-���� ���Ա����� ������ ��� �����԰����� <font color="red">�԰��� ���� ���԰� ��� ������ո��԰�</font>�� ����Ͽ� ���԰��� ����մϴ�.<br><br>
	* ����-���� �̵������� <font color="red">���������� �ۼ��� ���Ŀ�</font> ��ȸ�����մϴ�.<br><br>
	* <font color="red">������� �����</font>�� �����Ǿ� ���� �ʴ� ��ǰ�� ���ܵ˴ϴ�.<br />
	* ����� ������ ������, ����-���� ���Ա����� ��� �����̰�, ����������� ��� �������� ������� �����մϴ�.<br />
	* ���͸��Ա��� �����̰�, ����� �����̰�, ����������̸� ������Ա����� ������ �������� �����մϴ�.

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="0.����������(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsSetNotAssignedMWDiv();">
			<input type="button" class="button" value="1.�����԰�(<%= yyyy1 %>-<%= mm1 %>)" onclick="fnUpdateIpgoList('L');">
			<input type="button" class="button" value="2.������ո��԰�(<%= yyyy1 %>-<%= mm1 %>)" onclick="fnUpdateavgIpgoPrice('L');">
			&nbsp;
			<input type="button" class="button" value="3.�����԰�(<%= yyyy1 %>-<%= mm1 %>) " onclick="fnUpdateIpgoList('S');">
			<input type="button" class="button" value="4.������ո��԰�(<%= yyyy1 %>-<%= mm1 %>)" onclick="fnUpdateavgIpgoPrice('S');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<% if ((addInfoType = "sum") or (addInfoType = "sumshop") or (addInfoType = "diff")) and ojaego.FResultCount > 0 then %>
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60">�԰��</td>
		<td width="40">�԰�ó</td>
		<td width="100">����</td>
		<td width="40">�μ�</td>
		<td width="120">�귣��<br>(����ڻ�)</td>
		<td width="30">����</td>
		<td width=70>��ǰ�ڵ�</td>
		<td width=40>�ɼ�</td>
		<td width="80">�԰�<br>����</td>
		<td width="60">���Ա���<br>(����)</td>
		<td width="60">���Ա���<br>(�԰�ó)</td>
		<td width="60">��������</td>
		<td width="50"><b>�հ�<br>����</b></td>
		<td width="80"><b>���԰�<br>�հ�</b></td>
		<td width="80">����<br>��ո��԰�</td>
		<td width="150">�԰���</td>
		<td width="150">�ۼ���</td>
		<td>���</td>
	</tr>
<% if ojaegosum.FResultCount >0 then %>
	<% for i=0 to ojaegosum.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center><%= ojaegosum.FItemList(i).Fyyyymm %></td>
		<td align=center><%= ojaegosum.FItemList(i).GetStockPlaceName %></td>
		<td align=center><%= ojaegosum.FItemList(i).Fshopid %></td>
		<td align=center><%= ojaegosum.FItemList(i).FtargetGbn %></td>
		<td align=center><%= ojaegosum.FItemList(i).Flastmakerid %></td>
		<td align=center><%= ojaegosum.FItemList(i).Fitemgubun %></td>
		<td align="right"><%= ojaegosum.FItemList(i).Fitemid %></td>
		<td align=center><%= ojaegosum.FItemList(i).Fitemoption %></td>
		<td align=center><%= ojaegosum.FItemList(i).GetIpgoMWdivName %></td>
		<td align=center><%= ojaegosum.FItemList(i).FlastCenterMWDiv %></td>
		<td align=center><%= ojaegosum.FItemList(i).Flastmwdiv %></td>
		<td align=center><%= ojaegosum.FItemList(i).Flastvatinclude %></td>
		<td align="right">
			<%= FormatNumber(ojaegosum.FItemList(i).FtotItemNo, 0) %>
		</td>
		<td align="right">
			<%= FormatNumber(ojaegosum.FItemList(i).FtotBuyCash, 0) %>
		</td>
		<td align="right">
			<% if (ojaegosum.FItemList(i).FtotItemNo = 0) then %>
			0
			<% else %>
			<%= FormatNumber(ojaegosum.FItemList(i).FtotBuyCash/ojaegosum.FItemList(i).FtotItemNo, 0) %>
			<% end if %>
		</td>
		<td align=center></td>
		<td align=center></td>
		<td>
			<% if (addInfoType = "diff") then %>
			<%= ojaegosum.FItemList(i).FtotItemNo - ojaegosum.FItemList(i).FstockIpgoNo %>
			<% end if %>
	    </td>
	</tr>
	<%
	totItemNoSUM = totItemNoSUM + ojaegosum.FItemList(i).FtotItemNo
	totBuyCashSUM = totBuyCashSUM + ojaegosum.FItemList(i).FtotBuyCash
	%>
	<% next %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align="right"></td>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align=center></td>
		<td align="right">
			<%= FormatNumber(totItemNoSUM, 0) %>
		</td>
		<td align="right">
			<%= FormatNumber(totBuyCashSUM, 0) %>
		</td>
		<td align="right"></td>
		<td align=center></td>
		<td align=center></td>
		<td></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF" height=50>
		<td align=center colspan="17">������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>
<% end if %>

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60">�԰��</td>
		<td width="40">�԰�ó</td>
		<td width="100">����</td>
		<td width="40">�μ�</td>
		<td width="120">�귣��<br>(����ڻ�)</td>
		<td width="30">����</td>
		<td width=70>��ǰ�ڵ�</td>
		<td width=40>�ɼ�</td>
		<td width="80">�԰�<br>����</td>
		<td width="60">���Ա���<br>(����)</td>
		<td width="60">���Ա���<br>(�԰�ó)</td>
		<td width="60">��������</td>
		<td width="50"><b>�հ�<br>����</b></td>
		<td width="80"><b>���԰�<br>�հ�</b></td>
		<td width="80">����<br>��ո��԰�</td>
		<td width="150">�԰���</td>
		<td width="150">�ۼ���</td>
		<td>���</td>
	</tr>
	<% if ojaego.FResultCount >0 then %>
	<% for i=0 to ojaego.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center><%= ojaego.FItemList(i).Fyyyymm %></td>
		<td align=center><%= ojaego.FItemList(i).GetStockPlaceName %></td>
		<td align=center><%= ojaego.FItemList(i).Fshopid %></td>
		<td align=center><%= ojaego.FItemList(i).FtargetGbn %></td>
		<td align=center><%= ojaego.FItemList(i).Flastmakerid %></td>
		<td align=center><a href="javascript:popIpgoDetailList('<%= ojaego.FItemList(i).Fyyyymm %>', '<%= ojaego.FItemList(i).FipgoType %>', '<%= ojaego.FItemList(i).Fitemgubun %>', <%= ojaego.FItemList(i).Fitemid %>)"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
		<td align="right">
			<a href="javascript:popIpgoDetailList('<%= ojaego.FItemList(i).Fyyyymm %>', '<%= ojaego.FItemList(i).FipgoType %>', '<%= ojaego.FItemList(i).Fitemgubun %>', <%= ojaego.FItemList(i).Fitemid %>)"><%= ojaego.FItemList(i).Fitemid %></a>
		</td>
		<td align=center><%= ojaego.FItemList(i).Fitemoption %></td>
		<td align=center><%= ojaego.FItemList(i).GetIpgoMWdivName %></td>
		<td align=center><%= ojaego.FItemList(i).FlastCenterMWDiv %></td>
		<td align=center><%= ojaego.FItemList(i).Flastmwdiv %></td>
		<td align=center><%= ojaego.FItemList(i).Flastvatinclude %></td>
		<td align="right">
			<%= FormatNumber(ojaego.FItemList(i).FtotItemNo, 0) %>
		</td>
		<td align="right">
			<%= FormatNumber(ojaego.FItemList(i).FtotBuyCash, 0) %>
		</td>
		<td align="right">
			<% if (ojaego.FItemList(i).FtotItemNo = 0) then %>
			0
			<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FtotBuyCash/ojaego.FItemList(i).FtotItemNo, 0) %>
			<% end if %>

		</td>
		<td align=center><%= ojaego.FItemList(i).GetIpgoTypeName %></td>
		<td align=center><%= ojaego.FItemList(i).Flastupdate %></td>
		<td>

	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan=18 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18" align="center">
			<% if ojaego.HasPreScroll then %>
        		<a href="javascript:NextPage('<%= ojaego.StarScrollPage-1 %>')">[pre]</a>
        	<% else %>
        		[pre]
        	<% end if %>

        	<% for i=0 + ojaego.StarScrollPage to ojaego.FScrollCount + ojaego.StarScrollPage - 1 %>
        		<% if i>ojaego.FTotalpage then Exit for %>
        		<% if CStr(page)=CStr(i) then %>
        		<font color="red">[<%= i %>]</font>
        		<% else %>
        		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
        		<% end if %>
        	<% next %>

        	<% if ojaego.HasNextScroll then %>
        		<a href="javascript:NextPage('<%= i %>')">[next]</a>
        	<% else %>
        		[next]
        	<% end if %>
		</td>
	</tr>
</table>
<%
set ojaego = Nothing
set ojaegosum = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
