<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid,mode, searchtype
makerid = request("makerid")
mode = request("mode")
searchtype = request("searchtype")

searchtype = "err"

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype

if (makerid<>"") then
    osummarystock.GetDailyErrItemListByBrand
else
    osummarystock.GetDailyErrRealCheckItemListByBrandGroup
end if

dim i

%>
<script language='javascript'>

function PopErrItemLossInput(makerid){
	var popwin = window.open('/common/pop_erritem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_erritem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function SubmitSearchByBrandNew(makerid) {
	document.frm.makerid.value = makerid;
	document.frm.submit();
}

function ChangePage(v) {
	var frm = document.frm;

	if (v == "bad") {
		frm.action = "baditem_return_list.asp";
	} else {
		frm.action = "erritem_loss_list.asp";
	}

	frm.submit();
}

</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %> onClick="ChangePage('bad')" > �ҷ���ǰ
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %> onClick="ChangePage('err')"> ������ϻ�ǰ
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<% if makerid<>"" then %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
        	<input type="button" class="button" value="������ϻ�ǰ �ν�ó��" onclick="PopErrItemLossInput('<%= makerid %>')" border="0">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= osummarystock.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30">����</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="40">�ɼ�</td>
		<td width="50">�̹���</td>
    	<td width="100">�귣��ID</td>

		<td>�����۸�</td>
		<td>�ɼǸ�</td>
		<td width="40">���<br>����</td>

		<td width="50">�Һ��ڰ�</td>
		<td width="40">����<br>����</td>
    </tr>

	<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td><%= osummarystock.FItemList(i).FItemid %></td>
		<td><%= osummarystock.FItemList(i).FItemoption %></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>

		<td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
		<td><%= osummarystock.FItemList(i).GetMwDivName %></td>

		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		</td>
	</tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= osummarystock.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">�귣��</td>
		<td width="100">������ǰ��On</td>
		<td width="100">������ǰ��Off</td>
		<td >&nbsp;</td>
	</tr>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOnCnt %></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOffCnt %></td>
	    <td align="left">
        	<input type="button" class="button" value="������ϻ�ǰ �ν�ó��" onclick="PopErrItemLossInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
	    </td>
	</tr>
	<% next %>
</table>
<% end if %>

<p>




<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
