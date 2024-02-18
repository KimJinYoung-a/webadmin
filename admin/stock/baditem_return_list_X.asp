<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid,mode, searchtype, purchasetype
makerid = request("makerid")
mode = request("mode")
searchtype = request("searchtype")
purchasetype = request("purchasetype")

if (searchtype = "") then
	searchtype = "bad"
end if


'// ===========================================================================
dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype
osummarystock.FRectPurchaseType = purchasetype


if (makerid<>"") then
    osummarystock.GetBadOrErrItemListByBrand
else
    osummarystock.GetBadOrErrItemListByBrandGroup
end if

''response.end

'// ===========================================================================
dim BadOrErrText
if (searchtype="bad") then
    BadOrErrText = "�ҷ�"
else
    BadOrErrText = "�������"
end if


dim i

%>
<script language='javascript'>
function PopBadItemReInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid,'pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopBadItemLossInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function SubmitSearchByBrandNew(makerid) {
	var frm = document.frm;

	frm.makerid.value = makerid;
	frm.submit();
}

function ChangePage(v) {
	var frm = document.frm;

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
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %> onClick="ChangePage(this)" > �ҷ���ǰ
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %> onClick="ChangePage(this)"> ������ϻ�ǰ
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<br><br>
<font size="8">�۾���</font>
<br><br>

<% if makerid<>"" then %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�ҷ���ǰ��ǰ" onclick="PopBadItemReInput('<%= makerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="�ҷ���ǰ�ν�ó��" onclick="PopBadItemLossInput('<%= makerid %>')" border="0">
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
		<td>�귣��ID</td>
		<td width="50">�̹���</td>
		<td width="50">�ŷ�<br>����</td>
		<td width="30">��ǰ<br>����</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="40">�ɼ�</td>
		<td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>

		<td width="50">�Һ��ڰ�</td>
		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="60"><%= BadOrErrText %><br>����</td>
		<td width="80">������<br>(ON+OFF)</td>
    </tr>

	<% for i=0 to osummarystock.FResultCount - 1 %>
	<% if (osummarystock.FItemList(i).Fisusing = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>
    	<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
    	<td align="center"><font color="<%= osummarystock.FItemList(i).GetMwDivColor %>"><%= osummarystock.FItemList(i).GetMwDivName %></font></td>
    	<td align="center"><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemid %></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemoption %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemname %><br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font></td>

		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fsellyn %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fisusing %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fregitemno, 0) %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fchulgoitemno, 0) %></td>
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
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td rowspan="3">�귣��</td>
		<td rowspan="3">�귣���</td>
		<td width="40" rowspan="3">�귣��<br>���<br>����</td>
		<td colspan="8"><%= BadOrErrText %>��ǰ����</td>
		<td rowspan="3">���</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td colspan="3">10</td>
		<td>70</td>
		<td>80</td>
		<td colspan="2">90</td>
		<td rowspan="2" width="80">�Ұ�</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="55">����</td>
		<td width="55">Ư��</td>
		<td width="55">����</td>
		<td width="55"></td>
		<td width="55"></td>
		<td width="55">����</td>
		<td width="55">Ư��</td>
	</tr>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<% if (osummarystock.FItemList(i).Fuseyn = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td><%= osummarystock.FItemList(i).Fmakername %></td>
	    <td align="center"><%= osummarystock.FItemList(i).Fuseyn %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem10M, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem10W, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem10U, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem70, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem80, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem90M, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem90W, 0) %></td>
	    <td align="center"><%= FormatNumber((osummarystock.FItemList(i).FOnCnt + osummarystock.FItemList(i).FOffCnt), 0) %></td>
	    <td align="left">
			<input type="button" class="button" value="�ҷ���ǰ��ǰ" onclick="PopBadItemReInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="�ҷ���ǰ�ν�ó��" onclick="PopBadItemLossInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
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
