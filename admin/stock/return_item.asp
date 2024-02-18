<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%
dim page
dim makerid, itemid, sellyn, isusing, realstocknotzero


page    = request("page")
makerid = request("makerid")
itemid  = request("itemid")
sellyn  = request("sellyn")
isusing = request("isusing")
realstocknotzero = request("realstocknotzero")

if ((request("research") = "") and (isusing = "")) then
    isusing = "off"
end if

if ((request("research") = "") and (realstocknotzero = "")) then
    realstocknotzero = "on"
end if

if page="" then page=1

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FCurrPage=page
osummarystock.FPageSize=100
osummarystock.FRectMakerid = makerid
osummarystock.FRectItemID = itemid
osummarystock.FRectOnlyIsUsing = isusing
osummarystock.FRectrealstocknotzero = realstocknotzero

if (makerid<>"") then
    osummarystock.GetCurrentStockByOnlineBrandDanjong
else
    osummarystock.FPageSize=1000
    osummarystock.GetCurrentStockByOnlineBrandDanjong_GroupBrand
end if

dim i, ttlitemno

%>


<script language='javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function changecontent(){
	// nothing
}

function Research(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SearchByBrand(makerid){
    var frm = document.frm;
    frm.makerid.value = makerid;
	frm.submit();
}

function PopReturnItemByBrand(imakerid){
    var params = "menupos=" + frm.menupos.value + "&makerid=" + imakerid
    if (frm.isusing.checked==true){
        params = params + "&isusing=" + frm.isusing.value;
    }

    if (frm.realstocknotzero.checked==true){
        params = params + "&realstocknotzero=" + frm.realstocknotzero.value;
    }

    var popwin = window.open('/admin/stock/return_item.asp?' + params,'PopReturnItemByBrand','width=900, height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
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
        	��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
        	&nbsp;
        	<input type=checkbox name="isusing" value="on" <% if isusing="on" then response.write "checked" %> >����ǰ��
        	&nbsp;
        	<input type=checkbox name="realstocknotzero" value="on" <% if realstocknotzero="on" then response.write "checked" %> >�ǻ���� 0�� �ƴ� ��ǰ
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= FormatNumber(osummarystock.FTotalCount,0) %></b>
			&nbsp;
			������ :
			<% if osummarystock.FCurrPage > 1  then %>
				<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
			<% end if %>
			<b><%= page %> / <%= osummarystock.FTotalpage %></b>
			<% if (osummarystock.FTotalpage - osummarystock.FCurrPage)>0  then %>
				<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
			<% end if %>
		</td>
	</tr>
<% if makerid<>"" then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">��ǰ<br>�ڵ�</td>
		<td width="50">�̹���</td>
		<td width="70">�귣��</td>
		<td>��ǰ��<br>[�ɼǸ�]</td>
		<td width="30">���<br>����</td>
        <td width="30">��ü<br>�԰�<br>��ǰ</td>
        <td width="30">��ü<br>�Ǹ�<br>��ǰ</td>
        <td width="30">��ü<br>���<br>��ǰ</td>
        <td width="30">��Ÿ<br>���<br>��ǰ</td>
<!--    <td width="35"><b>�ý���<br>���</b></td>	-->
		<td width="30">��<br>�ҷ�</td>
<!--    <td width="35"><b>��ȿ<br>���</b></td>	-->
        <td width="30">��<br>�ǻ�<br>����</td>
        <td width="30"><b>�ǻ�<br>���</b></td>
        <td width="30">��<br>��ǰ<br>�غ�</td>
        <td width="30">��<br>�ֹ�<br>����</td>
        <td width="30"><b>����<br>���</b></td>

		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="30">����<br>����</td>
		<td width="30">����<br>����</td>
<!--	<td width="30">ǰ��<br>����</td>	-->
    </tr>
<% for i=0 to osummarystock.FresultCount-1 %>
	<% if osummarystock.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<td><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left"><a href="javascript:SearchByBrand('<%= osummarystock.FItemList(i).FMakerID %>');"><%= osummarystock.FItemList(i).FMakerID %></a></td>
		<td align="left">
			<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
			<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
			<br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
        </td>
        <td><%= osummarystock.FItemList(i).GetMwDivName %></td>
		<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
		<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
		<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
        <td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>
<!--    <td><%= osummarystock.FItemList(i).Ftotsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
<!--    <td><%= osummarystock.FItemList(i).Favailsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
        <td><b><%= osummarystock.FItemList(i).Frealstock %></b></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv4 + osummarystock.FItemList(i).Fipkumdiv2 + osummarystock.FItemList(i).Foffjupno %></td>
        <td><b><%= osummarystock.FItemList(i).GetMaystock %></b></td>

        <td><font color="<%= ynColor(osummarystock.FItemList(i).Fsellyn) %>"><%= osummarystock.FItemList(i).Fsellyn %></font></td>
        <td><font color="<%= ynColor(osummarystock.FItemList(i).Fisusing) %>"><%= osummarystock.FItemList(i).Fisusing %></font></td>
        <td>
        	<font color="<%= ynColor(osummarystock.FItemList(i).Flimityn) %>"><%= osummarystock.FItemList(i).Flimityn %>
			<% if (osummarystock.FItemList(i).Flimityn = "Y") then %>
				<br>
				(<%= osummarystock.FItemList(i).GetLimitStr %>)
			<% end if %>
			</font>
        </td>
        <td>
            <% if osummarystock.FItemList(i).FDanjongyn="Y" then %>
            <font color="#33CC33">����</font>
            <% elseif osummarystock.FItemList(i).FDanjongyn="M" then %>
            <font color="#33CC33">MD<br>ǰ��</font>
            <% elseif osummarystock.FItemList(i).FDanjongyn="S" then %>
            <font color="#33CC33">�Ͻ�<br>ǰ��</font>
            <% else %>
            <% end if %>
        </td>
<!--    <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">ǰ��</font><% end if %></td>	-->
	</tr>
	</form>
<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if osummarystock.HasPreScroll then %>
					<a href="javascript:GotoPage(<%= osummarystock.StartScrollPage-1 %>)">[pre]</a>
			<% else %>
					[pre]
			<% end if %>

			<% for i=0 + osummarystock.StartScrollPage to osummarystock.FScrollCount + osummarystock.StartScrollPage - 1 %>
			        <% if i>osummarystock.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
					<font color="red">[<%= i %>]</font>
				<% else %>
					<a href="javascript:GotoPage(<%= i %>)">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if osummarystock.HasNextScroll then %>
					<a href="javascript:GotoPage(<%= i %>)">[next]</a>
			<% else %>
					[next]
			<% end if %>
		</td>
	</tr>
</table>


<% else %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">�귣��</td>
		<td width="100">����ǰ��</td>
		<td >&nbsp;</td>
	</tr>
	<% for i= 0 to osummarystock.FResultCount-1 %>
	<%
	ttlitemno = ttlitemno + osummarystock.FItemList(i).FCnt
    %>
	<tr bgcolor="#FFFFFF" >
	    <td><a href="javascript:PopReturnItemByBrand('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).FCnt,0) %></td>
	    <td align="right"><!-- <img src="/images/icon_detail.gif"> --></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" >
	    <td></td>
	    <td align="center"><%= FormatNumber(ttlitemno,0) %></td>
	    <td></td>
	</tr>
</table>
<% end if %>


<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
