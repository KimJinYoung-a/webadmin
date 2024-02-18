<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/inspectstockcls.asp"-->
<%
dim makerid,page,isusing
dim mwdiv

makerid = request("makerid")
page = request("page")
isusing= request("isusing")
mwdiv = request("mwdiv")

if page="" then page=1
if mwdiv="" then mwdiv="T"

dim oinspectstock
set oinspectstock = New CInspectStock
oinspectstock.FPageSize = 200
oinspectstock.FCurrPage = page
oinspectstock.FRectMakerid = makerid
oinspectstock.FRectIsUsing = isusing
oinspectstock.FRectMwDiv = mwdiv
oinspectstock.GetErrRegItemList
dim i
%>
<script language='javascript'>
function reSearchBymakerid(imakerid){
    frm.makerid.value = imakerid;
    frm.submit();
}

function PopAdminSellEdit(itemgubun,itemid,itemoption){
	var popwin = window.open('popadminselledit.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption ,'popadminselledit','width=800,heght=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemStock(itemgubun,itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popitemstockdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function delStock(itemgubun,itemid, itemoption){
    if (confirm('��ϵ� ������� ���� �Ͻðڽ��ϱ�?\n(������� 0�� �ƴѰ�� �������� �ʽ��ϴ�.)')){
        var popwin = window.open('/admin/stock/delErrStock.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'delErrStock','width=100, height=100, scrollbars=yes, resizable=yes');
    	popwin.focus();
    }
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %> &nbsp;&nbsp;
	        	<br>
	        	��뿩�� :
	        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >ALL
	        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >Y
	        	
                &nbsp;&nbsp;
	        	���Ա��� :
	        	<input type="radio" name="mwdiv" value="A" <% if mwdiv="A" then response.write "checked" %> >All
	        	<input type="radio" name="mwdiv" value="T" <% if mwdiv="T" then response.write "checked" %> >�ٹ�
	        	<input type="radio" name="mwdiv" value="U" <% if mwdiv="U" then response.write "checked" %> >��ü


	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" onclick="document.frm.submit();"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        �˻���� : <%= oinspectstock.FTotalCount %> (�ִ� <%= oinspectstock.FPageSize %> ��)
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="20"></td>
    	<td width="50">ItemID</td>
    	<td width="50">ItemOption</td>
    	<td width="80">�귣��</td>
    	<td width="80">��ǰ��</td>
    	<td width="60">�ɼǸ�</td>
    	<td width="30">���</td>
    	<td width="20">�Ǹ�</td>
    	<td width="20">���</td>
    	<td width="20">�Ǹ�</td>
    	<td width="20">�԰�</td>
    	<td width="20">���</td>
    	<td >���<br>����</td>
    </tr>
<% for i=0 to oinspectstock.FresultCount-1 %>
    <tr bgcolor="#FFFFFF">
    	<td><%= oinspectstock.FItemList(i).FItemGubun %></td>
    	<td><a href="javascript:PopItemSellEdit('<%= oinspectstock.FItemList(i).FItemID %>');"><%= oinspectstock.FItemList(i).FItemID %></a></td>
    	<td><%= oinspectstock.FItemList(i).FItemOption %></td>
    	<td><a href="javascript:reSearchBymakerid('<%= oinspectstock.FItemList(i).FMakerid %>');"><%= oinspectstock.FItemList(i).FMakerid %></a></td>
    	<td><a href="javascript:PopItemStock('<%= oinspectstock.FItemList(i).FItemGubun %>','<%= oinspectstock.FItemList(i).FItemID %>','<%= oinspectstock.FItemList(i).FItemOption %>');"><%= oinspectstock.FItemList(i).FItemName %></a></td>
    	<td><%= oinspectstock.FItemList(i).FItemOptionName %></td>
    	<td><%= oinspectstock.FItemList(i).GetMwDivName %></td>
    	<td><%= oinspectstock.FItemList(i).FSellyn %></td>
    	<td><%= oinspectstock.FItemList(i).FIsusing %></td>
    	<td ><a href="javascript:PopAdminSellEdit('<%= oinspectstock.FItemList(i).FItemGubun %>','<%= oinspectstock.FItemList(i).FItemID %>','<%= oinspectstock.FItemList(i).FItemOption %>');"><%= oinspectstock.FItemList(i).Ftotsellno %></a></td>
    	<td ><a href="javascript:PopItemIpChulList('2001-01-01','<%= Left(now(),10) %>','<%= oinspectstock.FItemList(i).Fitemgubun %>','<%= oinspectstock.FItemList(i).Fitemid %>','<%= oinspectstock.FItemList(i).FItemoption %>','');"><%= oinspectstock.FItemList(i).Ftotipgono %></a></td>
    	<td ><a href="javascript:PopItemIpChulList('2001-01-01','<%= Left(now(),10) %>','<%= oinspectstock.FItemList(i).Fitemgubun %>','<%= oinspectstock.FItemList(i).Fitemid %>','<%= oinspectstock.FItemList(i).FItemoption %>','');"><%= oinspectstock.FItemList(i).Ftotchulgono %></a></td>
    	<td align="center" <%= chkIIF(oinspectstock.FItemList(i).Ftotsellno=0 and oinspectstock.FItemList(i).Ftotipgono=0 and oinspectstock.FItemList(i).Ftotchulgono=0,"bgcolor='#CC3333'","") %> ><a href="javascript:delStock('<%= oinspectstock.FItemList(i).FItemGubun %>','<%= oinspectstock.FItemList(i).FItemID %>','<%= oinspectstock.FItemList(i).FItemOption %>');"><img src="/images/icon_delete2.gif" border="0" width="20"></a></td>
    </tr>
<% next %>
</table>


<!-- ǥ �ϴܹ� ����-->
<!-- ����¡ ����.
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
	<% if oinspectstock.HasPreScroll then %>
		<a href="javascript:GotoPage('<%= oinspectstock.StarScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oinspectstock.StarScrollPage to oinspectstock.FScrollCount + oinspectstock.StarScrollPage - 1 %>
		<% if i>oinspectstock.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:GotoPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oinspectstock.HasNextScroll then %>
		<a href="javascript:GotoPage('<%= i %>')">[next]</a>
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
-->
<!-- ǥ �ϴܹ� ��-->

<%
set oinspectstock = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->