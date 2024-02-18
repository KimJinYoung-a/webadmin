<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/daumEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim allEP, page, i, makerid, itemname, itemid, onlyValidMargin
page				= request("page")
makerid				= Trim(requestCheckVar(Request("makerid"),32))
itemname   			= requestCheckvar(request("itemname"),64)
itemid				= requestCheckvar(request("itemid"),255)
onlyValidMargin		= request("onlyValidMargin")
If page = "" Then page = 1

Set allEP = new epShop
	allEP.FCurrPage				= page
	allEP.FRectMakerid			= makerid
	allEP.FRectItemname			= itemname
	allEP.FRectItemid			= itemid
	allEP.FRectOnlyValidMargin	= onlyValidMargin
	allEP.FPageSize	= 15
	allEP.AllEpItemList
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/daumEp/inc_daumHead.asp" -->
>> ��üEP����Ʈ
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		�� �� �� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		��ǰ��ȣ: <input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >���� 15%�̻� ��ǰ�� ����
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>
<br>
�ر⺻ �˻�����<br>
1.��ǰ�� �Ǹ���, �����<br>
2.��ǰ������������ ����ð����� 19���������̰ų� �ֱ��ǸŰ� 1���̻�(��üEP)<br>
3.�Ǹ����� �귣�尡 �ƴѰ�<br>
4.�Ǹ����� ��ǰ�� �ƴѰ�<br>
5.3Depth�̻� ���� ��ǰ<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(allEP.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(allEP.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>��ǰ��</td>
    <td>�귣��ID</td>
    <td>ǰ������</td>
	<td>��ǰ�����</td>
	<td>��ǰ����������</td>
	<td>�ǸŰ�</td>
	<td>����</td>
</tr>
<% For i=0 to allEP.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20" align="center">
	<td><img src="<%= allEP.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= allEP.FItemList(i).FItemid %></td>
    <td><%= allEP.FItemList(i).FItemname %></td>
    <td><%= allEP.FItemList(i).FMakerid %></td>
    <td>
        <% if allEP.FItemList(i).IsSoldOut then %>
            <% if allEP.FItemList(i).FSellyn="N" then %>
            <font color="red">ǰ��</font>
            <% else %>
            <font color="red">�Ͻ�<br>ǰ��</font>
            <% end if %>
        <% end if %>
    </td>
	<td><%= allEP.FItemList(i).FRegdate %></td>
	<td><%= allEP.FItemList(i).FLastupdate %></td>
	<td>
        <%= FormatNumber(allEP.FItemList(i).FSellcash,0) %>
	</td>
	<td>
        <% if allEP.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-allEP.FItemList(i).Fbuycash/allEP.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if allEP.HasPreScroll then %>
		<a href="javascript:goPage('<%= allEP.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + allEP.StartScrollPage to allEP.FScrollCount + allEP.StartScrollPage - 1 %>
    		<% if i>allEP.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if allEP.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->