<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ��ǰ ��� ��� ��ǰ 
' Hieditor : 2010.10.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/waitDIYitemCls.asp"-->

<%
Dim owaititem,ix,page ,sorttype, sortkey, sortkeyMid, currstate
	page = RequestCheckvar(request("page"),10)
	currstate = RequestCheckvar(request("currstate"),10)
	sorttype  = RequestCheckvar(request("sorttype"),10)
	sortkey = RequestCheckvar(request("sortkey"),32)
	sortkeyMid = RequestCheckvar(request("sortkeyMid"),10)
	
	if (page="") then page=1
	
	if sorttype="" then sorttype="C"
	if currstate="" then currstate="W"

set owaititem = new CWaitItemlist
	owaititem.FPageSize = 30
	owaititem.FCurrPage = page
	owaititem.FRectsortkey = sortkey
	owaititem.FRectsortkeyMid = sortkeyMid
	owaititem.FRectCurrState = currstate
	
	if sorttype="C" then
		owaititem.getWaitProductListByCategory
	elseif sorttype="B" then
		owaititem.getWaitProductListByBrand
	end if
%>

<script language='javascript'>

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function ViewItemDetail(itemno){
	window.open('/academy/itemmaster/viewDIYitem/viewDIYitem.asp?itemid='+itemno ,'window1','width=1024,height=960,scrollbars=yes,status=no');
}

function insertdb(itemid,itemname){
 //if (confirm(itemname + "�� ����Ͻðڽ��ϱ�?") == true){
    //location.href("item_insertdb.asp?itemid="+itemid);
 //}
}

function WaitState(itemid){
	var ret = confirm('��ϴ��� �����Ͻðڽ��ϱ�?');

	if (ret){
		document.location = 'doitemregboru.asp?mode=waitstate&idx=' + itemid;
	}
}

function popItemModify(itemid,designer){
	var popwin = window.open('wait_diyitem_modify.asp?itemid=' + itemid + '&designer=' + designer +'&fingerson=on','waititemmodify','width=860,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="sorttype" value="<%= sorttype %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<input type="radio" name="currstate" value="W" <% if currstate="W" then response.write "checked" %>>��ϴ���ǰ��
		<input type="radio" name="currstate" value="WR" <% if currstate="WR" then response.write "checked" %>>��ϴ��+��Ϻ���
		<input type="radio" name="currstate" value="A" <% if currstate="A" then response.write "checked" %>>��ü
		&nbsp;
		<% if sorttype="C" then %>
			ī�װ� :
			<% DrawSelectBoxCategoryLarge "sortkey" , sortkey %>&nbsp;
			<% DrawSelectBoxCategoryMid "sortkeyMid" , sortkey, sortkeyMid %>
		<% else %>
			�귣�� :
			<% drawSelectBoxLecturer "sortkey" , sortkey %>
		<% end if %>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">			
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<br>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if owaititem.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= owaititem.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= owaititem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">No.</td>
	<td align="center">��ǰ��</td>
	<td align="center">�̸�����</td>
	<td align="center">�ǸŰ�</td>
	<td align="center">���ް�</td>
	<td align="center">����</td>
	<td align="center">�����̳�</td>
	<td align="center">�����</td>
	<td align="center">����</td>
</tr>
<% for ix=0 to owaititem.FresultCount-1 %>

<tr align="center" bgcolor="#FFFFFF" >
	<td align="center"><%= owaititem.FItemList(ix).Fitemid %></td>
	<td align="left"><a href="javascript:popItemModify('<% =owaititem.FItemList(ix).Fitemid %>','<%= owaititem.FItemList(ix).Fmakerid %>')"><%= owaititem.FItemList(ix).Fitemname %></a></td>
	<td align="center">
		<% if owaititem.FItemList(ix).FCurrState="7" then %>
			<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<% =owaititem.FItemList(ix).Flinkitemid %>" target="_blank"><font color="blue">(����)</font></a>
		<% else %>
			<a href="javascript:ViewItemDetail('<% =owaititem.FItemList(ix).Fitemid %>')"><font color="blue">(�̸�����)</font></a>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(owaititem.FItemList(ix).Fsellcash,0) %></td>
	<td align="right"><%= FormatNumber(owaititem.FItemList(ix).Fsuplycash,0) %></td>
	<td align="center">
	<% if owaititem.FItemList(ix).Fsellcash<>0 then %>
	<%= 100 - CLng(owaititem.FItemList(ix).Fsuplycash/owaititem.FItemList(ix).Fsellcash*100*100)/100 %> %
	<% end if %>
	</td>
	<td align="center"><%= owaititem.FItemList(ix).Fmakerid %></td>
	<td align="center"><%= owaititem.FItemList(ix).Fregdate %></td>
	<td align="center"><font color="<%= owaititem.FItemList(ix).GetCurrStateColor %>"><%= owaititem.FItemList(ix).GetCurrStateName %></font>
	<% if (owaititem.FItemList(ix).FCurrState="2") or (owaititem.FItemList(ix).FCurrState="0") then %>
	<a href="javascript:WaitState('<%= owaititem.FItemList(ix).Fitemid %>')"><br><font color="#000000">[��ϴ�⺯��]</font></a>
	<% end if %>
	</td>
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if owaititem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= owaititem.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + owaititem.StarScrollPage to owaititem.StarScrollPage + owaititem.FScrollCount - 1 %>
			<% if (ix > owaititem.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(owaititem.FCurrPage) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>
	
		<% if owaititem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<%
	set owaititem = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->