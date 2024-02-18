<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/PlusSaleItemCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

<%
dim makerid, itemidArr, itemname, page
dim cdl, cdm, cds
'dim sellyn,usingyn

page        = RequestCheckVar(request("page"),9)
makerid     = RequestCheckVar(request("makerid"),32)
itemidArr   = RequestCheckVar(request("itemidArr"),1024)
itemname    = RequestCheckVar(request("itemname"),64)
cdl         = RequestCheckVar(request("cdl"),3)
cdm         = RequestCheckVar(request("cdm"),3)
cds         = RequestCheckVar(request("cds"),3)

if (page="") then page=1
itemidArr = Trim(itemidArr)
itemname  = Trim(itemname)
if (Right(itemidArr,1)=",") then itemidArr = Left(itemidArr,Len(itemidArr)-1)


dim oPsItemList
set oPsItemList = new CPlusSaleItem
oPsItemList.FPageSize     = 20
oPsItemList.FCurrPage     = page
oPsItemList.FRectMakerid  = makerid
oPsItemList.FRectCDL      = cdl
oPsItemList.FRectCDM      = cdm
oPsItemList.FRectCDS      = cds
oPsItemList.FRectItemIDArr= itemidArr
oPsItemList.FRectItemName = itemname

oPsItemList.GetPlusSaleMainItemList


dim i
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function PlusSaleItem_Main_New(){
    var popwin = window.open('/admin/shopmaster/PlusSaleItem_Edit.asp','PlusSaleItem_Main_New','');
    popwin.focus();
}

function showLinkedItemList(iitemid){
    var popwin = window.open('PlusSaleItem_Edit.asp?itemid=' + iitemid,'PlusSaleItem_Edit','');
    popwin.focus();
}


</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemidArr" value="<%= itemidArr %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!-- <input type="button" class="button" value="�űԵ��1" onClick="PlusSaleItem_Edit();"> -->
			<input type="button" class="button" value="�űԵ��" onClick="PlusSaleItem_Main_New();">
		</td>
		<td align="right">
		
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= oPsItemList.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oPsItemList.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">��ǰ�ڵ�</td>
    	<td width="50">�̹���</td>
      	<td width="100">�귣��ID</td>
      	<td>��ǰ��</td>
      	<td width="60">�ǸŰ�</td>
		<td width="60">���԰�</td>
		<td width="40">����</td>
		<td width="30">���<br>����</td>
      	<td width="100">��������<br>�߰�������ǰ��</td>
    </tr>
    <% for i=0 to oPsItemList.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %></td>
    	<td><img src="<%= oPsItemList.FItemList(i).Fsmallimage %>" width="50" height="50" ></td>
      	<td><%= oPsItemList.FItemList(i).FMakerid %></td>
      	<td align="left"><%= oPsItemList.FItemList(i).FitemName %></td>
      	<td align="right">
      	    <%= FormatNumber(oPsItemList.FItemList(i).FOrgPrice,0) %>
          	<% if oPsItemList.FItemList(i).IsCurrentSaleItem then %>
          		<br><font color=#F08050>(��)<%= FormatNumber(oPsItemList.FItemList(i).FSellcash,0) %></font>
          	<% end if %>
      	
      	    <% if oPsItemList.FItemList(i).IsCouponItem then %>
      	        <br><font color=#5080F0>(��)<%= FormatNumber(oPsItemList.FItemList(i).GetCouponAssignPrice,0) %></font>
      	    <% end if %>
      	</td>
      	<td align="right">
      		<%= FormatNumber(oPsItemList.FItemList(i).Forgsuplycash,0) %> 
      		<% if oPsItemList.FItemList(i).IsCurrentSaleItem then %>
      		<br><font color=#F08050>(��)<%= FormatNumber(oPsItemList.FItemList(i).FBuycash,0) %></font>
      	    <% end if %>
      	    
      	</td>
      	<td>
      		<%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %>
      		<% if oPsItemList.FItemList(i).IsCurrentSaleItem then %>
      		<br><font color=#F08050><%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %></font>
      	    <% end if %>
      	</td>
      	<td><%= fnColor(oPsItemList.FItemList(i).FMwdiv,"mw") %></td>
      	<td><a href="javascript:showLinkedItemList('<%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %>');" ><font color="red"><%= oPsItemList.FItemList(i).FPlusSaleItemCount %></font></a></td>
    </tr>
    <% next %>
    
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if oPsItemList.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oPsItemList.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    
    		<% for i=0 + oPsItemList.StarScrollPage to oPsItemList.FScrollCount + oPsItemList.StarScrollPage - 1 %>
    			<% if i>oPsItemList.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>
    
    		<% if oPsItemList.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>

<p>
<!--
*��ǰ�˻� ��, �˾�â���� �߰�������ǰ ���<br>
*<br>
-->

<%
set oPsItemList = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
