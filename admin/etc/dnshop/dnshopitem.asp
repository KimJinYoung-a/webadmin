<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim itemid, itemname, eventid, mode, delitemid, deljaehyu
dim itemidArr, eventidArr, makeridArr
dim page, makerid

page    = request("page")
itemid  = request("itemid")
delitemid = requestCheckvar(request("delitemid"),9)
itemname= request("itemname")
eventid = request("eventid")
mode    = request("mode")
itemidArr = Trim(request("itemidArr"))
eventidArr= Trim(request("eventidArr"))
makeridArr = Trim(request("makeridArr"))
makerid= request("makerid")
deljaehyu = request("deljaehyu")
 
if page="" then page=1
if Right(itemidArr,1)="," then itemidArr=Left(itemidArr,Len(itemidArr)-1)
if Right(eventidArr,1)="," then eventidArr=Left(eventidArr,Len(eventidArr)-1)


dim sqlStr, resultRow
if (mode="regByItemIDarr") then
    sqlStr = "insert into [db_item].[dbo].tbl_dnshop_reg_item " + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid) " + VbCrlf
    sqlStr = sqlStr + " select top 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + "     left join  [db_item].[dbo].tbl_dnshop_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='080' )" + VbCrlf ''and  i.cate_mid in ('010')
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='090' )" + VbCrlf ''and  i.cate_mid in ('010')
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','060','070','110'))" + VbCrlf
    sqlStr = sqlStr + " )"
    
    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.itemid in (" + itemidArr + ")" + VbCrlf
    sqlStr = sqlStr + " and i.sellyn='Y'"
    ''sqlStr = sqlStr + " and ((i.limityn='N') or (i.limityn='Y' and i.limitno-i.limitsold>50))" + VbCrlf
    ''sqlStr = sqlStr + " and recentsellcount>=3" + VbCrlf
    sqlStr = sqlStr + " and sellcash<>0" + VbCrlf
    sqlStr = sqlStr + " and ((sellcash-buycash)/sellcash)*100>=15" + VbCrlf
    
    ''��ü ������۵� ����
    sqlStr = sqlStr + " and i.makerid not in (select userid from [db_user].[dbo].tbl_user_c where defaultDeliveryType is not NULL)"
    
    ''���� ����Ŀ ����
    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun = 'dnshop')" + VbCrlf
    
    ''�ɼ� �߰��ݾ� �ִ°� ����
    sqlStr = sqlStr + " and i.itemid not in (select distinct itemid from db_item.dbo.tbl_item_option where optaddprice>0)"
    
    ''2009 ���̾ ����
    sqlStr = sqlStr + " and i.itemid not in (select  itemid from db_diary2010.dbo.tbl_diaryMaster)"
    
    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'vanillaspoon'" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'kongkkakji'" + VbCrlf
    
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"
    
''response.write sqlStr
    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
elseif (mode="regByEventIDarr") then
    
    sqlStr = "insert into [db_item].[dbo].tbl_dnshop_reg_item" + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid)" + VbCrlf
    sqlStr = sqlStr + " select top 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_event].[dbo].tbl_eventitem e," + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + " left join  [db_item].[dbo].tbl_dnshop_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where e.evt_code in (" + eventidArr + ")" + VbCrlf
    sqlStr = sqlStr + " and e.itemid=i.itemid" + VbCrlf
    sqlStr = sqlStr + " and (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='080')" + VbCrlf ' and  i.cate_mid in ('010')
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='090' )" + VbCrlf 'and  i.cate_mid in ('010')
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','060','070','110'))" + VbCrlf
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.sellcash<>0" + VbCrlf
    sqlStr = sqlStr + " and (( i.sellcash- i.buycash)/ i.sellcash)*100>=15" + VbCrlf
    
    ''��ü ������۵� ����
    sqlStr = sqlStr + " and i.makerid not in (select userid from [db_user].[dbo].tbl_user_c where defaultDeliveryType is not NULL)"
    
    ''���� ����Ŀ ����
    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun = 'dnshop')" + VbCrlf
    
    ''�ɼ� �߰��ݾ� �ִ°� ����
    sqlStr = sqlStr + " and i.itemid not in (select distinct itemid from db_item.dbo.tbl_item_option where optaddprice>0)"
    
    ''2009 ���̾ ����
    sqlStr = sqlStr + " and i.itemid not in (select  itemid from db_diary2010.dbo.tbl_diaryMaster)"
    
    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'vanillaspoon'" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'kongkkakji'" + VbCrlf
    
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"
    
    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
elseif (mode="recentBestSeller") then
    sqlStr = "insert into [db_item].[dbo].tbl_dnshop_reg_item" + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid)" + VbCrlf
    sqlStr = sqlStr + " select top 100 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item_contents c, [db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + " left join  [db_item].[dbo].tbl_dnshop_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where i.itemid=c.itemid" + VbCrlf
    sqlStr = sqlStr + " and (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='080' and  i.cate_mid in ('010'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='090' and  i.cate_mid in ('010'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='110' and  i.cate_mid in ('010','020','030'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
	sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','060','070','110'))" + VbCrlf
    sqlStr = sqlStr + " )"

''    sqlStr = sqlStr + " where i.cate_large" + VbCrlf
''    sqlStr = sqlStr + " in (" + VbCrlf
''    sqlStr = sqlStr + " '10','15','25','40','20'" + VbCrlf
''    sqlStr = sqlStr + " )" + VbCrlf
    
    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.sellyn='Y'" + VbCrlf
''    sqlStr = sqlStr + " and i.dispyn='Y'" + VbCrlf
    sqlStr = sqlStr + " and ((i.limityn='N') or (i.limityn='Y' and i.limitno-i.limitsold>=30))" + VbCrlf
    sqlStr = sqlStr + " and c.recentsellcount>=5" + VbCrlf
    sqlStr = sqlStr + " and sellcount>5" + VbCrlf
    sqlStr = sqlStr + " and sellcash<>0" + VbCrlf
    sqlStr = sqlStr + " and ((sellcash-buycash)/sellcash)*100>=20" + VbCrlf
    
    ''��ü ������۵� ����
    sqlStr = sqlStr + " and i.makerid not in (select userid from [db_user].[dbo].tbl_user_c where defaultDeliveryType is not NULL)"
    
    ''���� ����Ŀ ����
    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun = 'dnshop')" + VbCrlf
    
    ''�ɼ� �߰��ݾ� �ִ°� ����
    sqlStr = sqlStr + " and i.itemid not in (select distinct itemid from db_item.dbo.tbl_item_option where optaddprice>0)"
    
    ''2009 ���̾ ����
    sqlStr = sqlStr + " and i.itemid not in (select  itemid from db_diary2010.dbo.tbl_diaryMaster)"
    
    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'vanillaspoon'" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'kongkkakji'" + VbCrlf
    
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"
    
''response.write "������"    
    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
elseif (mode="regByMakerid") then
    sqlStr = "insert into [db_item].[dbo].tbl_dnshop_reg_item " + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid) " + VbCrlf
    sqlStr = sqlStr + " select top 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + "     left join  [db_item].[dbo].tbl_dnshop_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='080' )" + VbCrlf ''and  i.cate_mid in ('010')
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='090' )" + VbCrlf ''and  i.cate_mid in ('010')
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','060','070','110'))" + VbCrlf
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " and i.sellyn='Y'" + VbCrlf
    sqlStr = sqlStr + " and i.makerid ='" & makeridArr & "'" + VbCrlf
    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    ''sqlStr = sqlStr + " and i.itemid in (" + itemidArr + ")" + VbCrlf
    ''sqlStr = sqlStr + " and ((i.limityn='N') or (i.limityn='Y' and i.limitno-i.limitsold>50))" + VbCrlf
    ''sqlStr = sqlStr + " and recentsellcount>=3" + VbCrlf
    sqlStr = sqlStr + " and sellcash<>0" + VbCrlf
    sqlStr = sqlStr + " and ((sellcash-buycash)/sellcash)*100>=15" + VbCrlf
    
    ''��ü ������۵� ����
    sqlStr = sqlStr + " and i.makerid not in (select userid from [db_user].[dbo].tbl_user_c where defaultDeliveryType is not NULL)"
    
    ''���� ����Ŀ ����
    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun = 'dnshop')" + VbCrlf
    
    ''�ɼ� �߰��ݾ� �ִ°� ����
    sqlStr = sqlStr + " and i.itemid not in (select distinct itemid from db_item.dbo.tbl_item_option where optaddprice>0)"
    
    ''2009 ���̾ ����
    sqlStr = sqlStr + " and i.itemid not in (select  itemid from db_diary2010.dbo.tbl_diaryMaster)"
    
    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'vanillaspoon'" + VbCrlf
    sqlStr = sqlStr + " and i.makerid<>'kongkkakji'" + VbCrlf
    
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"
    
''response.write sqlStr
    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
elseif (mode = "delitem") then
	dbget.Execute "Delete [db_item].[dbo].tbl_dnshop_reg_item Where itemid = '" & delitemid & "'"
    response.write "<script >alert('�����Ǿ����ϴ�.')</script>"
end if

dim odnshopitem
set odnshopitem = new CExtSiteItem
odnshopitem.FPageSize       = 20
odnshopitem.FCurrPage       = page
odnshopitem.FRectItemID     = itemid
odnshopitem.FRectItemName   = itemname
odnshopitem.FRectEventid    = eventid
odnshopitem.FRectMakerid    = makerid
odnshopitem.FDelJaeHyu		= deljaehyu
odnshopitem.GetDnshopRegedItemList

dim i
%>
<script language='javascript'>
function goPage(page){
    frm.page.value = page;
    frm.submit();
}

function RegByItemID(frm){
    if (frm.itemidArr.value.length<1){
        alert('��ǰ��ȣ�� �Է��� �ּ���.');
        frm.itemidArr.focus();
        return;
    }
    
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "regByItemIDarr";
        frm.submit();
    }
}

function RegByEventID(frm){
    if (frm.eventidArr.value.length<1){
        alert('�̺�Ʈ ��ȣ��  �Է��� �ּ���.');
        frm.eventidArr.focus();
        return;
    }
    
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "regByEventIDarr";
        frm.submit();
    }
}
function RegByRecentSell(frm){
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "recentBestSeller";
        frm.submit();
    }
}

function RegByMakerID(frm){
    if (frm.makeridArr.value.length<1){
        alert('�귣�� ID��  �Է��� �ּ���.');
        frm.makeridArr.focus();
        return;
    }
    
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "regByMakerid";
        frm.submit();
    }
}

function DelItem(code)
{
	frm.delitemid.value = code;
    frm.mode.value = "delitem";
    frm.submit();
}

function NotInMakerid()
{
	window.open('/admin/etc/JaehyuMall_Not_In_Makerid.asp?mallgubun=dnshop','notin','width=200,height=400,scrollbars=yes');
}

function deleteitem()
{
	if(confirm("�˻��� �ϼ̽��ϱ�?\n�˻��� ���������� ������ ������ �˴ϴ�.") == true) {
		window.open('pop_dnshopitem.asp?makerid=<%=makerid%>&itemid=<%=itemid%>&itemname=<%=itemname%>&eventid=<%=eventid%>&deljaehyu=<%=deljaehyu%>','deleteitem','width=350,height=400,scrollbars=yes');
	} else {
		return false;
	}
}

function category_manager()
{
	window.open('DnshopCategory.asp','category_manager','width=1100,height=700,scrollbars=yes');
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="delitemid" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
		<td class="a">
		�귣�� :
		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		��ǰ��ȣ:
		<input type="text" name="itemid" value="<%= itemid %>" size="20" class="input">
		��ǰ��:
		<input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="input">
		�̺�Ʈ��ȣ:
		<input type="text" name="eventid" value="<%= eventid %>" size="6" maxlength="6" class="input">
		&nbsp;
		<input type="checkbox" name="deljaehyu" value="o" <% If deljaehyu = "o" Then %>checked<% End If %>>���޸��������ΰ�
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	<tr >
		<td colspan="2" align="right" class="a"><input class="button" type="button" value="�˻���� ��ǰ ����" onclick="deleteitem();"></td>
	</tr>
	</form>
</table>

<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmReg" method="post" action="dnshopitem.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr height="30" bgcolor="#FFFFFF">
    <td>
        ��ǰ�ڵ�� ��� &nbsp;&nbsp;&nbsp;&nbsp;: 
        <input class="input" type="input" name="itemidArr" value="" size="60"> <input class="button" type="button" value="���" onclick="RegByItemID(frmReg);">(�޸ӷ� ����)
        <br>
        �̺�Ʈ ��ȣ�� ��� : <input class="input" type="input" name="eventidArr" value="" size="60"> <input class="button" type="button" value="���" onclick="RegByEventID(frmReg);">(�޸ӷ� ����)
        <br>
        �귣��ID�� ��� &nbsp;&nbsp;&nbsp;&nbsp;: 
        <input class="input" type="text" name="makeridArr" value="" size="32" maxlength="32"> <input class="button" type="button" value="���" onclick="RegByMakerID(frmReg);">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
        <tr height="10"><td></td></tr>
        <tr>
        	<td>
        		<input class="button" type="button" value="�ֱ� ����Ʈ ���� ���" onclick="RegByRecentSell(frmReg);">
		        &nbsp;&nbsp;&nbsp;
		        <input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();">
        	</td>
        	<td align="right">
<%
'	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
%>
        	<input class="button" type="button" value="DnShopī�װ���Ī" onclick="category_manager();">
<%
'	End If
%>
        	</td>
        </tr>
        </table>
    </td>
</tr>
</form>
</table>
<br>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(odnshopitem.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(odnshopitem.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="50">�̹���</td>
	<td width="60">��ǰ��ȣ</td>   
	<td >��ǰ��</td>
	<td width="100">�����</td>
	<td width="100">�����ID</td>
	<td width="70">�ǸŰ�</td>
	<td width="70">����</td>
	<td width="70">ǰ������</td>
	<td width="70">ī�װ�����</td>
	<td width="10"></td>
</tr>
<% for i=0 to odnshopitem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
    <td><img src="<%= odnshopitem.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= odnshopitem.FItemList(i).FItemID %></td>
    <td><%= odnshopitem.FItemList(i).FItemName %></td>
    <td><%= odnshopitem.FItemList(i).FRegdate %></td>
    <td><%= odnshopitem.FItemList(i).Freguserid %></td>
    <td align="right"><%= FormatNumber(odnshopitem.FItemList(i).FSellcash,0) %></td>
    <td align="center">
        <% if odnshopitem.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-odnshopitem.FItemList(i).Fbuycash/odnshopitem.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if odnshopitem.FItemList(i).IsSoldOut then %>
        <font color="red">ǰ��</font>
        <% end if %>
    </td>
    <td align="center">
        <%= odnshopitem.FItemList(i).Fdnshopmngcategory %>
        <br><%= odnshopitem.FItemList(i).Fdnshopdispcategory  %>
        <br><%= odnshopitem.FItemList(i).Fdnshopstorecategory %>
    </td>
    <td align="center">
    <!--
    	<a href="javascript:DelItem('<%= odnshopitem.FItemList(i).FItemID %>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
    //-->
    </td>
</tr>
<% next %>
<tr height="20">
    <td colspan="10" align="center" bgcolor="#FFFFFF">
        <% if odnshopitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= odnshopitem.StarScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>
    
    	<% for i=0 + odnshopitem.StarScrollPage to odnshopitem.FScrollCount + odnshopitem.StarScrollPage - 1 %>
    		<% if i>odnshopitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>
    
    	<% if odnshopitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<%
set odnshopitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->