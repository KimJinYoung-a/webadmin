<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/atseoul_extsiteitemcls.asp"-->

<%
dim itemid, itemname, eventid, mode, vAddQuery, vInsertQuery
dim itemidArr, eventidArr, makeridArr, delitemid
dim page, makerid

page    = request("page")
itemid  = request("itemid")
delitemid = requestCheckvar(request("delitemid"),9)
itemname= request("itemname")
eventid = request("eventid")
mode    = NullFillWith(request("mode"),"")
itemidArr = Trim(request("itemidArr"))
eventidArr= Trim(request("eventidArr"))
makeridArr = Trim(request("makeridArr"))
makerid= request("makerid")

if page="" then page=1
if Right(itemidArr,1)="," then itemidArr=Left(itemidArr,Len(itemidArr)-1)
if Right(eventidArr,1)="," then eventidArr=Left(eventidArr,Len(eventidArr)-1)


	vInsertQuery = ""
	vInsertQuery = " INSERT INTO [db_item].[dbo].tbl_atseoul_reg_item(itemid,reguserid) " + VbCrlf


	'############################## ���� ǰ�� ������ ##############################
	vAddQuery = ""
	vAddQuery = vAddQuery + " and t.itemid is null" + VbCrlf
	vAddQuery = vAddQuery + " and i.sellyn='Y'" + VbCrlf
	vAddQuery = vAddQuery + " and i.sellcash<>0" + VbCrlf
	''��ü ������۵� ����
	vAddQuery = vAddQuery + " and i.makerid not in (select userid from [db_user].[dbo].tbl_user_c where defaultDeliveryType is not NULL)" + VbCrlf

    ''���� ����Ŀ ����
    vAddQuery = vAddQuery + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun = 'atseoul')" + VbCrlf

	''�ɼ� �߰��ݾ� �ִ°� ����
	vAddQuery = vAddQuery + " and i.itemid not in (select distinct itemid from db_item.dbo.tbl_item_option where optaddprice>0)" + VbCrlf

	''�ش� ī�װ���
	vAddQuery = vAddQuery + " and (i.cate_large + i.cate_mid + i.cate_small) = (m.tencdl + m.tencdm + m.tencdn)" + VbCrlf

	''Ư����ǰ����
	vAddQuery = vAddQuery + " and i.itemid<>114039" + VbCrlf
	vAddQuery = vAddQuery + " and i.makerid<>'vanillaspoon'" + VbCrlf
	vAddQuery = vAddQuery + " and i.makerid<>'kongkkakji'" + VbCrlf

	''�ؿܹ�ۻ�ǰ
	vAddQuery = vAddQuery + " and i.mwdiv <> 'U'" + VbCrlf
	vAddQuery = vAddQuery + " and i.itemWeight > 0" + VbCrlf
	vAddQuery = vAddQuery + " and i.deliverOverseas = 'Y'" + VbCrlf

    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    vAddQuery = vAddQuery + " and i.isExtusing = 'Y'" + VbCrlf
	'############################## ���� ǰ�� ������ ##############################


	dim sqlStr, resultRow
	If mode <> "" AND mode = "delitem" Then
		dbget.Execute "Delete [db_item].[dbo].tbl_atseoul_reg_item Where itemid = '" & delitemid & "'"
		response.write "<script >alert('�����Ǿ����ϴ�.')</script>"
	ElseIf mode <> "" Then
		sqlStr = vInsertQuery
		if (mode="regByItemIDarr") then
		'### ��ǰ�ڵ�� ��� ### <!-- //-->
		    sqlStr = sqlStr + " 	SELECT TOP 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
		    sqlStr = sqlStr + "  			FROM" + VbCrlf
		    sqlStr = sqlStr + " 		[db_item].[dbo].tbl_item AS i" + VbCrlf
		    sqlStr = sqlStr + "     	LEFT JOIN  [db_item].[dbo].tbl_atseoul_reg_item AS t on i.itemid=t.itemid" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN  [db_item].[dbo].tbl_atseoul_category_mapping AS m on i.cate_large = m.tencdl and i.cate_mid = m.tencdm and i.cate_small = m.tencdn" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN db_shop.dbo.tbl_shop_designer sd on sd.shopid='streetshop881' and i.makerid=sd.makerid" + VbCrlf
		    sqlStr = sqlStr + "		WHERE " + VbCrlf
		    sqlStr = sqlStr + " 		i.itemid in (" + itemidArr + ")" + VbCrlf
		    sqlStr = sqlStr + " 		AND ((sellcash-buycash)/sellcash)*100>=15" + VbCrlf
		elseif (mode="regByEventIDarr") then
		'### �̺�Ʈ ��ȣ�� ��� ### <!-- //-->
		    sqlStr = sqlStr + "		SELECT TOP 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
		    sqlStr = sqlStr + "				FROM" + VbCrlf
		    sqlStr = sqlStr + " 		[db_event].[dbo].tbl_eventitem AS e," + VbCrlf
		    sqlStr = sqlStr + " 		[db_item].[dbo].tbl_item AS i" + VbCrlf
		    sqlStr = sqlStr + " 		LEFT JOIN  [db_item].[dbo].tbl_atseoul_reg_item AS t on i.itemid=t.itemid" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN  [db_item].[dbo].tbl_atseoul_category_mapping AS m on i.cate_large = m.tencdl and i.cate_mid = m.tencdm and i.cate_small = m.tencdn" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN db_shop.dbo.tbl_shop_designer sd on sd.shopid='streetshop881' and i.makerid=sd.makerid" + VbCrlf
		    sqlStr = sqlStr + " 	WHERE " + VbCrlf
		    sqlStr = sqlStr + " 		e.evt_code in (" + eventidArr + ")" + VbCrlf
		    sqlStr = sqlStr + " 		AND e.itemid=i.itemid" + VbCrlf
		    sqlStr = sqlStr + " 		AND (( i.sellcash- i.buycash)/ i.sellcash)*100>=15" + VbCrlf
		elseif (mode="recentBestSeller") then
		'### �ֱ� ����Ʈ ���� ��� ### <!-- //-->
		    sqlStr = sqlStr + " 	SELECT TOP 100 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
		    sqlStr = sqlStr + "  			FROM" + VbCrlf
		    sqlStr = sqlStr + " 		[db_item].[dbo].tbl_item_contents AS c, [db_item].[dbo].tbl_item AS i" + VbCrlf
		    sqlStr = sqlStr + " 		LEFT JOIN  [db_item].[dbo].tbl_atseoul_reg_item AS t on i.itemid=t.itemid" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN  [db_item].[dbo].tbl_atseoul_category_mapping AS m on i.cate_large = m.tencdl and i.cate_mid = m.tencdm and i.cate_small = m.tencdn" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN db_shop.dbo.tbl_shop_designer sd on sd.shopid='streetshop881' and i.makerid=sd.makerid" + VbCrlf
		    sqlStr = sqlStr + " 	WHERE " + VbCrlf
		    sqlStr = sqlStr + " 		i.itemid=c.itemid" + VbCrlf
		    sqlStr = sqlStr + " 		AND ((i.limityn='N') or (i.limityn='Y' and i.limitno-i.limitsold>=30))" + VbCrlf
		    sqlStr = sqlStr + " 		AND c.recentsellcount>=1" + VbCrlf
		    sqlStr = sqlStr + " 		AND sellcount>1" + VbCrlf
		    sqlStr = sqlStr + " 		AND ((sellcash-buycash)/sellcash)*100>=20" + VbCrlf
		elseif (mode="regByMakerid") then
		'### �귣��ID�� ��� ### <!-- //-->
		    sqlStr = sqlStr + " 	SELECT TOP 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
		    sqlStr = sqlStr + "  			FROM" + VbCrlf
		    sqlStr = sqlStr + " 		[db_item].[dbo].tbl_item AS i" + VbCrlf
		    sqlStr = sqlStr + "     	LEFT JOIN  [db_item].[dbo].tbl_atseoul_reg_item AS t on i.itemid=t.itemid" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN  [db_item].[dbo].tbl_atseoul_category_mapping AS m on i.cate_large = m.tencdl and i.cate_mid = m.tencdm and i.cate_small = m.tencdn" + VbCrlf
			sqlStr = sqlStr + "     	INNER JOIN db_shop.dbo.tbl_shop_designer sd on sd.shopid='streetshop881' and i.makerid=sd.makerid" + VbCrlf
		    sqlStr = sqlStr + " 	WHERE " + VbCrlf
		    sqlStr = sqlStr + " 		i.makerid ='" & makeridArr & "'" + VbCrlf
		    sqlStr = sqlStr + " 		AND ((sellcash-buycash)/sellcash)*100>=15" + VbCrlf
		end if
		sqlStr = sqlStr + vAddQuery + VbCrlf
	    dbget.Execute sqlStr, resultRow
	    'response.write sqlStr
	    'dbget.close()
		'response.end
	    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.');location.href='atseoulitem.asp';</script>"
	End If

dim oatseoulitem
set oatseoulitem = new CExtSiteItem
oatseoulitem.FPageSize       = 20
oatseoulitem.FCurrPage       = page
oatseoulitem.FRectItemID     = itemid
oatseoulitem.FRectItemName   = itemname
oatseoulitem.FRectEventid    = eventid
oatseoulitem.FRectMakerid    = makerid
oatseoulitem.GetAtSeoulRegedItemList

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
	window.open('/admin/etc/JaehyuMall_Not_In_Makerid.asp?mallgubun=atseoul','notin','width=200,height=400,scrollbars=yes');
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
		<input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9" class="input">
		��ǰ��:
		<input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="input">
		�̺�Ʈ��ȣ:
		<input type="text" name="eventid" value="<%= eventid %>" size="6" maxlength="6" class="input">
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmReg" method="post" action="atseoulitem.asp">
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
        	<td align="right"></td>
        </tr>
        </table>
    </td>
</tr>
</form>
</table>
<br>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oatseoulitem.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oatseoulitem.FTotalCount,0) %></td>
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
<% for i=0 to oatseoulitem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
    <td><img src="<%= oatseoulitem.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= oatseoulitem.FItemList(i).FItemID %></td>
    <td><%= oatseoulitem.FItemList(i).FItemName %></td>
    <td><%= oatseoulitem.FItemList(i).FRegdate %></td>
    <td><%= oatseoulitem.FItemList(i).Freguserid %></td>
    <td align="right"><%= FormatNumber(oatseoulitem.FItemList(i).FSellcash,0) %></td>
    <td align="center">
        <% if oatseoulitem.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oatseoulitem.FItemList(i).Fbuycash/oatseoulitem.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if oatseoulitem.FItemList(i).IsSoldOut then %>
        <font color="red">ǰ��</font>
        <% end if %>
    </td>
    <td align="center">
        <%= oatseoulitem.FItemList(i).Fatseoulcategory %>
    </td>
    <td align="center">
    	<a href="javascript:DelItem('<%= oatseoulitem.FItemList(i).FItemID %>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
    </td>
</tr>
<% next %>
<tr height="20">
    <td colspan="10" align="center" bgcolor="#FFFFFF">
        <% if oatseoulitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oatseoulitem.StarScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oatseoulitem.StarScrollPage to oatseoulitem.FScrollCount + oatseoulitem.StarScrollPage - 1 %>
    		<% if i>oatseoulitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oatseoulitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<%
set oatseoulitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->