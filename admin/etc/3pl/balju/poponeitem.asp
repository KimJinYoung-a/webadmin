<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/tplbalju.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
'' ����/����/��ǰ ����ǰ ����.

dim mode, itemid, itemoption, reguserid, page, divcd
dim bf_itemid, bf_itemoption
page    	= request("page")
mode    	= request("mode")
itemid  	= Trim(request("itemid"))
itemoption  = request("itemoption")
divcd  		= request("divcd")
reguserid	= session("ssBctId")

if page="" then page=1


dim sqlStr
if (mode="del") and (itemid<>"") then
	if (itemoption = "") then
		sqlStr = " delete from [db_threepl].[dbo].tbl_baljureg_item " + VbCrlf
		sqlStr = sqlStr + " where itemid=" + Cstr(itemid)
	else
		sqlStr = " delete from [db_threepl].[dbo].tbl_baljureg_item " + VbCrlf
		sqlStr = sqlStr + " where itemid=" + Cstr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
	end if
    dbget.Execute sqlStr

    itemid = ""
end if

if (mode="add") and (itemid<>"") and (divcd<>"") then
	if (Len(itemid) <= 8) then
		sqlStr = " insert into [db_threepl].[dbo].tbl_baljureg_item " + VbCrlf
		sqlStr = sqlStr + " (itemid, itemoption, divcd, reguserid) "+ VbCrlf
		sqlStr = sqlStr + " select i.itemid, '', '" & divcd & "', '" + reguserid + "'" + VbCrlf
		sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_tpl_item i" + VbCrlf
		sqlStr = sqlStr + " left join [db_threepl].[dbo].tbl_baljureg_item b on i.itemid=b.itemid and b.itemoption = '' " + VbCrlf
		sqlStr = sqlStr + " where i.itemid=" + CStr(itemid)
		sqlStr = sqlStr + " and b.itemid is null"
	else
		if BF_IsMaybeTenBarcode(itemid) then
			bf_itemid = BF_GetItemId(itemid)
			bf_itemoption = BF_GetItemOption(itemid)
			'// ���� ���ڵ�
			sqlStr = " insert into [db_threepl].[dbo].tbl_baljureg_item " + VbCrlf
			sqlStr = sqlStr + " (itemid, itemoption, divcd, reguserid) "+ VbCrlf
			sqlStr = sqlStr + " select i.itemid, i.itemoption, '" & divcd & "', '" + reguserid + "'" + VbCrlf
			sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_tpl_item i" + VbCrlf
			sqlStr = sqlStr + " left join [db_threepl].[dbo].tbl_baljureg_item b on i.itemid=b.itemid and b.itemoption = i.itemoption " + VbCrlf
			sqlStr = sqlStr + " where i.itemid = " & bf_itemid & " and i.itemoption = '" & bf_itemoption & "' "
			sqlStr = sqlStr + " and b.itemid is null"
		else
			'// ������ڵ�
			sqlStr = " insert into [db_threepl].[dbo].tbl_baljureg_item " + VbCrlf
			sqlStr = sqlStr + " (itemid, itemoption, divcd, reguserid) "+ VbCrlf
			sqlStr = sqlStr + " select i.itemid, i.itemoption, '" & divcd & "', '" + reguserid + "'" + VbCrlf
			sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_tpl_item i" + VbCrlf
			sqlStr = sqlStr + " left join [db_threepl].[dbo].tbl_baljureg_item b on i.itemid=b.itemid and b.itemoption = i.itemoption " + VbCrlf
			sqlStr = sqlStr + " where i.barcode='" + CStr(itemid) + "'"
			sqlStr = sqlStr + " and b.itemid is null"
		end if
	end if

    dbget.Execute sqlStr
end if

dim odanpumbalju
set odanpumbalju = new CTenBalju
odanpumbalju.FPageSize=50
odanpumbalju.FCurrpage = page
odanpumbalju.FRectItemDivCD = divcd

odanpumbalju.GetDanpumBaljuItemList

dim i
%>
<script lanuage='javascript'>
function DelItem(iitemid, iitemoption){
   if (confirm('���� �Ͻðڽ��ϱ�?')){
        dellfrm.mode.value="del";
        dellfrm.itemid.value= iitemid;
		dellfrm.itemoption.value= iitemoption;
        dellfrm.submit();
    }
}


function AddItem(frm){
    if (frm.divcd.value == ""){
        alert('������ �����ϼ���.');
        return;
    }

    if (frm.itemid.value.length<3){
        alert('��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.');
        frm.itemid.focus();
        return;
    }

    frm.mode.value="add";
    frm.submit();
}

function NextPage(page){
    frmbar.page.value=page;
    frmbar.submit();
}
</script>
<!-- ������ ����� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("menubar") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>��ǰ����ǰ ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			����/����/��ǰ ��ǰ�� �����ϴ� ���Դϴ�.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>
<!-- ������ ����� ���� -->

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frmbar" method=get>
    <input type="hidden" name="mode" value="">
    <input type="hidden" name="page" value="">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
			<select name="divcd" >
			<option value="" 	<% if divcd="" then response.write "selected" %> >����</option>
			<option value="O" 	<% if divcd="O" then response.write "selected" %> >��ǰ��ǰ</option>
			<option value="E" 	<% if divcd="E" then response.write "selected" %> >���ܻ�ǰ</option>
			<option value="I" 	<% if divcd="I" then response.write "selected" %> >���Ի�ǰ</option>
			</select>
        	��ǰ�ڵ�(��ü�ɼ�)/���ڵ�(Ư���ɼ�) : <input type="text" name="itemid" value="<%= itemid %>" Maxlength="20" size="13" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ AddItem(frmbar); return false;}">
        	<input type="button" value="��ǰ�߰�" onclick="AddItem(frmbar)">
        </td>
        <td align="right">

        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">����</td>
    	<td width="60">��ǰ�ڵ�</td>
		<td width="50">�ɼ�</td>
    	<td width="50">�̹���</td>
    	<td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
      	<td>�����</td>
      	<td width="50">����</td>
    </tr>
    <% for i=0 to odanpumbalju.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= odanpumbalju.FItemList(i).GetDivCDString %></td>
    	<td><%= odanpumbalju.FItemList(i).FItemID %></td>
		<td>
			<% if (odanpumbalju.FItemList(i).FDivCD <> "O") then %>
			<%= odanpumbalju.FItemList(i).FItemOption %>
			<% end if %>
		</td>
    	<td><img src="<%= odanpumbalju.FItemList(i).FImageSmall %>" width="50"></td>
    	<td><%= odanpumbalju.FItemList(i).FItemName %><br><font color="blue">[<%= odanpumbalju.FItemList(i).FItemOptionName %>]</font></td>
      	<td><%= odanpumbalju.FItemList(i).Freguserid %></td>
      	<td><a href="javascript:DelItem(<%= odanpumbalju.FItemList(i).FItemID %>, '<%= odanpumbalju.FItemList(i).FItemOption %>');"><img src="/images/icon_delete.gif" border="0"></a></td>
    </tr>
    <% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <% if odanpumbalju.HasPreScroll then %>
        		<a href="javascript:NextPage('<%= odanpumbalju.StartScrollPage-1 %>')">[pre]</a>
        	<% else %>
        		[pre]
        	<% end if %>

        	<% for i=0 + odanpumbalju.StartScrollPage to odanpumbalju.FScrollCount + odanpumbalju.StartScrollPage - 1 %>
        		<% if i>odanpumbalju.FTotalpage then Exit for %>
        		<% if CStr(page)=CStr(i) then %>
        		<font color="red">[<%= i %>]</font>
        		<% else %>
        		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
        		<% end if %>
        	<% next %>

        	<% if odanpumbalju.HasNextScroll then %>
        		<a href="javascript:NextPage('<%= i %>')">[next]</a>
        	<% else %>
        		[next]
        	<% end if %>

        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->



<%
set odanpumbalju = Nothing
%>
<form name="dellfrm" method=get action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
