<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption
dim mode, chgmakerid
dim makerid, modiiteminfo
dim dataExists

stockPlace  = requestCheckvar(request("stockPlace"),20)
shopid  	= requestCheckvar(request("shopid"),20)
itemgubun  	= requestCheckvar(request("itemgubun"),20)
itemid  	= requestCheckvar(request("itemid"),20)
itemoption  = requestCheckvar(request("itemoption"),20)
mode  		= requestCheckvar(request("mode"),20)
chgmakerid  = requestCheckvar(request("chgmakerid"),20)
modiiteminfo= requestCheckvar(request("modiiteminfo"),20)

if (stockPlace = "") then
	stockPlace = "L"
end if

Dim sqlStr, ArrList, i, AssignedRow

IF (mode="add") then
	AssignedRow = 0

	if (stockPlace = "L") then
		sqlStr = " insert into db_summary.dbo.tbl_not_inc_SummaryStock(itemgubun, itemid, itemoption) " & vbCrLf
		sqlStr = sqlStr & " values('" & itemgubun & "', '" & itemid & "', '" & itemoption & "') "
		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

	if (stockPlace = "S") then
		sqlStr = " insert into db_summary.dbo.tbl_not_inc_SummaryStock_SHOP(shopid, itemgubun, itemid, itemoption) " & vbCrLf
		sqlStr = sqlStr & " values('" & shopid & "', '" & itemgubun & "', '" & itemid & "', '" & itemoption & "') "
		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

    IF (AssignedRow>0) then
        response.write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
ELSEIF (mode="del") then
	if (stockPlace = "L") then
		sqlStr = " delete from db_summary.dbo.tbl_not_inc_SummaryStock " & vbCrLf
		sqlStr = sqlStr & " where itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

	if (stockPlace = "S") then
		sqlStr = " delete from db_summary.dbo.tbl_not_inc_SummaryStock_SHOP " & vbCrLf
		sqlStr = sqlStr & " where shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

    IF (AssignedRow>0) then
        response.write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

if (stockPlace = "L") then
	sqlStr = " select top 1 itemgubun, itemid, itemoption from db_summary.dbo.tbl_not_inc_SummaryStock " & vbCrLf
	sqlStr = sqlStr & " 	where itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
	''rw sqlStr
end if

if (stockPlace = "S") then
	sqlStr = " select top 1 shopid, itemgubun, itemid, itemoption from db_summary.dbo.tbl_not_inc_SummaryStock_SHOP " & vbCrLf
	sqlStr = sqlStr & " 	where shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
	'rw sqlStr
	'response.end
end if

dataExists = False
rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	dataExists = True
	response.write "<script>alert('�̹� ��ϵ� ��ǰ�ڵ��Դϴ�.')</script>"
end if
rsget.Close

%>
<script language='javascript'>

function jsSearch() {
	var frm = document.frmAct;
	if ((frm.stockPlace.value == "S") && (frm.shopid.value == "")) {
		alert('������ �Է��ϼ���.');
		return;
	}

	frm.submit();
}

function delThis(stockPlace) {
	var frm = document.frmAct;
	frm.stockPlace.value = stockPlace;
	if ((frm.stockPlace.value == "S") && (frm.shopid.value == "")) {
		alert('������ �Է��ϼ���.');
		return;
	}

	if (confirm('����ڻ� ���ܻ�ǰ : �����Ͻðڽ��ϱ�?')) {
		frm.mode.value = "del";
		frm.method = 'post';
		frm.submit();
	}
}

function addThis(stockPlace) {
	var frm = document.frmAct;
	frm.stockPlace.value = stockPlace;
	if ((frm.stockPlace.value == "S") && (frm.shopid.value == "")) {
		alert('������ �Է��ϼ���.');
		return;
	}

	if (confirm('����ڻ� ���ܻ�ǰ : ����Ͻðڽ��ϱ�?')) {
		frm.mode.value = "add";
		frm.method = 'post';
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmAct" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubun" value="<%=itemgubun%>">
<input type="hidden" name="itemid" value="<%=itemid%>">
<input type="hidden" name="itemoption" value="<%=itemoption%>">
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >�����ġ </td>
	<td >
		<input type="radio" name="stockPlace" value="L" <%= CHKIIF(stockPlace="L", "checked", "") %>> ����
		<input type="radio" name="stockPlace" value="S" <%= CHKIIF(stockPlace="S", "checked", "") %>> ����
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >���� </td>
	<td >
		<input type="text" class="text" name="shopid" value="<%= shopid %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >��ǰ����</td>
	<td ><%= itemgubun %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >��ǰ�ڵ�</td>
	<td ><%= itemid %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >�ɼ�</td>
	<td ><%= itemoption %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >������� </td>
	<td >
		<%= CHKIIF(dataExists, "��ϿϷ�", "�̵��") %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
		<input type="button" class="button" value="�˻�" onClick="jsSearch()">
		&nbsp;
		<% if dataExists then %>
		<input type="button" class="button" value="����(<%= CHKIIF(stockPlace="L", "����", "����")%>)" onClick="delThis('<%= stockPlace %>')">
		<% else %>
		<input type="button" class="button" value="���(<%= CHKIIF(stockPlace="L", "����", "����")%>)" onClick="addThis('<%= stockPlace %>')">
		<% end if %>
    </td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
