<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption
dim mode, lastIpgodate
dim mwdiv, modXL

yyyymm 		= requestCheckvar(request("yyyymm"),20)
stockPlace  = requestCheckvar(request("stockPlace"),20)
shopid  	= requestCheckvar(request("shopid"),20)
itemgubun  	= requestCheckvar(request("itemgubun"),20)
itemid  	= requestCheckvar(request("itemid"),20)
itemoption  = requestCheckvar(request("itemoption"),20)

mode  		= requestCheckvar(request("mode"),20)
modXL  		= requestCheckvar(request("modXL"),20)

lastIpgodate  = requestCheckvar(request("lastIpgodate"),20)

if (stockPlace = "M") then
	stockPlace = "S"
end if


Dim sqlStr, ArrList, i, AssignedRow, AssignedRowSUM

IF (mode="act") then
	AssignedRow = 0
	AssignedRowSUM = 0

	if (stockPlace = "L") and (lastIpgodate <> "") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
		if (lastIpgodate = "NULL") then
			sqlStr = sqlStr & " set lastIpgodate = NULL " & vbCrLf
		else
			sqlStr = sqlStr & " set lastIpgodate = '" + CStr(lastIpgodate) + "' " & vbCrLf
		end if
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
		dbget.Execute sqlStr,AssignedRow

		AssignedRowSUM = AssignedRowSUM + AssignedRow

		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
		if (lastIpgodate = "NULL") then
			sqlStr = sqlStr & " set lastIpgodateLogics = NULL " & vbCrLf
		else
			sqlStr = sqlStr & " set lastIpgodateLogics = '" + CStr(lastIpgodate) + "' " & vbCrLf
		end if
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
		dbget.Execute sqlStr,AssignedRow

		AssignedRowSUM = AssignedRowSUM + AssignedRow

        if (modXL = "Y") then
		    sqlStr = " update [db_summary].dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 " & vbCrLf
		    if (lastIpgodate = "NULL") then
			    sqlStr = sqlStr & " set lastIpgodate = NULL " & vbCrLf
		    else
			    sqlStr = sqlStr & " set lastIpgodate = '" + CStr(lastIpgodate) + "' " & vbCrLf
		    end if
		    sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
		    dbget.Execute sqlStr,AssignedRow

		    AssignedRowSUM = AssignedRowSUM + AssignedRow
        end if
	end if

	if (stockPlace = "S") and (lastIpgodate <> "") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
		if (lastIpgodate = "NULL") then
			sqlStr = sqlStr & " set lastIpgodate = NULL " & vbCrLf
		else
			sqlStr = sqlStr & " set lastIpgodate = '" + CStr(lastIpgodate) + "' " & vbCrLf
		end if
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow

		AssignedRowSUM = AssignedRowSUM + AssignedRow
	end if

	if (stockPlace = "T") and (lastIpgodate <> "") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
		if (lastIpgodate = "NULL") then
			sqlStr = sqlStr & " set lastIpgodateLogics = NULL " & vbCrLf
		else
			sqlStr = sqlStr & " set lastIpgodateLogics = '" + CStr(lastIpgodate) + "' " & vbCrLf
		end if
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow

		AssignedRowSUM = AssignedRowSUM + AssignedRow
	end if

    IF (AssignedRowSUM>0) then
        response.write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

if (stockPlace = "L") then
	sqlStr = " select top 1 lastIpgodate from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
end if

if (stockPlace = "S") then
	sqlStr = " select top 1 lastIpgodate from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

	'rw sqlStr
	'response.end
end if

if (stockPlace = "T") then
	sqlStr = " select top 1 lastIpgodateLogics as lastIpgodate from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

	'rw sqlStr
	'response.end
end if

rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	lastIpgodate = rsget("lastIpgodate")
end if
rsget.Close

%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;

    if (frm.lastIpgodate.value.length<1){
        if (!confirm('�������԰���� �������� �ʾҽ��ϴ�. ����Ͻðڽ��ϱ�?')) return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value="act";
        frm.submit();
    }
}

</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmAct" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="yyyymm" value="<%=yyyymm%>">
<input type="hidden" name="stockPlace" value="<%=stockPlace%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="itemgubun" value="<%=itemgubun%>">
<input type="hidden" name="itemid" value="<%=itemid%>">
<input type="hidden" name="itemoption" value="<%=itemoption%>">
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="100" bgcolor="#F3F3FF" >����</td>
	<td ><%= yyyymm %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >�����ġ</td>
	<td ><%= stockPlace %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >����</td>
	<td ><%= shopid %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >�������԰��</td>
	<td ><%= lastIpgodate %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >�����ڷ�</td>
    <td >
        <input type="checkbox" name="modXL" value="Y" checked> �����ڷ�(���纻) ���ú���
    </td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >���� </td>
    <td >
        <input type="text" class="text" name="lastIpgodate" value="<%= lastIpgodate %>" size="7" maxlength="7">
		* ���� : NULL �Է�
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
    <input type="button" class="button" value="����" onClick="saveThis()">
    </td>
</tr>
</form>
</table>

* ���� �������԰���� �����ϴ� ��� ��ü ������ �����԰���� ���� ����˴ϴ�.

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
