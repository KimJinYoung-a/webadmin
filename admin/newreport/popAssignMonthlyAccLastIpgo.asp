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
        response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close()</script>"
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
        if (!confirm('마지막입고월이 지정되지 않았습니다. 계속하시겠습니까?')) return;
    }

    if (confirm('저장 하시겠습니까?')){
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
    <td width="100" bgcolor="#F3F3FF" >연월</td>
	<td ><%= yyyymm %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >재고위치</td>
	<td ><%= stockPlace %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >매장</td>
	<td ><%= shopid %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >마지막입고월</td>
	<td ><%= lastIpgodate %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >엑셀자료</td>
    <td >
        <input type="checkbox" name="modXL" value="Y" checked> 엑셀자료(복사본) 동시변경
    </td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >변경 </td>
    <td >
        <input type="text" class="text" name="lastIpgodate" value="<%= lastIpgodate %>" size="7" maxlength="7">
		* 삭제 : NULL 입력
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
    <input type="button" class="button" value="저장" onClick="saveThis()">
    </td>
</tr>
</form>
</table>

* 물류 마지막입고월을 변경하는 경우 전체 매장의 물류입고월도 같이 변경됩니다.

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
