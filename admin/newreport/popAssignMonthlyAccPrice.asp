<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption
dim mode, avgprice, lastprice, lastmwdiv

yyyymm 		= requestCheckvar(request("yyyymm"),20)
stockPlace  = requestCheckvar(request("stockPlace"),20)
shopid  	= requestCheckvar(request("shopid"),32)
itemgubun  	= requestCheckvar(request("itemgubun"),20)
itemid  	= requestCheckvar(request("itemid"),20)
itemoption  = requestCheckvar(request("itemoption"),20)
mode  		= requestCheckvar(request("mode"),32)
avgprice  	= requestCheckvar(request("avgprice"),32)
lastprice  	= requestCheckvar(request("lastprice"),32)
lastmwdiv  	= requestCheckvar(request("lastmwdiv"),32)

''if (stockPlace <> "L") then
''	response.write "에러"
''	dbget.close()
''	response.end
''end if


Dim sqlStr, ArrList, i, AssignedRow

IF (mode="act") then
	AssignedRow = 0

	if (stockPlace = "L") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
		sqlStr = sqlStr & " set lastbuyprice = " + CStr(lastprice) + ", avgipgoprice = " + CStr(avgprice) + " " & vbCrLf
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
		dbget.Execute sqlStr,AssignedRow

		if (lastmwdiv = "M") then
			sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
			sqlStr = sqlStr & " set avgshopipgoprice = " + CStr(avgprice) + " " & vbCrLf
			sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' and LstComm_cd = 'B031' "
			dbget.Execute sqlStr,AssignedRow
		end if
	end if

	if (stockPlace = "S") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
		sqlStr = sqlStr & " set avgshopipgoprice = " + CStr(avgprice) + ", LstBuyCash = " & lastprice & " " & vbCrLf
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' and LstComm_cd = 'B031' "
		dbget.Execute sqlStr,AssignedRow
	end if

    IF (AssignedRow>0) then
        response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

if (stockPlace = "L") then
	sqlStr = " select top 1 lastbuyprice, avgipgoprice, lastmwdiv from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
end if

if (stockPlace = "S") then
	sqlStr = " select top 1 LstBuyCash as lastbuyprice, avgShopipgoprice as avgipgoprice, LstComm_cd as lastmwdiv from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
	''rw sqlStr
end if

rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	lastprice = rsget("lastbuyprice")
	avgprice = rsget("avgipgoprice")
	lastmwdiv = rsget("lastmwdiv")
end if
rsget.Close

%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;

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
		<input type="hidden" name="lastmwdiv" value="<%=lastmwdiv%>">

		<tr align="center" bgcolor="#FFFFFF" height="20">
			<td width="80" bgcolor="#F3F3FF" >평균매입가</td>
			<td ><input type="text" size="6" name="avgprice" value="<%= avgprice %>"></td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF" height="20">
			<td width="80" bgcolor="#F3F3FF" >작성시매입가</td>
			<td ><input type="text" size="6" name="lastprice" value="<%= lastprice %>"></td>
		</tr>


		<tr bgcolor="#FFFFFF" height="40">
			<td colspan="7" align="center">
				<input type="button" class="button" value="저장" onClick="saveThis()">
			</td>
		</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
