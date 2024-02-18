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

yyyymm 		= requestCheckvar(request("yyyymm"),20)
stockPlace  = requestCheckvar(request("stockPlace"),20)
shopid  	= requestCheckvar(request("shopid"),20)
itemgubun  	= requestCheckvar(request("itemgubun"),20)
itemid  	= requestCheckvar(request("itemid"),20)
itemoption  = requestCheckvar(request("itemoption"),20)
mode  		= requestCheckvar(request("mode"),20)
chgmakerid  = requestCheckvar(request("chgmakerid"),20)
modiiteminfo= requestCheckvar(request("modiiteminfo"),20)


Dim sqlStr, ArrList, i, AssignedRow

IF (mode="act") then
	AssignedRow = 0

	if (stockPlace = "L") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
		sqlStr = sqlStr & " set lastmakerid = '" + CStr(chgmakerid) + "' " & vbCrLf
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

	if (stockPlace = "S") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
		sqlStr = sqlStr & " set LstMakerid = '" + CStr(chgmakerid) + "' " & vbCrLf
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

	if (modiiteminfo = "Y") then
		sqlStr = " update [db_shop].[dbo].[tbl_shop_item] " & vbCrLf
		sqlStr = sqlStr & " set makerid = '" + CStr(chgmakerid) + "' " & vbCrLf
		sqlStr = sqlStr & " 	where itemgubun = '" + CStr(itemgubun) + "' and shopitemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow
		''response.write sqlStr
	end if

    IF (AssignedRow>0) then
        response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

if (stockPlace = "L") then
	sqlStr = " select top 1 lastmwdiv as mwdiv, lastmakerid as LstMakerid from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
	''rw sqlStr
end if

if (stockPlace = "S") then
	sqlStr = " select top 1 LstComm_cd as mwdiv, LstMakerid, lstbuycash from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

	'rw sqlStr
	'response.end
end if

rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	makerid = rsget("LstMakerid")
end if
rsget.Close


if isNULL(makerid) then makerid=""

%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;

    if (frm.chgmakerid.value.length<1){
        if (!confirm('출고 구분이 지정되지 않았습니다. 계속하시겠습니까?')) return;
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
    <td width="80" bgcolor="#F3F3FF" >현재 </td>
	<td ><%= makerid %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >변경 </td>
    <td >
		<input type="text" size="20" name="chgmakerid" value="<%= makerid %>">
		<input type="checkbox" name="modiiteminfo" value="Y" <%= CHKIIF(itemgubun="10", "disabled", "") %>> 상품정보 브랜드 동시변경
	</td>
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
