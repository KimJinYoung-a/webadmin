<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption
dim mode, chgmwdiv
dim mwdiv

yyyymm 		= requestCheckvar(request("yyyymm"),20)
stockPlace  = requestCheckvar(request("stockPlace"),20)
shopid  	= requestCheckvar(request("shopid"),20)
itemgubun  	= requestCheckvar(request("itemgubun"),20)
itemid  	= requestCheckvar(request("itemid"),20)
itemoption  = requestCheckvar(request("itemoption"),20)
mode  		= requestCheckvar(request("mode"),20)
chgmwdiv  = requestCheckvar(request("chgmwdiv"),20)


Dim sqlStr, ArrList, i, AssignedRow

IF (mode="act") then
	AssignedRow = 0

	if (stockPlace = "L") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
		sqlStr = sqlStr & " set lastmwdiv = '" + CStr(chgmwdiv) + "' " & vbCrLf
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow
	end if

	if (stockPlace = "S") then
		sqlStr = " update db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
		sqlStr = sqlStr & " set LstComm_cd = '" + CStr(chgmwdiv) + "' " & vbCrLf
		sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

		dbget.Execute sqlStr,AssignedRow
	end if

    IF (AssignedRow>0) then
        response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

if (stockPlace = "L") then
	sqlStr = " select top 1 lastmwdiv as mwdiv from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "
end if

if (stockPlace = "S") then
	sqlStr = " select top 1 LstComm_cd as mwdiv, lstmakerid, lstbuycash from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary " & vbCrLf
	sqlStr = sqlStr & " 	where yyyymm = '" + CStr(yyyymm) + "' and shopid = '" + CStr(shopid) + "' and itemgubun = '" + CStr(itemgubun) + "' and itemid = '" + CStr(itemid) + "' and itemoption = '" + CStr(itemoption) + "' "

	'rw sqlStr
	'response.end
end if

rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	mwdiv = rsget("mwdiv")
end if
rsget.Close


if isNULL(mwdiv) then mwdiv=""

function getMaeipGubunName(mwdiv)
	if mwdiv="M" then
		getMaeipGubunName = "매입"
	elseif mwdiv="W" then
		getMaeipGubunName = "위탁"
	elseif mwdiv="U" then
		getMaeipGubunName = "업체"
	elseif mwdiv="Z" then
		getMaeipGubunName = "-"
	elseif mwdiv="B011" then
		getMaeipGubunName = "위탁판매"
	elseif mwdiv="B012" then
		getMaeipGubunName = "업체위탁"
	elseif mwdiv="B013" then
		getMaeipGubunName = "출고위탁"
	elseif mwdiv="B021" then
		getMaeipGubunName = "오프매입"
	elseif mwdiv="B022" then
		getMaeipGubunName = "매장매입"
	elseif mwdiv="B023" then
		getMaeipGubunName = "가맹점매입"
	elseif mwdiv="B031" then
		getMaeipGubunName = "출고매입"
	elseif mwdiv="B032" then
		getMaeipGubunName = "센터매입"
	else
		getMaeipGubunName = mwdiv
	end if

	IF isNULL(mwdiv) then
		getMaeipGubunName ="-"
	end if

	IF mwdiv = "" then
		getMaeipGubunName ="??"
	end if
end function

%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;

    if (frm.chgmwdiv.value.length<1){
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
	<td ><%= getMaeipGubunName(mwdiv) %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >변경 </td>
    <td >
        <select name="chgmwdiv">
			<option value=""> 미지정</option>
			<% if (stockPlace = "L") then %>
			<option value="M" <%=CHKIIF(mwdiv="M","selected","")%> >매입</option>
			<option value="W" <%=CHKIIF(mwdiv="W","selected","")%> >위탁</option>
			<% end if %>
			<% if (stockPlace = "S") then %>
			<option value="B012" <%=CHKIIF(mwdiv="B012","selected","")%> >업체위탁</option>
			<option value="B031" <%=CHKIIF(mwdiv="B031","selected","")%> >출고매입</option>
			<option value="">---------</option>
			<option value="B013" <%=CHKIIF(mwdiv="B013","selected","")%> >출고위탁</option>
			<% end if %>
        </select>
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
