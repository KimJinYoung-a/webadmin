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
        response.write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close()</script>"
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
		getMaeipGubunName = "����"
	elseif mwdiv="W" then
		getMaeipGubunName = "��Ź"
	elseif mwdiv="U" then
		getMaeipGubunName = "��ü"
	elseif mwdiv="Z" then
		getMaeipGubunName = "-"
	elseif mwdiv="B011" then
		getMaeipGubunName = "��Ź�Ǹ�"
	elseif mwdiv="B012" then
		getMaeipGubunName = "��ü��Ź"
	elseif mwdiv="B013" then
		getMaeipGubunName = "�����Ź"
	elseif mwdiv="B021" then
		getMaeipGubunName = "��������"
	elseif mwdiv="B022" then
		getMaeipGubunName = "�������"
	elseif mwdiv="B023" then
		getMaeipGubunName = "����������"
	elseif mwdiv="B031" then
		getMaeipGubunName = "������"
	elseif mwdiv="B032" then
		getMaeipGubunName = "���͸���"
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
        if (!confirm('��� ������ �������� �ʾҽ��ϴ�. ����Ͻðڽ��ϱ�?')) return;
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
    <td width="80" bgcolor="#F3F3FF" >���� </td>
	<td ><%= getMaeipGubunName(mwdiv) %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >���� </td>
    <td >
        <select name="chgmwdiv">
			<option value=""> ������</option>
			<% if (stockPlace = "L") then %>
			<option value="M" <%=CHKIIF(mwdiv="M","selected","")%> >����</option>
			<option value="W" <%=CHKIIF(mwdiv="W","selected","")%> >��Ź</option>
			<% end if %>
			<% if (stockPlace = "S") then %>
			<option value="B012" <%=CHKIIF(mwdiv="B012","selected","")%> >��ü��Ź</option>
			<option value="B031" <%=CHKIIF(mwdiv="B031","selected","")%> >������</option>
			<option value="">---------</option>
			<option value="B013" <%=CHKIIF(mwdiv="B013","selected","")%> >�����Ź</option>
			<% end if %>
        </select>
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
    <input type="button" class="button" value="����" onClick="saveThis()">
    </td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
