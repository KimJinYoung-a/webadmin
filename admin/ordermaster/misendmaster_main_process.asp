<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
dim mode      : mode = request("mode")
dim detailIDx : detailIDx = request("detailIDx")
dim orderserial, didx, finishstr, state, prevstate

dim writeuser,contents_jupsu
dim itemid, itemoption, itemname, itemoptionname, modiuserid

orderserial = request("orderserial")
didx        = request("didx")
finishstr   = request("finishstr")
state       = request("state")
prevstate   = request("prevstate")

didx = didx + ",,"
didx = Split(didx, ",")

finishstr = finishstr + ",,"
finishstr = Split(finishstr, ",")

state = state + ",,"
state = Split(state, ",")

prevstate = prevstate + ",,"
prevstate = Split(prevstate, ",")

dim sqlStr,i

if (mode="SendCallChange") then
    sqlStr = "update [db_temp].[dbo].tbl_mibeasong_list " &VbCRLF
	sqlStr = sqlStr + " set isSendCall = 'Y' " &VbCRLF
	sqlStr = sqlStr + " ,state=4" &VbCRLF
	sqlStr = sqlStr + " ,sendCount=IsNull(sendCount,0) + 1 " &VbCRLF
	sqlStr = sqlStr + " ,lastSendUserid='" + CStr(session("ssBctId")) + "' " &VbCRLF
	sqlStr = sqlStr + " ,lastSendDate=getdate()" &VbCRLF
	sqlStr = sqlStr + " where detailidx=" + CStr(detailIDx) &VbCRLF
    dbget.Execute sqlStr

    '// ǰ�����Ұ� ����� ����(�ش� �ֹ� ��ü ǰ�����Ұ� �ȳ���)
	sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] " & detailIDx & " "
	dbget.Execute sqlStr
elseif (mode="cancelFin") then
    sqlStr = "update [db_temp].[dbo].tbl_mibeasong_list " &VbCRLF
	sqlStr = sqlStr + " set state=9" &VbCRLF
	sqlStr = sqlStr + " where detailidx=" + CStr(detailIDx) &VbCRLF

    dbget.Execute sqlStr
else
    for i = 0 to UBound(didx)
    	if (trim(didx(i)) <> "") then
			if (trim(prevstate(i)) = "4") and (trim(state(i)) = "4") and C_ADMIN_AUTH then
				'// ���ȳ� => ��ó�� ��ȯ

				'// CS������� ���ζ�� �ϼż� �ϴ� ���(2015-06-11, skyer9)
				response.end

				sqlStr = "select top 1 itemid, itemoption, itemname, itemoptionname, modiuserid from [db_temp].[dbo].tbl_mibeasong_list "
				sqlStr = sqlStr + " where detailidx=" + CStr(didx(i)) &VbCRLF
				rsget.Open sqlStr, dbget

				itemid = ""
				if  not rsget.EOF  then
					itemid 			= rsget("itemid")
					itemoption 		= rsget("itemoption")
					itemname 		= db2html(rsget("itemname"))
					itemoptionname 	= db2html(rsget("itemoptionname"))
					modiuserid 		= rsget("modiuserid")

					if IsNull(modiuserid) then
						modiuserid = session("ssBctId")
					end if
				end if
				rsget.close

				if (itemid <> "") then
					writeuser = modiuserid
					contents_jupsu = "[TEST] ������� ���ȳ� =&gt; ��ó�� ��ȯ<br>��ǰ�ڵ� : " & itemid & "[" & itemoption & "]<br>" & itemname & "[" & itemoptionname & "]<br>ó���� : " & session("ssBctId")

					call AddCsMemo(orderserial, "2", "", writeuser, contents_jupsu)
				end if
			end if

    		sqlStr = "update [db_temp].[dbo].tbl_mibeasong_list " &VbCRLF
    		sqlStr = sqlStr + " set finishstr = '" + trim(finishstr(i)) + "'"&VbCRLF
    		sqlStr = sqlStr + " , state = '" + trim(state(i)) + "' " &VbCRLF
			sqlStr = sqlStr + "	,modiuserid = '" + CStr(session("ssBctId")) + "' "
			sqlStr = sqlStr + "	,modidate = getdate() "
    		sqlStr = sqlStr + " where detailidx=" + CStr(didx(i)) &VbCRLF

    		dbget.Execute sqlStr
    	end if
    next
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
<% if (mode="cancelFin") then %>
    alert('ó�� �Ǿ����ϴ�.');
    window.close()
<% else %>
    alert('���� �Ǿ����ϴ�.');
    location.replace('misendmaster_main.asp?orderserial=<%= orderserial %>');
<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
