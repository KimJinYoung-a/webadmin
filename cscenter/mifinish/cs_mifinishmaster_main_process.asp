<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode      : mode = request("mode")
dim csdetailidx : csdetailidx = request("csdetailidx")
dim asid, arrcsdetailidx, finishstr, state


asid 			= request("asid")
arrcsdetailidx  = request("arrcsdetailidx")
finishstr   	= request("finishstr")
state       	= request("state")


arrcsdetailidx = arrcsdetailidx + ",,"
arrcsdetailidx = Split(arrcsdetailidx, ",")

finishstr = finishstr + ",,"
finishstr = Split(finishstr, ",")



state = state + ",,"
state = Split(state, ",")

dim sqlStr,i

if (mode="SendCallChange") then
    sqlStr = "update [db_temp].[dbo].tbl_csmifinish_list " &VbCRLF
	sqlStr = sqlStr + " set isSendCall = 'Y' " &VbCRLF
	sqlStr = sqlStr + " ,state=4" &VbCRLF
	sqlStr = sqlStr + " where csdetailidx=" + CStr(csdetailidx) &VbCRLF
    dbget.Execute sqlStr

    '// ǰ�����Ұ� ����� ����(�ش� �ֹ� ��ü ǰ�����Ұ� �ȳ���)
	sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeIDaaa] " & csdetailidx & " "
	''dbget.Execute sqlStr

	response.write "<script language='javascript'>alert('TODO : ǰ�����Ұ� ����� ����.');</script>"
elseif (mode="cancelFin") then
    sqlStr = "update [db_temp].[dbo].tbl_csmifinish_list " &VbCRLF
	sqlStr = sqlStr + " set state=9" &VbCRLF
	sqlStr = sqlStr + " where csdetailidx=" + CStr(csdetailidx) &VbCRLF

    dbget.Execute sqlStr
else
    for i = 0 to UBound(arrcsdetailidx)
    	if (trim(arrcsdetailidx(i)) <> "") then
    		finishstr(i) = Replace(finishstr(i), "_XX_", ",")

    		sqlStr = "update [db_temp].[dbo].tbl_csmifinish_list " &VbCRLF
    		sqlStr = sqlStr + " set finishstr = '" + trim(finishstr(i)) + "'"&VbCRLF
    		sqlStr = sqlStr + " , state = '" + trim(state(i)) + "' " &VbCRLF
    		sqlStr = sqlStr + " where csdetailidx=" + CStr(arrcsdetailidx(i)) &VbCRLF

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
    location.replace('cs_mifinishmaster_main.asp?asid=<%= asid %>');
<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->