<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 ERP연동
' History : 2011.12.16 eastone  생성 erpLink_Process.asp
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
Dim mode : mode = reQuestCheckvar(request("mode"),32)
Dim chk : chk = request("chk")                      ''Idx Array

dim idx, divcd, appDate, SELLBIZSECTION_CD, BUYBIZSECTION_CD, totalSum, supplySum, taxSum, vatyn, reguserid, selluserid, acc_cd

idx = reQuestCheckvar(request("idx"),32)
divcd = reQuestCheckvar(request("divcd"),32)
appDate = reQuestCheckvar(request("appDate"),32)
SELLBIZSECTION_CD = reQuestCheckvar(request("SELLBIZSECTION_CD"),32)
BUYBIZSECTION_CD = reQuestCheckvar(request("BUYBIZSECTION_CD"),32)
totalSum = reQuestCheckvar(request("totalSum"),32)
supplySum = reQuestCheckvar(request("supplySum"),32)
taxSum = reQuestCheckvar(request("taxSum"),32)
vatyn = reQuestCheckvar(request("vatyn"),32)
reguserid = session("ssBctId")
selluserid = session("ssBctId")
acc_cd = reQuestCheckvar(request("acc_cd"),32)


'chk = Split(chk,",")

dim i, j, sqlStr

'response.write chk
'response.write chk

if (mode = "delselectedarr") then

	sqlStr = " update db_partner.dbo.tbl_InternalOrder "
	sqlStr = sqlStr & " set useyn = 'N', deldate = getdate() "
	sqlStr = sqlStr & " where idx in (" & CStr(chk) & ") "
	dbget.Execute sqlStr

	response.write "<script>alert('삭제되었습니다.'); history.back();</script>"

elseif (mode = "ins") then

	if (taxSum = 0) then
		vatyn = "N"
	else
		vatyn = "Y"
	end if

	sqlStr = " insert into db_partner.dbo.tbl_InternalOrder(divcd, appDate, SELLBIZSECTION_CD, BUYBIZSECTION_CD, totalSum, supplySum, taxSum, vatyn, reguserid, selluserid, acc_cd) "
	sqlStr = sqlStr & " values('" + CStr(divcd) + "', '" + CStr(appDate) + "', '" + CStr(SELLBIZSECTION_CD) + "', '" + CStr(BUYBIZSECTION_CD) + "', " + CStr(totalSum) + ", " + CStr(supplySum) + ", " + CStr(taxSum) + ", '" + CStr(vatyn) + "', '" + CStr(reguserid) + "', '" + CStr(selluserid) + "', '" + CStr(acc_cd) + "') "
	'response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다..'); opener.location.reload(); opener.focus(); windows.close();</script>"

elseif (mode = "confirminnerorder") then

	sqlStr = " update db_partner.dbo.tbl_InternalOrder "
	sqlStr = sqlStr & " set buyuserid = '" + CStr(session("ssBctId")) + "' "
	sqlStr = sqlStr & " where idx = " + CStr(idx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다..'); opener.location.reload(); opener.focus(); windows.close();</script>"

else
	'
end if




%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
