<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
dim vipone2oneyn, vvipone2oneyn, one2oneyn, michulgoyn, stockoutyn, returnyn, indexno, userid, useyn, SQL, menupos, mode
	indexno 		= requestCheckVar(request("indexno"), 32)
	userid 			= requestCheckVar(request("userid"), 32)
	vipone2oneyn 	= requestCheckVar(request("vipone2oneyn"), 32)
	vvipone2oneyn 	= requestCheckVar(request("vvipone2oneyn"), 1)
	one2oneyn 		= requestCheckVar(request("one2oneyn"), 32)
	michulgoyn 		= requestCheckVar(request("michulgoyn"), 32)
	stockoutyn 		= requestCheckVar(request("stockoutyn"), 32)
	returnyn 		= requestCheckVar(request("returnyn"), 32)
	useyn 			= requestCheckVar(request("useyn"), 32)
	menupos 		= requestCheckVar(request("menupos"), 32)
	mode 			= requestCheckVar(request("mode"), 32)

if (mode = "modify") then
	SQL = " update db_cs.dbo.tbl_cs_board_user " & VbCRLF
	SQL = SQL & "set userid='" & trim(userid) & "' " & VbCRLF
	SQL = SQL & "	, vipone2oneyn='" & trim(vipone2oneyn) & "' " & VbCRLF
	SQL = SQL & "	, vvipone2oneyn='" & trim(vvipone2oneyn) & "' " & VbCRLF
	SQL = SQL & "	, one2oneyn='" & trim(one2oneyn) & "' " & VbCRLF
	SQL = SQL & "	, michulgoyn='" & trim(michulgoyn) & "' " & VbCRLF
	SQL = SQL & "	, stockoutyn='" & trim(stockoutyn) & "' " & VbCRLF
	SQL = SQL & "	, returnyn='" & trim(returnyn) & "' " & VbCRLF
	SQL = SQL & "	, useyn='" & trim(useyn) & "' " & VbCRLF
	SQL = SQL & "	, lastupdate=getdate() " & VbCRLF
	SQL = SQL & " where indexno='"& CStr(indexno)& "'" & VbCRLF

	'response.write SQL & "<Br>"
	dbget.execute SQL

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizevvipone2one") then
	SQL = " exec db_cs.[dbo].[sp_Ten_MyQna_SetChargeID_NEW] 0, 'VV' "

	'response.write SQL & "<Br>"
	dbget.execute SQL

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizevipone2one") then
	SQL = " exec db_cs.[dbo].[sp_Ten_MyQna_SetChargeID_NEW] 0, 'V' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizevipone2onenocharge") then
	SQL = " exec db_cs.[dbo].[sp_Ten_MyQna_SetChargeID_NEW] 0, 'V', 'Y' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizeone2one") then
	SQL = " exec db_cs.[dbo].[sp_Ten_MyQna_SetChargeID_NEW] 0, 'R' "
	rsget.Open SQL, dbget, 1

	call AddCsMemo("","1","system",session("ssBctId"),"1:1��� ��й�" + VbCrlf + "1:1����� ��й�Ǿ����ϴ�.")

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizeone2onenocharge") then
	SQL = " exec db_cs.[dbo].[sp_Ten_MyQna_SetChargeID_NEW] 0, 'R', 'Y' "
	rsget.Open SQL, dbget, 1

	call AddCsMemo("","1","system",session("ssBctId"),"1:1��� ������ ��й�" + VbCrlf + "1:1����� ������ ��й�Ǿ����ϴ�.")

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizemichulgonocharge") then
	SQL = " exec db_cs.[dbo].[sp_Ten_Michulgo_Upche_SetChargeID] 'N' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizemichulgoall") then
	SQL = " exec db_cs.[dbo].[sp_Ten_Michulgo_Upche_SetChargeID] 'Y' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizemichulgoavg") then
	SQL = " exec db_cs.[dbo].[sp_Ten_Michulgo_Upche_SetChargeID] 'A' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizestockoutall") then
	SQL = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] 0 "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizenotreturnnocharge") then
	SQL = " exec db_cs.[dbo].[sp_Ten_NotReturn_Upche_SetChargeID] 'N' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

elseif (mode = "reorganizenotreturnall") then
	SQL = " exec db_cs.[dbo].[sp_Ten_NotReturn_Upche_SetChargeID] 'Y' "
	rsget.Open SQL, dbget, 1

	'// CS���� ������ ������Ʈ
	application("csTimeUserList") = DateAdd("d", -1, now)

else
	response.write "�����ڰ� �����ϴ�"
	dbget.close() : response.end
end if

response.write	"<script type='text/javascript'>" &_
				"	alert('�����Ǿ����ϴ�.'); location.replace('/cscenter/board/boarduserlist.asp?menupos=" & menupos & "'); " &_
				"</script>"

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
