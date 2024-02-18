<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 입금내역
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim mode
dim jungsanidx, inoutidx, matchprice
dim masteridx
dim totjungsanprice, orgtotmatchedipkumsum
dim tx_amt, matchstate, orgdetailtotmatchedipkumsum
dim matchdetailidx
dim ipkumstatecd, ipkumdate
	mode 			= requestCheckVar(request.Form("mode"),32)
	jungsanidx 		= requestCheckVar(request.Form("jungsanidx"),10)
	inoutidx 		= requestCheckVar(request.Form("inoutidx"),10)
	matchprice 		= requestCheckVar(request.Form("matchprice"),20)
	matchdetailidx 	= requestCheckVar(request.Form("matchdetailidx"),10)

dim i,cnt,sqlStr

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim reguserid
reguserid = session("ssBctId")

if mode="addmatch" then

	'// =======================================================================
	totjungsanprice = -1

	sqlStr = " select top 1 totalsum from [db_shop].[dbo].tbl_fran_meachuljungsan_master "
	sqlStr = sqlStr + " where idx = " + CStr(jungsanidx) + " "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		totjungsanprice = rsget("totalsum")
	end if
	rsget.close


	'// =======================================================================
	masteridx = 0
	orgtotmatchedipkumsum = 0

	sqlStr = " select top 1 idx, totmatchedipkumsum from db_jungsan.dbo.tbl_ipkum_match_master "
	sqlStr = sqlStr + " where jungsanidx = " + CStr(jungsanidx) + " and useyn = 'Y' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		masteridx = rsget("idx")
		orgtotmatchedipkumsum = rsget("totmatchedipkumsum")
	end if
	rsget.close


	'// 초과입금도 가능해야 한다.(환차익)
	'if (matchprice + orgtotmatchedipkumsum) > totjungsanprice then
	'	response.write "<script>alert('입금확인액이 정산액을 초과합니다.'); history.back();</script>"
	'	response.end
	'end if


	'// =======================================================================
	tx_amt = 0
	matchstate = ""
	orgdetailtotmatchedipkumsum = 0


	sqlStr = " select top 1 i.inoutidx, i.TX_AMT, i.matchstate, IsNull(sum(d.matchprice), 0) as totmatchprice " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT i " + vbCrlf
	sqlStr = sqlStr + " 	left join db_jungsan.dbo.tbl_ipkum_match_detail d " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 		and i.inoutidx = d.ipkumidx " + vbCrlf
	sqlStr = sqlStr + " 		and d.ipkummethod = 'BNK' " + vbCrlf
	sqlStr = sqlStr + " 		and d.useyn = 'Y' " + vbCrlf
	sqlStr = sqlStr + " where " + vbCrlf
	sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 	and i.inoutidx = " + CStr(inoutidx) + " " + vbCrlf
	sqlStr = sqlStr + " group by " + vbCrlf
	sqlStr = sqlStr + " 	i.inoutidx, i.TX_AMT, i.matchstate " + vbCrlf
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		tx_amt = rsget("tx_amt")
		matchstate = rsget("matchstate")
		orgdetailtotmatchedipkumsum = rsget("totmatchprice")
	end if
	rsget.close


	if (tx_amt = 0) then
		response.write "<script>alert('잘못된 입금내역입니다.'); history.back();</script>"
		response.end
	end if

	if (matchstate = "Y") then
		response.write "<script>alert('매칭완료된 입금내역입니다.'); history.back();</script>"
		response.end
	end if

	if (tx_amt < (matchprice + orgdetailtotmatchedipkumsum)) then
		response.write "<script>alert('실제 입금내역보다 매칭하려는 금액이 더 큽니다.'); history.back();</script>"
		response.end
	end if


	'// =======================================================================
	if (masteridx = 0) then

		sqlStr = " select * from db_jungsan.dbo.tbl_ipkum_match_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("jungsanidx") = jungsanidx
		rsget("totmatchedipkumsum") = 0
		rsget("reguserid") = reguserid

		rsget.update
		masteridx = rsget("idx")
		rsget.close

	end if


	'// =======================================================================
	sqlStr = " insert into db_jungsan.dbo.tbl_ipkum_match_detail " + vbCrlf
	sqlStr = sqlStr + " (masteridx, ipkummethod, ipkumidx, matchprice, reguserid) " + vbCrlf
	sqlStr = sqlStr + " values( " + vbCrlf
	sqlStr = sqlStr + " 	" + CStr(masteridx) + " " + vbCrlf
	sqlStr = sqlStr + " 	, 'BNK' " + vbCrlf
	sqlStr = sqlStr + " 	, " + CStr(inoutidx) + " " + vbCrlf
	sqlStr = sqlStr + " 	, " + CStr(matchprice) + " " + vbCrlf
	sqlStr = sqlStr + " 	, '" + CStr(reguserid) + "' " + vbCrlf
	sqlStr = sqlStr + " ) " + vbCrlf
	'response.write sqlStr
	rsget.Open sqlStr, dbget, 1

	if (tx_amt <= (orgdetailtotmatchedipkumsum + matchprice)) then

		sqlStr = " update [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT "
		sqlStr = sqlStr + " set matchstate = 'Y' "
		sqlStr = sqlStr + " where inoutidx = " + CStr(inoutidx) + " "
		rsget.Open sqlStr, dbget, 1

	else

		sqlStr = " update [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT "
		sqlStr = sqlStr + " set matchstate = 'H' "
		sqlStr = sqlStr + " where inoutidx = " + CStr(inoutidx) + " "
		rsget.Open sqlStr, dbget, 1

	end if


	sqlStr = " update " + vbCrlf
	sqlStr = sqlStr + " 	m " + vbCrlf
	sqlStr = sqlStr + " set " + vbCrlf
	sqlStr = sqlStr + " 	m.totmatchedipkumsum = T.totmatchprice " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	db_jungsan.dbo.tbl_ipkum_match_master m " + vbCrlf
	sqlStr = sqlStr + " 	join ( " + vbCrlf
	sqlStr = sqlStr + " 		select " + vbCrlf
	sqlStr = sqlStr + " 			masteridx, sum(matchprice) as totmatchprice " + vbCrlf
	sqlStr = sqlStr + " 		from " + vbCrlf
	sqlStr = sqlStr + " 			db_jungsan.dbo.tbl_ipkum_match_detail " + vbCrlf
	sqlStr = sqlStr + " 		where " + vbCrlf
	sqlStr = sqlStr + " 			masteridx = " + CStr(masteridx) + " and useyn = 'Y' " + vbCrlf
	sqlStr = sqlStr + " 		group by " + vbCrlf
	sqlStr = sqlStr + " 			masteridx " + vbCrlf
	sqlStr = sqlStr + " 	) T " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		m.idx = T.masteridx " + vbCrlf
	rsget.Open sqlStr, dbget, 1

	'// 입금일, 입금상태
	ipkumstatecd = ""
	ipkumdate = ""
	sqlStr = " select " + vbCrlf
	sqlStr = sqlStr + " 	max(i.acct_txday) as ipkumdate, sum(d.matchprice) as totipkum " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	db_jungsan.dbo.tbl_ipkum_match_master m " + vbCrlf
	sqlStr = sqlStr + " 	join db_jungsan.dbo.tbl_ipkum_match_detail d " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		m.idx = d.masteridx " + vbCrlf
	sqlStr = sqlStr + " 	join [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT i " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		d.ipkumidx = i.inoutidx " + vbCrlf
	sqlStr = sqlStr + " where " + vbCrlf
	sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 	and m.jungsanidx = " + CStr(jungsanidx) + " " + vbCrlf
	sqlStr = sqlStr + " 	and m.useyn = 'Y' " + vbCrlf
	sqlStr = sqlStr + " 	and d.useyn = 'Y' " + vbCrlf
	''response.write sqlStr
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ipkumdate = rsget("ipkumdate")
		ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

		if (totjungsanprice <= rsget("totipkum")) then
			ipkumstatecd = "9"
		else
			ipkumstatecd = "5"
		end if

	end if
	rsget.close

	if (ipkumstatecd <> "") then
		sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	ipkumstatecd = '" + CStr(ipkumstatecd) + "' "
		sqlStr = sqlStr + " 	, ipkumdate = '" + CStr(ipkumdate) + "' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " + CStr(jungsanidx) + " "
		rsget.Open sqlStr, dbget, 1

		sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
		sqlStr = sqlStr + " set statecd = 7 "
		sqlStr = sqlStr + " where idx = " + CStr(jungsanidx) + " and issuestatecd = 9 and ipkumstatecd = 9 "
		rsget.Open sqlStr, dbget, 1
	end if

	response.write "<script>alert('저장 되었습니다.');</script>"

elseif (mode = "delmatch") then

	'// =======================================================================
	sqlStr = " update db_jungsan.dbo.tbl_ipkum_match_detail "
	sqlStr = sqlStr + " set useyn = 'N' "
	sqlStr = sqlStr + " where idx = " + CStr(matchdetailidx) + " "
	rsget.Open sqlStr, dbget, 1


	'// =======================================================================
	masteridx = 0
	orgtotmatchedipkumsum = 0

	sqlStr = " select top 1 idx, totmatchedipkumsum from db_jungsan.dbo.tbl_ipkum_match_master "
	sqlStr = sqlStr + " where jungsanidx = " + CStr(jungsanidx) + " and useyn = 'Y' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		masteridx = rsget("idx")
	end if
	rsget.close


	'// =======================================================================
	tx_amt = 0
	matchstate = ""
	orgdetailtotmatchedipkumsum = 0


	sqlStr = " select top 1 i.inoutidx, i.TX_AMT, IsNull(sum(d.matchprice), 0) as totmatchprice " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT i " + vbCrlf
	sqlStr = sqlStr + " 	left join db_jungsan.dbo.tbl_ipkum_match_detail d " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 		and i.inoutidx = d.ipkumidx " + vbCrlf
	sqlStr = sqlStr + " 		and d.ipkummethod = 'BNK' " + vbCrlf
	sqlStr = sqlStr + " 		and d.useyn = 'Y' " + vbCrlf
	sqlStr = sqlStr + " where " + vbCrlf
	sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 	and i.inoutidx = " + CStr(inoutidx) + " " + vbCrlf
	sqlStr = sqlStr + " group by " + vbCrlf
	sqlStr = sqlStr + " 	i.inoutidx, i.TX_AMT " + vbCrlf
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		tx_amt = rsget("tx_amt")
		orgdetailtotmatchedipkumsum = rsget("totmatchprice")
	end if
	rsget.close


	'// =======================================================================
	if (orgdetailtotmatchedipkumsum = 0) then

		sqlStr = " update [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT "
		sqlStr = sqlStr + " set matchstate = 'N' "
		sqlStr = sqlStr + " where inoutidx = " + CStr(inoutidx) + " "
		rsget.Open sqlStr, dbget, 1

	elseif (tx_amt > orgdetailtotmatchedipkumsum) then

		sqlStr = " update [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT "
		sqlStr = sqlStr + " set matchstate = 'H' "
		sqlStr = sqlStr + " where inoutidx = " + CStr(inoutidx) + " "
		rsget.Open sqlStr, dbget, 1

	end if


	sqlStr = " update " + vbCrlf
	sqlStr = sqlStr + " 	m " + vbCrlf
	sqlStr = sqlStr + " set " + vbCrlf
	sqlStr = sqlStr + " 	m.totmatchedipkumsum = IsNull(T.totmatchprice, 0) " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	db_jungsan.dbo.tbl_ipkum_match_master m " + vbCrlf
	sqlStr = sqlStr + " 	left join ( " + vbCrlf
	sqlStr = sqlStr + " 		select " + vbCrlf
	sqlStr = sqlStr + " 			masteridx, sum(matchprice) as totmatchprice " + vbCrlf
	sqlStr = sqlStr + " 		from " + vbCrlf
	sqlStr = sqlStr + " 			db_jungsan.dbo.tbl_ipkum_match_detail " + vbCrlf
	sqlStr = sqlStr + " 		where " + vbCrlf
	sqlStr = sqlStr + " 			masteridx = " + CStr(masteridx) + " and useyn = 'Y' " + vbCrlf
	sqlStr = sqlStr + " 		group by " + vbCrlf
	sqlStr = sqlStr + " 			masteridx " + vbCrlf
	sqlStr = sqlStr + " 	) T " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		m.idx = T.masteridx " + vbCrlf
	sqlStr = sqlStr + " where " + vbCrlf
	sqlStr = sqlStr + " 	m.idx = " + CStr(masteridx) + " " + vbCrlf
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	ipkumstatecd = '1' "
	sqlStr = sqlStr + " 	, ipkumdate = NULL "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	idx = " + CStr(jungsanidx) + " "
	rsget.Open sqlStr, dbget, 1

	response.write "<script>alert('삭제 되었습니다.');</script>"

elseif (mode = "dismatch") then

	'// =======================================================================
	sqlStr = " update [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT "
	sqlStr = sqlStr + " set matchstate = 'X' "
	sqlStr = sqlStr + " where inoutidx = " + CStr(inoutidx) + " and IsNull(matchstate, 'N') = 'N' "
	rsget.Open sqlStr, dbget, 1

	response.write "<script>alert('제외 되었습니다.');</script>"

end if

%>
<script language="javascript">
// alert('저장 되었습니다.');
opener.location.reload();
opener.focus();
window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
