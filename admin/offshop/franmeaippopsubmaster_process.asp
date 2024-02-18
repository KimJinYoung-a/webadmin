<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 정산
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

'==============================================================================
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim menupos
dim mode, i, j, sqlStr
dim masteridx, workidx, orgworkidx
dim shopid, yyyymm, deliverpay
dim errMsg
dim submasteridx, subdetailidx

'==============================================================================
menupos		= requestCheckVar(request("menupos"),10)
mode		= requestCheckVar(request("mode"),32)
masteridx	= requestCheckVar(request("masteridx"),10)
workidx		= requestCheckVar(request("workidx"),10)
orgworkidx	= requestCheckVar(request("orgworkidx"),10)

'==============================================================================
if (mode = "insertworkidx") then

	sqlStr = " update db_shop.dbo.tbl_fran_meachuljungsan_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	workidx = " + CStr(workidx) + " "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(masteridx) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

	sqlStr = " update db_storage.dbo.tbl_cartoonbox_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	jungsanidx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(workidx) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

	sqlStr = " update j "
	sqlStr = sqlStr + " set j.invoceidx = c.invoceidx, j.issuestatecd = '0' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_shop.dbo.tbl_fran_meachuljungsan_master j "
	sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_master c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and j.idx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " 		and c.idx = " + CStr(workidx) + " "
	sqlStr = sqlStr + " 		and c.invoceidx is not NULL "
	sqlStr = sqlStr + " 		and j.invoceidx is NULL "
	dbget.Execute sqlStr

elseif (mode = "updateworkidx") then

	sqlStr = " update db_shop.dbo.tbl_fran_meachuljungsan_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	workidx = " + CStr(workidx) + " "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(masteridx) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

	sqlStr = " update db_storage.dbo.tbl_cartoonbox_master "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	jungsanidx = " + CStr(masteridx) + " "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(workidx) + " "
	'response.write "aaaaaaaaaaaa" & sqlStr
	dbget.Execute sqlStr

	if (orgworkidx <> workidx) then
		sqlStr = " update db_storage.dbo.tbl_cartoonbox_master "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	jungsanidx = NULL "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and idx = " + CStr(orgworkidx) + " "
		'response.write "aaaaaaaaaaaa" & sqlStr
		dbget.Execute sqlStr

		sqlStr = " update j "
		sqlStr = sqlStr + " set j.invoceidx = c.invoceidx, j.issuestatecd = '0' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_shop.dbo.tbl_fran_meachuljungsan_master j "
		sqlStr = sqlStr + " 	join db_storage.dbo.tbl_cartoonbox_master c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and j.idx = " + CStr(masteridx) + " "
		sqlStr = sqlStr + " 		and c.idx = " + CStr(workidx) + " "
		sqlStr = sqlStr + " 		and c.invoceidx is not NULL "
		sqlStr = sqlStr + " 		and IsNull(j.invoceidx,0) <> c.invoceidx "
		dbget.Execute sqlStr
	end if

elseif (mode = "addemsprice") then

	'// =======================================================================
    errMsg = ""

	sqlStr = " select top 1 j.*, IsNull(c.delivermethod,'') as delivermethod, IsNull(c.deliverpay,0) as deliverpay "
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_master j"
	sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_cartoonbox_master c "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	j.workidx = c.idx "
	sqlStr = sqlStr + " where j.idx=" + CStr(masteridx)
	sqlStr = sqlStr + " and j.statecd=0"
	rsget.Open sqlStr, dbget, 1

	if  not rsget.EOF  then
		if (rsget("delivermethod") <> "E") then
			errMsg = "운송방법이 EMS일때만 사용가능합니다."
		end if

		if (errMsg = "") and (rsget("deliverpay") = 0) then
			errMsg = "EMS운송비용이 입력되어 있지 않습니다."
		end if

		shopid = rsget("shopid")
		yyyymm = rsget("yyyymm")
		deliverpay = rsget("deliverpay")

	else
		errMsg = "수정중 상태에서만 추가 가능합니다."
	end if

	rsget.Close

	if errMsg <> "" then
		response.write "<script type='text/javascript'>alert('" + errMsg + "');</script>"
		response.write "<script type='text/javascript'>window.close();</script>"
		dbget.close()	:	response.End
	end if

	'// =======================================================================
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = masteridx
	rsget("linkidx") = 0
	rsget("shopid") = shopid
	rsget("code01") = yyyymm
	rsget("code02") = "temp"
	rsget("execdate") = yyyymm + "-01"
	rsget("totalcount") = 0
	rsget("totalsellcash") = 0
	rsget("totalbuycash") = 0
	rsget("totalsuplycash") = 0
	rsget("totalorgsellcash") = 0
	rsget.update
	submasteridx = rsget("idx")
	rsget.close

	'// =======================================================================
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = submasteridx
	rsget("topmasteridx") = masteridx
	rsget("linkbaljucode") = "etc2"			'// 기타
	rsget("linkmastercode") = "0"
	rsget("linkdetailidx") = 0
	rsget("itemgubun") = "00"
	rsget("itemid") = 0
	rsget("itemoption") = "0000"
	rsget("itemname") = "EMS 운송요금"
	rsget("itemoptionname") = ""
	rsget("makerid") = "temp"
	rsget("itemno") = 1
	rsget("sellcash") = deliverpay
	rsget("suplycash") = deliverpay
	rsget("buycash") = deliverpay
	rsget("orgsellcash") = deliverpay

	rsget.update
	subdetailidx = rsget("idx")
	rsget.close

	'// =======================================================================
	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(masteridx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1

	'// =======================================================================
	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

end if

%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
