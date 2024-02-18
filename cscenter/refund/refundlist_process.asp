<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbAcademyHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcls.asp" -->

<%
dim ckidx, referer, mode, asid, upfiledate, sitegubun
dim rebankaccount, arrckidx, arrrebankaccount

referer = request.ServerVariables("HTTP_REFERER")

ckidx           	= Trim(request("ckidx"))
arrrebankaccount    = Trim(request("arrrebankaccount"))

mode    = request("mode")
asid    = request("asid")
upfiledate = request("upfiledate")
sitegubun  = request("sitegubun")

dim sqlStr, rowCount, sqlStrFrom, sqlStrTIDX, sqlStrTIDX_dbACADEMYget, sqlStrTIDX_TENSTATUS
dim errcode
dim retURL, paramInfo, retParam, rowVal
dim i

'' **********************************************
'' 환불 유효성 검사
'' - 환불액 합계가 기 결제액보다 큰지 검사
'' - // 주문번호의 구매자, 수령인명과 일치하는지 검사. - 안함.
if (mode="regfile") then

end if

''이체파일 작성일 저장
if (mode="regfile") then

	'==========================================================================
	if (sitegubun = "10x10") then

	    sqlStrFrom = " FROM " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	[db_cs].[dbo].tbl_as_refund_info r " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	JOIN [db_cs].[dbo].tbl_new_as_list a " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	ON " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 		r.asid=a.id " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	JOIN db_log.dbo.tbl_IBK_BANKCD b " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	ON " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 		(Case When r.rebankname='현대스위스상호저축은행' then '상호저축' WHEN r.rebankname='조흥' then '신한' When r.rebankname='외환' then '하나' When r.rebankname='KEB하나' then '하나' else r.rebankname end) = b.BANK_NAME_TEN " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " WHERE " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	a.divcd='A003' " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	and a.currstate='B001' " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	and a.deleteyn='N' " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	and r.asid in (" & ckidx & ") " + VbCrlf

	elseif (sitegubun = "academy") then

	    sqlStrFrom = " FROM " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	[ACADEMYDB].[db_academy].[dbo].tbl_academy_as_refund_info r " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	JOIN [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	ON " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 		r.asid=a.id " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	JOIN db_log.dbo.tbl_IBK_BANKCD b " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	ON " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 		(Case When r.rebankname='조흥' then '신한' When r.rebankname='외환' then '하나' When r.rebankname='KEB하나' then '하나' else r.rebankname end) = b.BANK_NAME_TEN " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " WHERE " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	a.divcd='A003' " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	and a.currstate='B001' " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	and a.deleteyn='N' " + VbCrlf
	    sqlStrFrom = sqlStrFrom + " 	and r.asid in (" & ckidx & ") " + VbCrlf

	else

		'에러

	end if



	'==========================================================================
    sqlStr = " select convert(varchar(19),getdate(),21) as upfiledate, count(r.asid) as cnt "
    sqlStr = sqlStr + sqlStrFrom

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly
    if Not rsget.Eof then
        upfiledate = rsget("upfiledate")
        rowCount = rsget("cnt")

        if (rowCount > 3000) then
        	rowCount = 3000
        end if
    end if
    rsget.Close

    if (rowCount = "") then
    	rowCount = 0
    end if



	'==========================================================================
	'이체데이타 작성

	arrckidx			= Split(ckidx, ",")
	arrrebankaccount	= Split(arrrebankaccount, ",")

''	if (sitegubun = "10x10") then
''		'텐바이텐만 복호화한다. ==> 환불 완료건에 대해서만 암호화 함.
''		for i = 0 to UBound(arrrebankaccount)
''			if (Trim(arrrebankaccount(i)) <> "") then
''				sqlStr = " update " + VbCrlf
''				sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_as_refund_info " + VbCrlf
''				sqlStr = sqlStr + " set " + VbCrlf
''				sqlStr = sqlStr + " 	rebankaccount = '" & Trim(arrrebankaccount(i)) & "' " + VbCrlf
''				sqlStr = sqlStr + " where " + VbCrlf
''				sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
''				sqlStr = sqlStr + " 	and asid = " & Trim(arrckidx(i)) & " " + VbCrlf
''				sqlStr = sqlStr + " 	and encmethod = 'TBT' " + VbCrlf
''				dbget.Execute sqlStr
''			end if
''		next
''	end if

    sqlStr = " INSERT INTO db_log.dbo.tbl_IBK_ERP_ICHE_DATA( " + VbCrlf
    sqlStr = sqlStr + "  	SITE_NO " + VbCrlf
    sqlStr = sqlStr + " 	, FL_DATE " + VbCrlf
    sqlStr = sqlStr + " 	, FL_TIME " + VbCrlf
    sqlStr = sqlStr + " 	, FL_CNT " + VbCrlf
    sqlStr = sqlStr + " 	, FL_SEQ " + VbCrlf
    sqlStr = sqlStr + " 	, SEND_GB " + VbCrlf
    sqlStr = sqlStr + " 	, IN_BANK_CD " + VbCrlf
    sqlStr = sqlStr + " 	, IN_ACCT_NO " + VbCrlf
    sqlStr = sqlStr + " 	, TRAN_AMT " + VbCrlf
    sqlStr = sqlStr + " 	, PRE_RECI_MAN " + VbCrlf
    sqlStr = sqlStr + " 	, IN_PRT " + VbCrlf
    sqlStr = sqlStr + " 	, OUT_PRT " + VbCrlf
    sqlStr = sqlStr + " 	, REG_DATE " + VbCrlf
    sqlStr = sqlStr + " 	, TEN_CSID " + VbCrlf
    sqlStr = sqlStr + " 	, SITEGUBUN " + VbCrlf
    sqlStr = sqlStr + " ) " + VbCrlf
    sqlStr = sqlStr + " SELECT TOP 3000  " + VbCrlf
    sqlStr = sqlStr + " 	'2118700620' " + VbCrlf
    sqlStr = sqlStr + " 	, Replace(Replace(convert(varchar(10),'" & upfiledate & "',21),'-',''),' ','') as FL_DATE " + VbCrlf
    sqlStr = sqlStr + " 	, Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') as FL_TIME " + VbCrlf
    sqlStr = sqlStr + " 	, ROW_NUMBER() OVER(ORDER BY r.asid DESC) as FL_CNT " + VbCrlf
    sqlStr = sqlStr + " 	, 1 as FL_SEQ " + VbCrlf	 																					'3000건 넘을경우 증가
    sqlStr = sqlStr + " 	, 4 as SEND_GB " + VbCrlf																					    '조회후 환불
    sqlStr = sqlStr + " 	, b.EB_BANK_CD as IN_BANK_CD " + VbCrlf
    ''''''''''sqlStr = sqlStr + " 	, Replace(Replace(r.rebankaccount,' ',''),'-','') as IN_ACCT_NO " + VbCrlf
    sqlStr = sqlStr + " 	, convert(varchar(20), Replace(Replace( (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), r.rebankaccount) WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), r.rebankaccount) ELSE r.rebankaccount END ) ,' ',''),'-','')) as IN_ACCT_NO " + VbCrlf
    sqlStr = sqlStr + " 	, r.refundrequire as TRAN_AMT " + VbCrlf
    sqlStr = sqlStr + " 	, convert(varchar(32),r.rebankownername) as PRE_RECI_MAN " + VbCrlf
    sqlStr = sqlStr + " 	, '텐바이텐' as IN_PRT " + VbCrlf
    sqlStr = sqlStr + " 	, convert(varchar(18),r.rebankownername) as OUT_PRT " + VbCrlf
    sqlStr = sqlStr + " 	, Replace(Replace(Replace(convert(varchar(20),getdate(),20),'-',''),':',''),' ','') as REG_DATE " + VbCrlf
    sqlStr = sqlStr + " 	, r.asid " + VbCrlf
    sqlStr = sqlStr + " 	, '" & sitegubun & "' " + VbCrlf
	sqlStr = sqlStr + sqlStrFrom

	'TODO : GetRefundRequireList 와 위 쿼리의 정렬순서는 일치해야 한다.
	sqlStr = sqlStr + " ORDER BY a.id asc " + VbCrlf

	dbget.Execute sqlStr

''	if (sitegubun = "10x10") then
''		'텐바이텐만 복호화했던 데이타를 날린다.
''		sqlStr = " update " + VbCrlf
''		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_as_refund_info " + VbCrlf
''		sqlStr = sqlStr + " set " + VbCrlf
''		sqlStr = sqlStr + " 	rebankaccount = '' " + VbCrlf
''		sqlStr = sqlStr + " where " + VbCrlf
''		sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
''		sqlStr = sqlStr + " 	and asid in (" & ckidx & ") " + VbCrlf
''		sqlStr = sqlStr + " 	and encmethod = 'TBT' " + VbCrlf
''		dbget.Execute sqlStr
''	end if

	'==========================================================================
	'입력된 내역 표시
	if (sitegubun = "10x10") then

	    sqlStr = " UPDATE" + VbCrlf
	    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_as_refund_info" + VbCrlf
	    sqlStr = sqlStr + " SET" + VbCrlf
	    sqlStr = sqlStr + " 	upfiledate='" & upfiledate & "' " + VbCrlf
	    sqlStr = sqlStr + " 	, IBK_TIDX=T.TIDX" + VbCrlf
	    sqlStr = sqlStr + " FROM" + VbCrlf
	    sqlStr = sqlStr + " 	db_log.dbo.tbl_IBK_ERP_ICHE_DATA T" + VbCrlf
	    sqlStr = sqlStr + " WHERE" + VbCrlf
	    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_as_refund_info.asid in (" + ckidx + ")" + VbCrlf
	    sqlStr = sqlStr + " 	and [db_cs].[dbo].tbl_as_refund_info.asid=T.TEN_CSID" + VbCrlf
	    sqlStr = sqlStr + " 	and T.SITE_NO='2118700620'" + VbCrlf
	    sqlStr = sqlStr + " 	and T.FL_DATE=Replace(Replace(convert(varchar(10),'" & upfiledate & "',21),'-',''),' ','') " + VbCrlf
	    sqlStr = sqlStr + " 	and T.FL_TIME=Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') " + VbCrlf
	    sqlStr = sqlStr + " 	and T.FL_SEQ=1" + VbCrlf
	    sqlStr = sqlStr + " 	and T.TEN_STATUS=0" + VbCrlf
	    sqlStr = sqlStr + " 	and IsNull(T.SITEGUBUN, '10x10') = '10x10' " + VbCrlf

	    dbget.Execute sqlStr

	elseif (sitegubun = "academy") then

	    sqlStr = " UPDATE" + VbCrlf
	    sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_as_refund_info " + VbCrlf
	    sqlStr = sqlStr + " SET" + VbCrlf
	    sqlStr = sqlStr + " 	upfiledate='" & upfiledate & "' " + VbCrlf
	    sqlStr = sqlStr + " 	, IBK_TIDX=T.TIDX" + VbCrlf
	    sqlStr = sqlStr + " FROM" + VbCrlf
	    sqlStr = sqlStr + " 	[TENDB].db_log.dbo.tbl_IBK_ERP_ICHE_DATA T" + VbCrlf
	    sqlStr = sqlStr + " WHERE" + VbCrlf
	    sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_as_refund_info.asid in (" + ckidx + ")" + VbCrlf
	    sqlStr = sqlStr + " 	and [db_academy].[dbo].tbl_academy_as_refund_info.asid=T.TEN_CSID" + VbCrlf
	    sqlStr = sqlStr + " 	and T.SITE_NO='2118700620'" + VbCrlf
	    sqlStr = sqlStr + " 	and T.FL_DATE=Replace(Replace(convert(varchar(10),'" & upfiledate & "',21),'-',''),' ','') " + VbCrlf
	    sqlStr = sqlStr + " 	and T.FL_TIME=Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') " + VbCrlf
	    sqlStr = sqlStr + " 	and T.FL_SEQ=1" + VbCrlf
	    sqlStr = sqlStr + " 	and T.TEN_STATUS=0" + VbCrlf
	    sqlStr = sqlStr + " 	and IsNull(T.SITEGUBUN, '10x10') = 'academy' " + VbCrlf

	    dbACADEMYget.Execute sqlStr

'response.write sqlStr
'dbACADEMYget.close
'dbget.close
'response.end

	else

		'에러

	end if



	'    sqlStr = " db_cs.dbo.sp_TEN_CS_ASRefundFile_ArrayProc"
	'
	'    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	'			,Array("@idxArray"	    , adVarchar	, adParamInput	, 	4000  , ckidx) _
	'			,Array("@UPFILEDATE"	, adVarchar	, adParamOutput	,   19    , 0) _
	'	)
	'
	'    if (Len(ckidx)>0) then
	'        retParam = fnExecSPOutput(sqlStr, paramInfo)
	'        rowCount = GetValue(retParam,"@RETURN_VALUE")
	'        upfiledate = GetValue(retParam,"@UPFILEDATE")
	'    end if




    if (rowCount>0) then
        retURL = "/cscenter/refund/refundlist.asp?menupos=972&upfiledate="&upfiledate&"&upfilestate=uploaded&sitegubun=" + sitegubun
        response.write "<script language='javascript'>alert('환불 이체 파일이 ("&rowCount&"건) 작성 되었습니다.');"
        response.write "location.replace('" & retURL & "');"
    else
        response.write "<script language='javascript'>alert('!! 이체 파일이 작성되지 않앗습니다.');"
        response.write "location.replace('" & referer & "');"
    end if

    response.write "</script>"



    ''파일로 작성시..
elseif (mode="regfileOLD") then
    sqlStr = " DECLARE @UPFILEDATE varchar(19)" + VbCrlf
    sqlStr = sqlStr + " set @UPFILEDATE=convert(varchar(19),getdate(),21)" + VbCrlf

	if (sitegubun = "10x10") then

	    sqlStr = sqlStr + " update [db_cs].[dbo].tbl_as_refund_info" + VbCrlf
	    sqlStr = sqlStr + " set upfiledate=@UPFILEDATE" + VbCrlf
	    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a"
	    sqlStr = sqlStr + " where [db_cs].[dbo].tbl_as_refund_info.asid in (" + ckidx + ")" + VbCrlf
	    sqlStr = sqlStr + " and [db_cs].[dbo].tbl_as_refund_info.asid=a.id" + VbCrlf
	    sqlStr = sqlStr + " and a.divcd='A003'" + VbCrlf
	    sqlStr = sqlStr + " and a.currstate='B001'" + VbCrlf
	    sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
	    sqlStr = sqlStr + " and [db_cs].[dbo].tbl_as_refund_info.returnmethod='R007'" + VbCrlf

        if (Len(ckidx)>0) then
	        dbget.Execute sqlStr, rowCount
	    end if

	elseif (sitegubun = "academy") then

	    sqlStr = sqlStr + " update [db_academy].[dbo].tbl_academy_as_refund_info" + VbCrlf
	    sqlStr = sqlStr + " set upfiledate=@UPFILEDATE" + VbCrlf
	    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list a"
	    sqlStr = sqlStr + " where [db_academy].[dbo].tbl_academy_as_refund_info.asid in (" + ckidx + ")" + VbCrlf
	    sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_as_refund_info.asid=a.id" + VbCrlf
	    sqlStr = sqlStr + " and a.divcd='A003'" + VbCrlf
	    sqlStr = sqlStr + " and a.currstate='B001'" + VbCrlf
	    sqlStr = sqlStr + " and a.deleteyn='N'" + VbCrlf
	    sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_as_refund_info.returnmethod='R007'" + VbCrlf

        if (Len(ckidx)>0) then
	        dbACADEMYget.Execute sqlStr, rowCount
	    end if

	else

		'에러

	end if

elseif (mode="rollbackfile") then
    ''작성이전 변경. - tbl_IBK_ERP_ICHE_DATA 상태 확인..
    dim TIDX, PROC_YN, EB_USED, ERR_MSG
    TIDX = 0
    PROC_YN = ""
    EB_USED = ""

    sqlStr = " select TIDX,IsNULL(PROC_YN,'') as PROC_YN, IsNULL(EB_USED,'') as EB_USED, IsNull(ERR_MSG, '') as ERR_MSG"
    sqlStr = sqlStr + " from db_log.dbo.tbl_IBK_ERP_ICHE_DATA"
    sqlStr = sqlStr + " where TEN_CSID=" + CStr(asid)

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly
    if Not rsget.Eof then
        TIDX = rsget("TIDX")
        PROC_YN = Trim(rsget("PROC_YN"))
        EB_USED = Trim(rsget("EB_USED"))
        ERR_MSG = Trim(rsget("ERR_MSG"))
    end if
    rsget.Close

    if (ERR_MSG = "자료가져온후 삭제됨") then
        '// 롤백 허용
    elseif (EB_USED="Y") or (PROC_YN<>"") then
        response.write "<script>alert('결제 요청중 또는 처리 내역이므로 변경 할 수 없습니다.');</script>"
        response.write "<script>history.back();</script>"
        dbget.Close : response.end
    end if



    ''트랜잭션
    on Error Resume Next
    dbget.BeginTrans
    dbACADEMYget.BeginTrans



    if (TIDX<>0) then
        sqlStr = " delete from db_log.dbo.tbl_IBK_ERP_ICHE_DATA"
        sqlStr = sqlStr + " where TIDX="&TIDX

        dbget.Execute sqlStr
    end if

    if (sitegubun = "10x10") then

	    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info" + VbCrlf
	    sqlStr = sqlStr + " set upfiledate=NULL" + VbCrlf
	    sqlStr = sqlStr + " where asid=" + CStr(asid)

	    dbget.Execute sqlStr, rowCount

    elseif (sitegubun = "academy") then

	    sqlStr = " update [db_academy].[dbo].tbl_academy_as_refund_info" + VbCrlf
	    sqlStr = sqlStr + " set upfiledate=NULL" + VbCrlf
	    sqlStr = sqlStr + " where asid=" + CStr(asid)

	    dbACADEMYget.Execute sqlStr, rowCount

    else

	end if



    IF Err then
        dbget.RollBackTrans
        dbACADEMYget.RollBackTrans
    ELSE
        dbget.CommitTrans
        dbACADEMYget.CommitTrans
    end IF
    on Error Goto 0

elseif (mode="finisharray") then
    response.write ckidx

'    On Error Resume Next
'    dbget.beginTrans
'    dbACADEMYget.beginTrans



'    If (Err.Number = 0)Then
'        errcode = "001"
        '' CS Master 완료처리

	    if (sitegubun = "10x10") then

	        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
	        sqlStr = sqlStr + " set finishuser='" + session("ssBctid") + "'" + VbCrlf
	        sqlStr = sqlStr + " , contents_finish='대량이체 환불 처리'" + VbCrlf
	        sqlStr = sqlStr + " , finishdate=getdate()" + VbCrlf
	        sqlStr = sqlStr + " , currstate='B007'" + VbCrlf
	        sqlStr = sqlStr + " , opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '" & "환불(취소) 완료: ' + convert(varchar,convert(int,r.refundrequire))  else '" & VbCrlf & "환불(취소) 완료: ' + convert(varchar,convert(int,r.refundrequire))  End )" + VbCrlf
	        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info r"
	        sqlStr = sqlStr + " where [db_cs].[dbo].tbl_new_as_list.id in (" + ckidx + ")" + VbCrlf
	        sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.id=r.asid"
	        sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.divcd='A003'" + VbCrlf
	        sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.currstate='B001'" + VbCrlf
	        sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.deleteyn='N'" + VbCrlf

	        dbget.Execute sqlStr, rowCount

	    elseif (sitegubun = "academy") then

	        sqlStr = " update [db_academy].[dbo].tbl_academy_as_list" + VbCrlf
	        sqlStr = sqlStr + " set finishuser='" + session("ssBctid") + "'" + VbCrlf
	        sqlStr = sqlStr + " , contents_finish='대량이체 환불 처리'" + VbCrlf
	        sqlStr = sqlStr + " , finishdate=getdate()" + VbCrlf
	        sqlStr = sqlStr + " , currstate='B007'" + VbCrlf
	        sqlStr = sqlStr + " , opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '" & "환불(취소) 완료: ' + convert(varchar,convert(int,r.refundrequire))  else '" & VbCrlf & "환불(취소) 완료: ' + convert(varchar,convert(int,r.refundrequire))  End )" + VbCrlf
	        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_refund_info r"
	        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_academy_as_list.id in (" + ckidx + ")" + VbCrlf
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_as_list.id=r.asid"
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_as_list.divcd='A003'" + VbCrlf
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_as_list.currstate='B001'" + VbCrlf
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_as_list.deleteyn='N'" + VbCrlf

	        dbACADEMYget.Execute sqlStr, rowCount

	    else

		end if

'    end if


'    If (Err.Number = 0) and (rowCount>0) Then
'        errcode = "002"
        '' CS Master 완료처리

	    if (sitegubun = "10x10") then

		    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info" + VbCrlf
	        sqlStr = sqlStr + " set refundresult=refundrequire" + VbCrlf
	        sqlStr = sqlStr + " where asid in (" + ckidx + ")" + VbCrlf

		    dbget.Execute sqlStr

	    elseif (sitegubun = "academy") then

		    sqlStr = " update [db_academy].[dbo].tbl_academy_as_refund_info" + VbCrlf
	        sqlStr = sqlStr + " set refundresult=refundrequire" + VbCrlf
	        sqlStr = sqlStr + " where asid in (" + ckidx + ")" + VbCrlf

		    dbACADEMYget.Execute sqlStr

	    else

		end if

'    end if
'
'     If (Err.Number = 0) and (rowCount>0) Then
'        errcode = "003"

	    if (sitegubun = "10x10") then

	        ''sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "

	        '' 2015/08/17 수정
            sqlStr = "Insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran "
            sqlStr = sqlStr + " (recipient_num, callback, msg_status, date_client_req, content,  service_type, broadcast_yn ) "

	        sqlStr = sqlStr + " select m.buyhp," + VbCrlf
	        sqlStr = sqlStr + " '1644-6030'," + VbCrlf
	        sqlStr = sqlStr + " '1'," + VbCrlf
	        sqlStr = sqlStr + " getdate()," + VbCrlf
	        sqlStr = sqlStr + " '[텐바이텐] 고객님 ' +  convert(varchar,convert(int,r.refundrequire))  + '원 환불이 완료되었습니다. 즐거운 하루 되세요.'" + VbCrlf
	        sqlStr = sqlStr + " ,'0','N'"       ''new_info_SMS
	        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a," + VbCrlf
	        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m," + VbCrlf
	        sqlStr = sqlStr + "  [db_cs].[dbo].tbl_as_refund_info r" + VbCrlf
	        sqlStr = sqlStr + " where  a.id in (" + ckidx + ")" + VbCrlf
	        sqlStr = sqlStr + " and a.id=r.asid" + VbCrlf
	        sqlStr = sqlStr + " and a.orderserial=m.orderserial" + VbCrlf
	        sqlStr = sqlStr + " and a.divcd='A003'" + VbCrlf
	        sqlStr = sqlStr + " and a.deleteyn='N' " + VbCrlf
	        sqlStr = sqlStr + " and a.finishdate>=convert(varchar(10),getdate(),21)" ''2017/08/09 추가 금일 완료된것만 발송.
	        sqlStr = sqlStr + " and r.returnmethod='R007'" + VbCrlf

	    	dbget.Execute sqlStr

	    elseif (sitegubun = "academy") then

	        ''sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "

	        '' 2015/08/17 수정
            sqlStr = "Insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran "
            sqlStr = sqlStr + " (recipient_num, callback, msg_status, date_client_req, content,  service_type, broadcast_yn ) "

	        sqlStr = sqlStr + " select m.buyhp," + VbCrlf
	        sqlStr = sqlStr + " '02-741-9070'," + VbCrlf
	        sqlStr = sqlStr + " '1'," + VbCrlf
	        sqlStr = sqlStr + " getdate()," + VbCrlf
	        sqlStr = sqlStr + " '[아카데미] 고객님 ' +  convert(varchar,convert(int,r.refundrequire))  + '원 환불이 완료되었습니다. 즐거운 하루 되세요.'" + VbCrlf
	        sqlStr = sqlStr + " ,'0','N'"       ''new_info_SMS
	        sqlStr = sqlStr + " from [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a," + VbCrlf
	        sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master m," + VbCrlf
	        sqlStr = sqlStr + "  [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_refund_info r" + VbCrlf
	        sqlStr = sqlStr + " where  a.id in (" + ckidx + ")" + VbCrlf
	        sqlStr = sqlStr + " and a.id=r.asid" + VbCrlf
	        sqlStr = sqlStr + " and a.orderserial=m.orderserial" + VbCrlf
	        sqlStr = sqlStr + " and a.divcd='A003'" + VbCrlf
	        sqlStr = sqlStr + " and a.deleteyn='N' " + VbCrlf
	        sqlStr = sqlStr + " and a.finishdate>=convert(varchar(10),getdate(),21)" ''2017/08/09 추가 금일 완료된것만 발송.
	        sqlStr = sqlStr + " and r.returnmethod='R007'" + VbCrlf

	    ''	dbget.Execute sqlStr

	    else

		end if


        '' db_log.dbo.tbl_IBK_ERP_ICHE_DATA 에서 삭제
        sqlStr = " update db_log.dbo.tbl_IBK_ERP_ICHE_DATA"
        sqlStr = sqlStr & " SET IN_ACCT_NO=''"
        sqlStr = sqlStr & " where TEN_CSID in (" + ckidx + ")" + VbCrlf
        sqlStr = sqlStr & " and sitegubun = '"&sitegubun&"'" + VbCrlf

        dbget.Execute sqlStr

'     end if

'    If (Err.Number = 0) and (ScanErr="") Then
'        dbget.CommitTrans
'        dbACADEMYget.CommitTrans
'    Else
'        dbget.RollBackTrans
'        dbACADEMYget.RollBackTrans
'        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + ")" + Chr(34) + ")</script>"
'        'response.write "<script>history.back()</script>"
'        dbget.close()	:	response.End
'    End If
'    on error Goto 0

elseif (mode="finishfile") then
    '' 작성일로 완료처리
    response.write upfiledate

'    On Error Resume Next
'    dbget.beginTrans
'    dbACADEMYget.beginTrans
'
'    If (Err.Number = 0)Then
'        errcode = "001"
		'IBK 완료

	    if (sitegubun = "10x10") then

			sqlStrTIDX = " select K.TIDX, K.TEN_CSID " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " from db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	Join db_cs.dbo.tbl_new_as_list a " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	on K.TEN_CSID=a.id " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	Join [db_cs].[dbo].tbl_as_refund_info r " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	on a.id=r.asid " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	and r.upfiledate='" & upfiledate & "' " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " where K.SITE_NO='2118700620' " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " and K.PROC_YN='Y' " + VbCrlf
			''sqlStrTIDX = sqlStrTIDX + " and K.PROC_DATE<>'' " + VbCrlf												'20090616 추가// 20110919제거
			sqlStrTIDX = sqlStrTIDX + " and K.FL_DATE=Replace(convert(varchar(10),'" & upfiledate & "',21),'-','')  " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " and K.FL_TIME=Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " and IsNull(K.SITEGUBUN, '10x10') = '10x10' " + VbCrlf

	    elseif (sitegubun = "academy") then

			sqlStrTIDX = " select K.TIDX, K.TEN_CSID " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " from db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	Join [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	on K.TEN_CSID=a.id " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	Join [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_refund_info r " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	on a.id=r.asid " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " 	and r.upfiledate='" & upfiledate & "' " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " where K.SITE_NO='2118700620' " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " and K.PROC_YN='Y' " + VbCrlf
			''sqlStrTIDX = sqlStrTIDX + " and K.PROC_DATE<>'' " + VbCrlf												'20090616 추가
			sqlStrTIDX = sqlStrTIDX + " and K.FL_DATE=Replace(convert(varchar(10),'" & upfiledate & "',21),'-','')  " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " and K.FL_TIME=Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') " + VbCrlf
			sqlStrTIDX = sqlStrTIDX + " and IsNull(K.SITEGUBUN, '10x10') = 'academy' " + VbCrlf

			sqlStrTIDX_dbACADEMYget = " select K.TIDX, K.TEN_CSID " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " from [TENDB].db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " 	Join [db_academy].[dbo].tbl_academy_as_list a " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " 	on K.TEN_CSID=a.id " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " 	Join [db_academy].[dbo].tbl_academy_as_refund_info r " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " 	on a.id=r.asid " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " 	and r.upfiledate='" & upfiledate & "' " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " where K.SITE_NO='2118700620' " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " and K.PROC_YN='Y' " + VbCrlf
			''sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " and K.PROC_DATE<>'' " + VbCrlf					'20090616 추가
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " and K.FL_DATE=Replace(convert(varchar(10),'" & upfiledate & "',21),'-','')  " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " and K.FL_TIME=Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') " + VbCrlf
			sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " and IsNull(K.SITEGUBUN, '10x10') = 'academy' " + VbCrlf

	    else

		end if



		'고객센터에서 완료처리한것은 SMS/리펀드정보 별도로 저장할 필요가 없으므로 제외
		sqlStr = " update db_log.dbo.tbl_IBK_ERP_ICHE_DATA " + VbCrlf
		sqlStr = sqlStr + " set TEN_STATUS=1 " + VbCrlf
		sqlStr = sqlStr + " where TIDX in (select T.TIDX from (" & sqlStrTIDX & " and K.TEN_STATUS=0 and a.currstate='B001') T)  " + VbCrlf
		dbget.Execute sqlStr, rowCount

		if (rowCount < 1) then
			rowVal = -1
		else
			rowVal = rowCount
		end if

		'sqlStrTIDX = sqlStrTIDX + " and K.TEN_STATUS=1 " + VbCrlf
		'sqlStrTIDX_dbACADEMYget = sqlStrTIDX_dbACADEMYget + " and K.TEN_STATUS=1 " + VbCrlf

'	end if
'
'    If (Err.Number = 0) and (rowCount>0) Then
'        errcode = "002"
		'CS 마스타

	    if (sitegubun = "10x10") then

			sqlStr = " update db_cs.dbo.tbl_new_as_list " + VbCrlf
			sqlStr = sqlStr + " set currstate='B007' " + VbCrlf
			sqlStr = sqlStr + " , finishuser='" & session("ssBctid") & "' " + VbCrlf
			sqlStr = sqlStr + " , contents_finish='E-Branch 환불처리' " + VbCrlf
			sqlStr = sqlStr + " , finishdate=getdate() " + VbCrlf
			sqlStr = sqlStr + " , opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '환불(취소) 완료: ' + convert(varchar,convert(int,T.TRAN_AMT))  else char(13) + '환불(취소) 완료: ' + convert(varchar,convert(int,T.TRAN_AMT))  End ) " + VbCrlf
			sqlStr = sqlStr + " from ( " + VbCrlf
			sqlStr = sqlStr + " 	select K.TEN_CSID, K.TRAN_AMT " + VbCrlf
			sqlStr = sqlStr + " 	from db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStr = sqlStr + " 	where K.TIDX in (select T.TIDX from (" & sqlStrTIDX & ") T) " + VbCrlf
			sqlStr = sqlStr + " ) T	 " + VbCrlf
			sqlStr = sqlStr + " where id=T.TEN_CSID " + VbCrlf
			sqlStr = sqlStr + " and divcd='A003' " + VbCrlf
			sqlStr = sqlStr + " and deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " and currstate='B001' " + VbCrlf

			dbget.Execute sqlStr

	    elseif (sitegubun = "academy") then

			sqlStr = " update [db_academy].[dbo].tbl_academy_as_list " + VbCrlf
			sqlStr = sqlStr + " set currstate='B007' " + VbCrlf
			sqlStr = sqlStr + " , finishuser='" & session("ssBctid") & "' " + VbCrlf
			sqlStr = sqlStr + " , contents_finish='E-Branch 환불처리' " + VbCrlf
			sqlStr = sqlStr + " , finishdate=getdate() " + VbCrlf
			sqlStr = sqlStr + " , opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '환불(취소) 완료: ' + convert(varchar,convert(int,T.TRAN_AMT))  else char(13) + '환불(취소) 완료: ' + convert(varchar,convert(int,T.TRAN_AMT))  End ) " + VbCrlf
			sqlStr = sqlStr + " from ( " + VbCrlf
			sqlStr = sqlStr + " 	select K.TEN_CSID, K.TRAN_AMT " + VbCrlf
			sqlStr = sqlStr + " 	from [TENDB].db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStr = sqlStr + " 	where K.TIDX in (select T.TIDX from (" & sqlStrTIDX_dbACADEMYget & ") T) " + VbCrlf
			sqlStr = sqlStr + " ) T	 " + VbCrlf
			sqlStr = sqlStr + " where id=T.TEN_CSID " + VbCrlf
			sqlStr = sqlStr + " and divcd='A003' " + VbCrlf
			sqlStr = sqlStr + " and deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " and currstate='B001' " + VbCrlf

			dbACADEMYget.Execute sqlStr

	    else

		end if

'	end if
'
'    If (Err.Number = 0) and (rowCount>0) Then
'        errcode = "003"
        '' CS RefundInfo 완료

	    if (sitegubun = "10x10") then

			sqlStr = " update [db_cs].[dbo].tbl_as_refund_info " + VbCrlf
			sqlStr = sqlStr + " set refundresult=T.TRAN_AMT " + VbCrlf
			sqlStr = sqlStr + " from ( " + VbCrlf
			sqlStr = sqlStr + " 	select K.TEN_CSID, K.TRAN_AMT " + VbCrlf
			sqlStr = sqlStr + " 	from db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStr = sqlStr + " 	where K.TIDX in (select T.TIDX from (" & sqlStrTIDX & ") T) " + VbCrlf
			sqlStr = sqlStr + " ) T " + VbCrlf
			sqlStr = sqlStr + " where [db_cs].[dbo].tbl_as_refund_info.asid=T.TEN_CSID " + VbCrlf

			dbget.Execute sqlStr

	    elseif (sitegubun = "academy") then

			sqlStr = " update [db_academy].[dbo].tbl_academy_as_refund_info " + VbCrlf
			sqlStr = sqlStr + " set refundresult=T.TRAN_AMT " + VbCrlf
			sqlStr = sqlStr + " from ( " + VbCrlf
			sqlStr = sqlStr + " 	select K.TEN_CSID, K.TRAN_AMT " + VbCrlf
			sqlStr = sqlStr + " 	from [TENDB].db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
			sqlStr = sqlStr + " 	where K.TIDX in (select T.TIDX from (" & sqlStrTIDX_dbACADEMYget & ") T) " + VbCrlf
			sqlStr = sqlStr + " ) T " + VbCrlf
			sqlStr = sqlStr + " where [db_academy].[dbo].tbl_academy_as_refund_info.asid=T.TEN_CSID " + VbCrlf

			dbACADEMYget.Execute sqlStr

	    else

		end if

'    end if
'
'     If (Err.Number = 0) and (rowCount>0) Then
'        errcode = "004"

	    if (sitegubun = "10x10") then

	        ''sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "

	        '' 2015/08/17 수정
            sqlStr = "Insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran "
            sqlStr = sqlStr + " (recipient_num, callback, msg_status, date_client_req, content,  service_type, broadcast_yn ) "

	        sqlStr = sqlStr + " select m.buyhp," + VbCrlf
	        sqlStr = sqlStr + " '1644-6030'," + VbCrlf
	        sqlStr = sqlStr + " '1'," + VbCrlf
	        sqlStr = sqlStr + " getdate()," + VbCrlf
	        'sqlStr = sqlStr + " '[텐바이텐]고객님 ' +  convert(varchar,convert(int,r.refundrequire))  + '원 환불이 5월 25일자로 완료되었습니다. 즐거운 하루 되세요.'" + VbCrlf
	        sqlStr = sqlStr + " '[텐바이텐] 고객님 ' +  convert(varchar,convert(int,r.refundrequire))  + '원 환불이 완료되었습니다. 즐거운 하루 되세요.'" + VbCrlf
	        sqlStr = sqlStr + " ,'0','N'"       ''new_info_SMS
	        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a," + VbCrlf
	        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m," + VbCrlf
	        sqlStr = sqlStr + "  [db_cs].[dbo].tbl_as_refund_info r" + VbCrlf
	        sqlStr = sqlStr + " where  a.id in (select T.TEN_CSID from (" & sqlStrTIDX & ") T)" + VbCrlf
	        sqlStr = sqlStr + " and a.id=r.asid" + VbCrlf
	        sqlStr = sqlStr + " and a.orderserial=m.orderserial" + VbCrlf
	        sqlStr = sqlStr + " and a.divcd='A003'" + VbCrlf
	        sqlStr = sqlStr + " and a.deleteyn='N' " + VbCrlf
	        sqlStr = sqlStr + " and a.finishdate>=convert(varchar(10),getdate(),21)" ''2017/08/09 추가 금일 완료된것만 발송.
	        sqlStr = sqlStr + " and r.returnmethod='R007'" + VbCrlf

	    	dbget.Execute sqlStr

	    elseif (sitegubun = "academy") then

	        ''sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "

	        '' 2015/08/17 수정
            sqlStr = "Insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran "
            sqlStr = sqlStr + " (recipient_num, callback, msg_status, date_client_req, content,  service_type, broadcast_yn ) "

	        sqlStr = sqlStr + " select m.buyhp," + VbCrlf
	        sqlStr = sqlStr + " '02-741-9070'," + VbCrlf
	        sqlStr = sqlStr + " '1'," + VbCrlf
	        sqlStr = sqlStr + " getdate()," + VbCrlf
	        sqlStr = sqlStr + " '[핑거스 아카데미] 고객님 ' +  convert(varchar,convert(int,r.refundrequire))  + '원 환불이 완료되었습니다. 즐거운 하루 되세요.'" + VbCrlf
	        sqlStr = sqlStr + " ,'0','N'"       ''new_info_SMS
	        sqlStr = sqlStr + " from [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a," + VbCrlf
	        sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master m," + VbCrlf
	        sqlStr = sqlStr + "  [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_refund_info r" + VbCrlf
	        sqlStr = sqlStr + " where  a.id in (select T.TEN_CSID from (" & sqlStrTIDX & ") T)" + VbCrlf
	        sqlStr = sqlStr + " and a.id=r.asid" + VbCrlf
	        sqlStr = sqlStr + " and a.orderserial=m.orderserial" + VbCrlf
	        sqlStr = sqlStr + " and a.divcd='A003'" + VbCrlf
	        sqlStr = sqlStr + " and a.deleteyn='N' " + VbCrlf
	        sqlStr = sqlStr + " and a.finishdate>=convert(varchar(10),getdate(),21)" ''2017/08/09 추가 금일 완료된것만 발송.
	        sqlStr = sqlStr + " and r.returnmethod='R007'" + VbCrlf

	    ''	dbget.Execute sqlStr

	    else

		end if

      '' TMPDB 계좌 삭제
      if (sitegubun = "10x10") then
            sqlStr = " Update K " + VbCrlf
            sqlStr = sqlStr + " SET IN_ACCT_NO=''" + VbCrlf
    		sqlStr = sqlStr + " from db_log.dbo.tbl_IBK_ERP_ICHE_DATA K " + VbCrlf
    		sqlStr = sqlStr + " 	Join db_cs.dbo.tbl_new_as_list a " + VbCrlf
    		sqlStr = sqlStr + " 	on K.TEN_CSID=a.id " + VbCrlf
    		sqlStr = sqlStr + " 	Join [db_cs].[dbo].tbl_as_refund_info r " + VbCrlf
    		sqlStr = sqlStr + " 	on a.id=r.asid " + VbCrlf
    		sqlStr = sqlStr + " 	and r.upfiledate='" & upfiledate & "' " + VbCrlf
    		sqlStr = sqlStr + " where K.SITE_NO='2118700620' " + VbCrlf
    		sqlStr = sqlStr + " and K.PROC_YN='Y' " + VbCrlf
    		''sqlStr = sqlStr + " and K.PROC_DATE<>'' " + VbCrlf												'20090616 추가
    		sqlStr = sqlStr + " and K.FL_DATE=Replace(convert(varchar(10),'" & upfiledate & "',21),'-','')  " + VbCrlf
    		sqlStr = sqlStr + " and K.FL_TIME=Replace(Right(convert(varchar(20),'" & upfiledate & "',108),8),':','') " + VbCrlf
    		sqlStr = sqlStr + " and IsNull(K.SITEGUBUN, '10x10') = '10x10' " + VbCrlf

    		dbget.Execute sqlStr
      end if

'     end if
'
'
'
'
'    If (Err.Number = 0) and (ScanErr="") Then
'        dbget.CommitTrans
'        dbACADEMYget.CommitTrans
'    Else
'        dbget.RollBackTrans
'        dbACADEMYget.RollBackTrans
'        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + ")" + Chr(34) + ")</script>"
'        'response.write "<script>history.back()</script>"
'        dbget.close()	:	response.End
'    End If
'    on error Goto 0


    ''''''''''''''sqlStr = " db_cs.dbo.sp_TEN_CS_ASRefundFile_FinishProc"

	'    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	'			,Array("@UPFILEDATE"	, adVarchar	, adParamInput	,   19    , upfiledate) _
	'			,Array("@finishuser"	, adVarchar	, adParamInput	, 	32  , session("ssBctid")) _
	'	)
	'
	'    retParam = fnExecSPOutput(sqlStr, paramInfo)
	'    rowVal = GetValue(retParam,"@RETURN_VALUE")

    if (rowVal=-1) then
        response.write "<script language='javascript'>alert('처리할 내역이 없습니다.');</script>"
        response.write "<script language='javascript'>location.replace('"& referer &"');</script>"
    elseif (rowVal=-2) then
        response.write "<script language='javascript'>alert('처리중 오류가 발생 하였습니다.\n관리자 문의 요망');</script>"
        response.write "<script language='javascript'>location.replace('"& referer &"');</script>"
    else
        response.write "<script language='javascript'>alert('"&rowVal&"건 처리되었습니다.');</script>"
        response.write "<script language='javascript'>location.replace('"& referer &"');</script>"

    end if

    dbget.close()	:	response.End
end if
%>

<script language='javascript'>
alert('<%= rowCount %>개 적용 되었습니다.');
location.replace('<%= referer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
