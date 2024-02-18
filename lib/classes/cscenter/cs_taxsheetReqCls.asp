<%

class CTaxRequestItem
	public Fidx
	public Forderserial
	public Fuserid
	public Fgoodname
	public Fgroupid

	public FbusiIdx
	public FbusiNO
	public FtaxIdx
	public FuseYN
	public Fregdate

	public Fbuyname
	public FOrdItemCNT
	public FCHuLItemCNT
	public FchulgoPriceSum
	public Fcompany_name
	public Fvatinclude

	public FlastChulgoDate

	public function GetGoodNameStr()
		if (Abs(FCHuLItemCNT) = 1) or (Abs(FCHuLItemCNT) = 0) then
			GetGoodNameStr = Fgoodname
		elseif (FCHuLItemCNT > 0) then
			GetGoodNameStr = Fgoodname & " 외 " & (FCHuLItemCNT - 1) & " 개"
		else
			GetGoodNameStr = Fgoodname & " 외 " & (FCHuLItemCNT + 1) & " 개"
		end if
	end function

	public function GetChulgoState()
		if (FOrdItemCNT = 0) then
			GetChulgoState = "취소"
		elseif (FOrdItemCNT = FCHuLItemCNT) then
			GetChulgoState = "출고완료"
		elseif (Abs(FOrdItemCNT) > Abs(FCHuLItemCNT)) then
			if (FCHuLItemCNT = 0) then
				GetChulgoState = "출고이전"
			else
				GetChulgoState = "일부출고"
			end if
		elseif (Abs(FOrdItemCNT) < Abs(FCHuLItemCNT)) then
			GetChulgoState = "ERROR"
		end if
	end function

	public function GetChulgoStateColor()
		if (FOrdItemCNT = 0) then
			GetChulgoStateColor = "red"
		elseif (FOrdItemCNT = FCHuLItemCNT) then
			GetChulgoStateColor = "blue"
		elseif (Abs(FOrdItemCNT) > Abs(FCHuLItemCNT)) then
			GetChulgoStateColor = "black"
		elseif (Abs(FOrdItemCNT) < Abs(FCHuLItemCNT)) then
			GetChulgoStateColor = "red"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CTaxRequest
	public FTaxList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectIDX
	public FRectOrderserial
	public FRectUseYN
	public FRectFinishYN
	public FRectExcIssueFinish

	public Sub GetTaxRequestList()
		dim strSql, addSql, i

		'검색 추가 쿼리
		if FRectOrderserial<>"" then
			addSql = addSql & " and t.orderserial='" & FRectOrderserial & "' "
		end if

		if FRectUseYN<>"" then
			addSql = addSql & " and t.useYN='" & FRectUseYN & "' "
		end if

		if FRectFinishYN<>"" then
			addSql = addSql & " and t.finishYN='" & FRectFinishYN & "' "
		end if

		'// ====================================================================
		'@ 총데이터수

		strSql = " Select count(t.idx) as cnt "
		strSql = strSql + " from "
		strSql = strSql + " 	db_log.dbo.tbl_tax_issue_request as t "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + addSql

		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		'// ====================================================================
		strSql = " select  top " + CStr(CStr(FPageSize*FCurrPage)) + " "

		strSql = strSql + " 	t.idx, t.orderserial, t.userid, t.busiIdx, t.useYN, t.regdate "

		''strSql = strSql + " 	t.idx, t.orderserial, t.userid, t.goodname, t.vatYN, t.req_itemno, t.chulgo_itemno, t.req_price, t.chulgo_price, t.req_tax "
		''strSql = strSql + " 	, t.chulgo_tax, t.lastChulgoDate, t.groupid, t.busiIdx, t.taxIdx, t.useYN, t.regdate "
		strSql = strSql + " from "
		strSql = strSql + " 	db_log.dbo.tbl_tax_issue_request as t "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "

		strSql = strSql + addSql

		strSql = strSql + " order by "
		strSql = strSql + " 	t.idx desc "


		''response.write strSql
		rsget.pagesize = FPageSize
		rsget.Open strSql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.AbsolutePage = FCurrPage
			do until rsget.eof
				set FTaxList(i) = new CTaxRequestItem

				FTaxList(i).Fidx				= rsget("idx")
				FTaxList(i).Forderserial		= rsget("orderserial")
				FTaxList(i).Fuserid				= rsget("userid")
				'FTaxList(i).Fgoodname			= rsget("goodname")
				'FTaxList(i).FvatYN				= rsget("vatYN")
				'FTaxList(i).Freq_itemno			= rsget("req_itemno")
				'FTaxList(i).Fchulgo_itemno		= rsget("chulgo_itemno")
				'FTaxList(i).Freq_price			= rsget("req_price")
				'FTaxList(i).Fchulgo_price		= rsget("chulgo_price")
				'FTaxList(i).Freq_tax			= rsget("req_tax")
				'FTaxList(i).Fchulgo_tax			= rsget("chulgo_tax")
				'FTaxList(i).FlastChulgoDate		= rsget("lastChulgoDate")
				'FTaxList(i).Fgroupid			= rsget("groupid")
				FTaxList(i).FbusiIdx			= rsget("busiIdx")
				'FTaxList(i).FtaxIdx				= rsget("taxIdx")
				FTaxList(i).FuseYN				= rsget("useYN")
				FTaxList(i).Fregdate			= rsget("regdate")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	public Sub GetTaxRequestOneOrder()
		dim strSql, addSql, i

		'// ====================================================================
		strSql = " select R.*, IsNull(T.taxIdx, -1) as taxIdx "
		strSql = strSql + " from "
		strSql = strSql + " 	( "
		strSql = strSql + " 		select "
		strSql = strSql + " 			r.idx, r.busiIdx, r.useYN, r.regdate, m.orderserial "
		strSql = strSql + " 			, IsNull(m.userid, '') as userid "
		strSql = strSql + " 			, m.buyname "
		strSql = strSql + " 			, sum(CASE WHEN d.cancelyn<>'Y' and d.itemid<>0 then d.itemno ELSE 0 END) as OrdItemCNT "
		strSql = strSql + " 			, sum(CASE WHEN d.cancelyn<>'Y' and d.itemid<>0 and d.beasongdate is Not NULL then d.itemno ELSE 0 END) as CHuLItemCNT "
		strSql = strSql + " 			, sum(case when d.cancelyn<>'Y' and (d.itemid=0 or d.beasongdate is Not NULL) then reducedprice*d.itemno else 0 end) as chulgoPriceSum "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN 'G00456' ELSE p.groupid END) as groupid "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN '(주)텐바이텐' ELSE g.company_name END) as company_name "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN '211-87-00620' ELSE g.company_no END) as busiNO "
		strSql = strSql + " 			, IsNull(max(CASE WHEN d.cancelyn<>'Y' and d.itemid<>0 then d.itemname ELSE '' END), '') as goodname "
		strSql = strSql + " 			, d.vatinclude "
		strSql = strSql + " 			, max(case when d.cancelyn<>'Y' and d.itemid<>0 and d.beasongdate is Not NULL then convert(varchar(10), d.beasongdate, 121) else '2000-01-01' end) as lastChulgoDate "
		strSql = strSql + " 		from "
		strSql = strSql + " 			db_log.[dbo].[tbl_tax_issue_request] r "
		strSql = strSql + " 			join db_order.dbo.tbl_order_master m "
		strSql = strSql + " 			on "
		strSql = strSql + " 				r.orderserial = m.orderserial "
		strSql = strSql + " 			join db_order.dbo.tbl_order_detail d "
		strSql = strSql + " 			on "
		strSql = strSql + " 				m.orderserial = d.orderserial "
		strSql = strSql + " 			Join db_partner.dbo.tbl_partner p "
		strSql = strSql + " 			on "
		strSql = strSql + " 				d.makerid=p.id "
		strSql = strSql + " 			join db_partner.dbo.tbl_partner_group g "
		strSql = strSql + " 			on "
		strSql = strSql + " 				p.groupid = g.groupid "
		strSql = strSql + " 		where "
		strSql = strSql + " 			1 = 1 "
		strSql = strSql + " 			and r.orderserial = '" + CStr(FRectOrderserial) + "' "
		strSql = strSql + " 			and r.useYN = 'Y' "
		strSql = strSql + " 		group by "
		strSql = strSql + " 			r.idx, r.busiIdx, r.useYN, r.regdate, m.orderserial "
		strSql = strSql + " 			, IsNull(m.userid, '') "
		strSql = strSql + " 			, m.buyname "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN 'G00456' ELSE p.groupid END) "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN 'T' ELSE 'WU' END) "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN '(주)텐바이텐' ELSE g.company_name END) "
		strSql = strSql + " 			, (CASE WHEN d.omwdiv='M' or d.itemid=0 and isNULL(d.makerid,'')='' THEN '211-87-00620' ELSE g.company_no END) "
		strSql = strSql + " 			, d.vatinclude "
		strSql = strSql + " 	) R "
		strSql = strSql + " 	left join ( "
		strSql = strSql + " 		select t.orderserial, s.busiNO, t.taxIdx, t.taxtype "
		strSql = strSql + " 		from "
		strSql = strSql + " 			db_log.[dbo].[tbl_tax_issue_request] r "
		strSql = strSql + " 			join db_order.dbo.tbl_taxSheet t "
		strSql = strSql + " 			on "
		strSql = strSql + " 				r.orderserial = t.orderserial "
		strSql = strSql + " 			join db_order.dbo.tbl_busiInfo s "
		strSql = strSql + " 			on "
		strSql = strSql + " 				1 = 1 "
		strSql = strSql + " 				and r.orderserial = '" + CStr(FRectOrderserial) + "' "
		strSql = strSql + " 				and r.useYN = 'Y' "
		strSql = strSql + " 				and t.supplyBusiIdx = s.busiidx "
		strSql = strSql + " 				and t.delYN <> 'Y' "
		strSql = strSql + " 	) T "
		strSql = strSql + " 	on "
		strSql = strSql + " 		1 = 1 "
		strSql = strSql + " 		and T.orderserial = R.orderserial "
		strSql = strSql + " 		and T.busiNO = R.busiNO "
		strSql = strSql + " 		and R.vatinclude = T.taxtype "
		strSql = strSql + " order by "
		strSql = strSql + " 	R.groupid "

		''response.write strSql
		rsget.pagesize = FPageSize
		rsget.Open strSql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.AbsolutePage = FCurrPage
			do until rsget.eof
				set FTaxList(i) = new CTaxRequestItem

				''FTaxList(i).Fidx				= rsget("idx")
				FTaxList(i).Forderserial		= rsget("orderserial")
				FTaxList(i).Fuserid				= rsget("userid")
				FTaxList(i).FbusiIdx			= rsget("busiIdx")
				FTaxList(i).FuseYN				= rsget("useYN")
				FTaxList(i).Fregdate			= rsget("regdate")

				FTaxList(i).Fbuyname			= rsget("buyname")
				FTaxList(i).FOrdItemCNT			= rsget("OrdItemCNT")
				FTaxList(i).FCHuLItemCNT		= rsget("CHuLItemCNT")
				FTaxList(i).FchulgoPriceSum		= rsget("chulgoPriceSum")
				FTaxList(i).Fgroupid			= rsget("groupid")
				FTaxList(i).Fcompany_name		= rsget("company_name")
				FTaxList(i).FbusiNO				= rsget("busiNO")
				FTaxList(i).Fgoodname			= rsget("goodname")
				FTaxList(i).Fvatinclude			= rsget("vatinclude")
				FTaxList(i).FtaxIdx				= rsget("taxIdx")

				FTaxList(i).FlastChulgoDate		= rsget("lastChulgoDate")



				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub


	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FTaxList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

%>
