<%

Class CCSRefundCheckItem
	Public Fasid
	Public Fdivcd
	Public FOrderserial

	Public Fdivcdname
	Public Fgubun01name
	Public Fgubun02name
	Public Ftitle

	Public Freturnmethod
	Public FreturnmethodName
	Public Frefundresult
	Public FOrgRefundRequire

	Public Fregdate
	Public Ffinishdate

	Public Fadd_upchejungsandeliverypay
	public Fadd_upchejungsancause
	public Freturndeliverpay

	public FappPrice

	Private Sub Class_Initialize()

	end Sub

	Private Sub Class_Terminate()

	End Sub

End Class

class CCSRefundCheck
	Public FItemList()
	Public FOneItem

	Public FPageSize
	Public FTotalPage
    Public FPageCount
	Public FTotalCount
	Public FResultCount
    Public FScrollCount
	Public FCurrPage

	Public FRectDivCD
	Public FRectReturnMethod
	Public FRectStartDate
	Public FRectEndDate
	Public FRectOrderSerial

	Public FRectChkGubun

	public FRectRefundMin
	public FRectRefundMax

	public FRectExCheckFinish
	public FRectReturnMethodIN

	public FrefundSUM
	public FaddjungSUM
	public FRectDategbn

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage 		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetRefundCheckList
	    dim sqlStr, addSql, i

		addSql = " where "
		addSql = addSql + " 	1 = 1 "
		addSql = addSql + " 	and a.currstate = 'B007' "
		addSql = addSql + " 	and a.deleteyn = 'N' "

		if FRectDategbn="regdate" then
			If (FRectStartDate <> "") Then
				addSql = addSql + " 	and a.regdate >= '" & FRectStartDate & "' "
			End If
			If (FRectEndDate <> "") Then
				addSql = addSql + " 	and a.regdate < '" & FRectEndDate & "' "
			End If
		else
			If (FRectStartDate <> "") Then
				addSql = addSql + " 	and a.finishdate >= '" & FRectStartDate & "' "
			End If
			If (FRectEndDate <> "") Then
				addSql = addSql + " 	and a.finishdate < '" & FRectEndDate & "' "
			End If
		end if

		If (FRectDivCD <> "") Then
			addSql = addSql + " 	and a.divcd = '" & FRectDivCD & "' "
		End If

		If (FRectReturnMethod <> "") Then
			if (FRectReturnMethod = "REXC") then
				addSql = addSql + " 	and r.returnmethod not in ('R007', 'R910', 'R900') "
			else
				addSql = addSql + " 	and r.returnmethod = '" & FRectReturnMethod & "' "
			end if
		End If

		If (FRectReturnMethodIN <> "") Then
			addSql = addSql + " 	and r.returnmethod in (" & FRectReturnMethodIN & ") "
		end if

		If (FRectOrderSerial <> "") Then
			addSql = addSql + " 	and a.orderserial = '" & FRectOrderSerial & "' "
		End If

		If (FRectChkGubun <> "") Then
			Select Case FRectChkGubun
				Case "err"
					addSql = addSql + "		and a.divcd = 'A003' "
					addSql = addSql + " 	and (case when a.divcd in ('A003', 'A007') then IsNULL(r.refundresult,0) else IsNULL(r.refundrequire,0) end) <> (case when a.divcd in ('A004', 'A010') then IsNULL(r.refundrequire,0) else IsNULL(r2.refundrequire,0) end) "
				Case "ret"
					'// 반품
					addSql = addSql + "		and a.divcd in ('A004', 'A010') "
					addSql = addSql + "		and (IsNull(u.add_upchejungsandeliverypay, 0) <> 0) "
				Case "etc"
					'// 업체기타정산
					addSql = addSql + "		and a.divcd = 'A700' "
					addSql = addSql + "		and (IsNull(u.add_upchejungsandeliverypay, 0) <> 0) "
					''A700
				Case "addjung"
					'// 업체추가정산(반품 등)
					''addSql = addSql + "		and a.divcd = 'A700' "
					addSql = addSql + "		and (IsNull(u.add_upchejungsandeliverypay, 0) <> 0) "
					''A700
				Case "retbea"
					'// 반품 + 불량제외 + 반품배송비차감없음
					''C004	CD01	변심
					''C005,C006,C007	상품관련, 물류관련, 택배사관련
					addSql = addSql + "		and a.divcd = 'A004' "
					addSql = addSql + " 	and (a.gubun01 = 'C004' and a.gubun02 in ('CD01','CD06')) "
					''addSql = addSql + " 	and a.gubun01 not in ('C005', 'C006', 'C007') "
					''addSql = addSql + " 	and (case when a.divcd <> 'A003' then (IsNull(r.refundbeasongpay,0) + IsNull(r.refunddeliverypay,0)) else 0 end) >= 0 "
				Case "retbeaTen"
					'// 반품 + 불량제외 + 반품배송비차감없음
					''C004	CD01	변심
					''C005,C006,C007	상품관련, 물류관련, 택배사관련
					addSql = addSql + "		and a.divcd = 'A010' "
					addSql = addSql + " 	and (a.gubun01 = 'C004' and a.gubun02 in ('CD01','CD06')) "
					''addSql = addSql + " 	and a.gubun01 not in ('C005', 'C006', 'C007') "
					''addSql = addSql + " 	and (case when a.divcd <> 'A003' then (IsNull(r.refundbeasongpay,0) + IsNull(r.refunddeliverypay,0)) else 0 end) >= 0 "
				Case ""
					''
				Case ""
					''
				Case Else
					''
			End Select
		End If

		If (FRectRefundMin <> "") Then
			addSql = addSql + " 	and (case when a.divcd in ('A003', 'A007') then IsNULL(r.refundresult,0) else IsNULL(r.refundrequire,0) end) >= '" & FRectRefundMin & "' "
		End If

		If (FRectRefundMax <> "") Then
			addSql = addSql + " 	and (case when a.divcd in ('A003', 'A007') then IsNULL(r.refundresult,0) else IsNULL(r.refundrequire,0) end) <= '" & FRectRefundMax & "' "
		End If

		if (FRectExCheckFinish = "Y") then
			addSql = addSql + " 	and a.title <> '예치금을 무통장으로 환불' "
			addSql = addSql + " 	and a.title <> '제휴몰 구매확정 후 환불' "
			addSql = addSql + " 	and a.title <> '고객입금 차액환불' "
			addSql = addSql + " 	and a.title <> '업체정산 및 고객환불' "
			addSql = addSql + " 	and a.title <> 'CS서비스 - 무통장 환불(배송비)' "
		end if

		sqlStr = "select IsNull(sum(case when a.divcd in ('A003', 'A007') then IsNULL(r.refundresult,0) else IsNULL(r.refundrequire,0) end),0) as totSUM "
		sqlStr = sqlStr + " , IsNull(sum(IsNull(u.add_upchejungsandeliverypay, 0)), 0) as totAddJung "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_as_refund_info] r "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = r.asid "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_as_upcheAddjungsan] u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = u.asid "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].[tbl_as_refund_info] r2 on a.refasid = r2.asid "
		sqlStr = sqlStr + addSql

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FrefundSUM = rsget("totSUM")
			FaddjungSUM = rsget("totAddJung")
		rsget.Close

		sqlStr = "select count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_as_refund_info] r "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = r.asid "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_as_upcheAddjungsan] u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = u.asid "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].[tbl_as_refund_info] r2 on a.refasid = r2.asid "
		sqlStr = sqlStr + addSql

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & (FPageSize * FCurrPage) & " "
		sqlStr = sqlStr + " a.id as asid, a.* "
		sqlStr = sqlStr + " , r.returnmethod, C4.comm_name as returnmethodName, (case when a.divcd in ('A003', 'A007') then IsNULL(r.refundresult,0) else IsNULL(r.refundrequire,0) end) as refundresult "
		sqlStr = sqlStr + " , C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name "
		sqlStr = sqlStr + " , (case when a.divcd in ('A004', 'A010') then IsNULL(r.refundrequire,0) else IsNULL(r2.refundrequire,0) end) as orgRefundRequire "
		sqlStr = sqlStr + " , IsNull(u.add_upchejungsandeliverypay, 0) as add_upchejungsandeliverypay "
		sqlStr = sqlStr + " , IsNull(u.add_upchejungsancause, '') as add_upchejungsancause "
		sqlStr = sqlStr + " , (case when a.divcd <> 'A003' then (IsNull(r.refundbeasongpay,0) + IsNull(r.refunddeliverypay,0)) else 0 end) as returndeliverpay "
		''sqlStr = sqlStr + " , IsNull((select top 1 l.appPrice from [db_order].[dbo].[tbl_onlineApp_log] l where l.pggubun = 'bankipkum' and l.orderserial = a.orderserial),0) as appPrice "
		sqlStr = sqlStr + " , IsNull((select sum(l.appPrice) as appPrice from [db_order].[dbo].[tbl_onlineApp_log] l "
		sqlStr = sqlStr + " where l.pggubun = 'bankipkum' and l.orderserial = a.orderserial "
		sqlStr = sqlStr + " and l.appDate >= '" & FRectStartDate & "' "
		sqlStr = sqlStr + " and l.appDate < '" & FRectEndDate & "'),0) as appPrice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_as_refund_info] r "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = r.asid "
		sqlStr = sqlStr + " 	left join [db_cs].[dbo].[tbl_as_upcheAddjungsan] u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = u.asid "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C1 on a.divcd=C1.comm_cd "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C2 on a.gubun01=C2.comm_cd "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C3 on a.gubun02=C3.comm_cd "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C4 on r.returnmethod=C4.comm_cd and C4.comm_group='Z090' "
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].[tbl_as_refund_info] r2 on a.refasid = r2.asid "
		sqlStr = sqlStr + addSql
		sqlStr = sqlStr + " order by a.finishdate desc, a.id desc "

		'response.write sqlStr & "<Br>"
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			i = 0

			do until (i >= FResultCount)

				set FItemList(i) = new CCSRefundCheckItem

				FItemList(i).Fasid 			= rsget("asid")
				FItemList(i).Fdivcd 		= rsget("divcd")
				FItemList(i).FOrderserial 	= rsget("orderserial")

				FItemList(i).Fdivcdname 	= rsget("divcdname")
				FItemList(i).Fgubun01name 	= rsget("gubun01name")
				FItemList(i).Fgubun02name 	= rsget("gubun02name")
				FItemList(i).Ftitle 		= db2html(rsget("title"))

				FItemList(i).Freturnmethod 		= rsget("returnmethod")
				FItemList(i).FreturnmethodName 	= rsget("returnmethodName")
				FItemList(i).Frefundresult 		= rsget("refundresult")
				FItemList(i).FOrgRefundRequire 	= rsget("orgRefundRequire")

				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Ffinishdate 	= rsget("finishdate")

				FItemList(i).Fadd_upchejungsandeliverypay 	= rsget("add_upchejungsandeliverypay")
				FItemList(i).Fadd_upchejungsancause 		= rsget("add_upchejungsancause")
				FItemList(i).Freturndeliverpay 				= rsget("returndeliverpay")

				FItemList(i).FappPrice			= rsget("appPrice")

				rsget.movenext
				i=i+1

			loop
		end if
		rsget.Close

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


%>
