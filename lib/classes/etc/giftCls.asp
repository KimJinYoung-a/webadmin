<%
'#######################################################
' Description : 기프티콘/기프팅 금액권내역
' History	:  강준구 생성
'              2023.05.23 한용민 수정(엑셀다운로드 재개발. 페이징방식변경해서 전체 다운로드 가능하게 변경함)
'#######################################################

Class cGift_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fgubun
	public fitemid
	public fitemname
	public fsmallimage
	public ftot_sellcash
	public fsellcash
	public fdili_itemcost
	public fsoldout
	public fuseyn
	public FDeliverytype
	public FDefaultfreeBeasongLimit
	public FTenSellcash
	public FItemcost

	public faccountdiv
	public faccountname
	public fuserid
	public fusername
	public ftotalsum
	public fsubtotalprice
	public fcardStatus
	public fregdate

	public fcount
	public ftotal

	public foidx
	public fridx
	public fcardprice

	public fcouponno
	public fcouponidx
	public fcouponvalue
	public fcouponname
	public forderserial

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "업체착불"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "업체"
		Else
			getDeliverytypeName = ""
		End If
	End Function

End Class

Class ClsGift

	public FItemList()
	public FOneItem
	public FIdx
	public FGubun
	public FUseYN
	public FItemID
	public FItemName
	public FTot_Sellcash
	public FSellcash
	public FDiliItemcost
	public FSoldOUT
	public FOrderSerial
	public FUserID
	public FUSerName
	public FReqHP
	public FSDate
	public FEDate
	public fArrList
	public FTotalSum
	public FTCouponNo
	public FNoCouponno

	public FRectdiffPrc
	public FRectdiffCost

	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	public sub FGiftList
		Dim sqlStr, i, vSubQuery

		IF application("Svr_Info") = "Dev" THEN
			vSubQuery = " AND GI.itemid NOT IN(374487,374488,374489,374490,374491) "
		Else
			vSubQuery = " AND GI.itemid NOT IN(588084,588085,588088,588089,588095) "
		End If

		If FGubun <> "" Then
			vSubQuery = vSubQuery & " AND GI.gubun = '" & FGubun & "' "
		End If

		If FItemID <> "" Then
			vSubQuery = vSubQuery & " AND GI.itemid = '" & FItemID & "' "
		End If

		If FItemName <> "" Then
			vSubQuery = vSubQuery & " AND I.itemname like '%" & FItemName & "%' "
		End If

		If FUseYN <> "" Then
			vSubQuery = vSubQuery & " AND GI.useyn = '" & FUseYN & "' "
		End If

		If FSoldOUT = "Y" Then
			vSubQuery = vSubQuery & " AND I.sellyn <> 'Y' "
		ElseIf FSoldOUT = "N" Then
			vSubQuery = vSubQuery & " AND I.sellyn = 'Y' "
		End If

		If FRectdiffPrc <> "" Then
			vSubQuery = vSubQuery & " and GI.sellcash <> I.sellcash "
		End If

		If FRectdiffCost <> "" Then
			vSubQuery = vSubQuery & " and GI.dili_itemcost <> (CASE WHEN I.deliverytype in (2,4,7) THEN 0 "
			vSubQuery = vSubQuery & " WHEN I.deliverytype=1 and I.sellcash >= 30000 then 0  "
			vSubQuery = vSubQuery & " WHEN I.deliverytype=9 and I.sellcash >= c.defaultFreeBeasongLimit then 0  "
			vSubQuery = vSubQuery & " WHEN I.deliverytype=9 THEN c.defaultDeliverPay ELSE " + Cstr(getDefaultBeasongPayByDate(now())) + " END) "
		End If

		sqlStr = "SELECT COUNT(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & _
				 "			FROM [db_order].[dbo].[tbl_mobile_gift_item] AS GI " & _
				 "	INNER JOIN [db_item].[dbo].[tbl_item] AS I ON GI.itemid = I.itemid " & _
				 "	INNER JOIN db_user.dbo.tbl_user_c c on I.makerid = c.userid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " GI.idx, GI.gubun, GI.itemid, GI.tot_sellcash, GI.sellcash, GI.dili_itemcost, GI.useyn, I.smallimage, I.itemname, " & _
				 "				I.limitno, I.limitsold, I.sellyn, I.limityn, I.deliverytype, c.defaultfreeBeasongLimit, I.sellcash as tenSellcash, " & _
				 "				(CASE WHEN I.deliverytype in (2,4,7) THEN 0 " &_
				 "						WHEN I.deliverytype=1 and I.sellcash >= 30000 then 0 " &_
				 "						WHEN I.deliverytype=9 and I.sellcash >= c.defaultFreeBeasongLimit then 0 " &_
				 "						WHEN I.deliverytype=9 THEN c.defaultDeliverPay " &_
				 "						ELSE " + Cstr(getDefaultBeasongPayByDate(now())) + " END) as itemcost " & _
				 "			FROM [db_order].[dbo].[tbl_mobile_gift_item] AS GI " & _
				 "	INNER JOIN [db_item].[dbo].[tbl_item] AS I ON GI.itemid = I.itemid " & _
				 "	INNER JOIN db_user.dbo.tbl_user_c c on I.makerid = c.userid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY GI.idx DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget ,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new cGift_oneitem

					FItemList(i).fidx			= rsget("idx")
					FItemList(i).fgubun			= rsget("gubun")
					FItemList(i).fitemid		= rsget("itemid")
					FItemList(i).fitemname		= rsget("itemname")
					FItemList(i).fsmallimage	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
					FItemList(i).ftot_sellcash	= rsget("tot_sellcash")
					FItemList(i).fsellcash		= rsget("sellcash")
					FItemList(i).fdili_itemcost	= rsget("dili_itemcost")
					FItemList(i).fuseyn			= rsget("useyn")

					IF rsget("limitno")<>"" and rsget("limitsold")<>"" Then
						FItemList(i).fsoldout = (rsget("sellyn")<>"Y") or ((rsget("limityn") = "Y") and (clng(rsget("limitno"))-clng(rsget("limitsold"))<1))
					Else
						FItemList(i).fsoldout = (rsget("sellyn")<>"Y")
					End If
					If (rsget("sellyn") = "S") Then
						FItemList(i).fsoldout = (rsget("sellyn") = "S")
					End IF
					FItemList(i).FDeliverytype	= rsget("deliverytype")
					FItemList(i).FDefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
					FItemList(i).FTenSellcash	= rsget("tenSellcash")
					FItemList(i).FItemcost	= rsget("itemcost")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end sub


	public Function FGiftCont
	Dim strSql
		IF FIdx = "" THEN Exit Function
		strSql = " SELECT GI.idx, GI.gubun, GI.itemid, GI.tot_sellcash, GI.sellcash, GI.dili_itemcost, GI.useyn "&_
				 " 		FROM [db_order].[dbo].[tbl_mobile_gift_item] AS GI "&_
				 " WHERE GI.idx = '" & FIdx & "' "
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FGubun			= rsget("gubun")
			FItemID			= rsget("itemid")
			FTot_Sellcash	= rsget("tot_sellcash")
			FSellcash		= rsget("sellcash")
			FDiliItemcost	= rsget("dili_itemcost")
			FUseYN			= rsget("useyn")

		End IF
		rsget.Close
	End Function


	public Function FGiftStatisticShortList
		Dim sqlStr, i, vSubQuery
		If FGubun <> "" Then
			vSubQuery = " AND O.accountdiv = '" & FGubun & "' "
		Else
			vSubQuery = " AND O.accountdiv IN('550','560') "
		End If

		If FOrderSerial <> "" Then
			vSubQuery = vSubQuery & " AND O.giftOrderSerial = '" & FOrderSerial & "' "
		End If

		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND R.userid = '" & FUserID & "' "
		End If

		If FUSerName <> "" Then
			vSubQuery = vSubQuery & " AND U.username = '" & FUSerName & "' "
		End If

		If FReqHP <> "" Then
			vSubQuery = vSubQuery & " AND U.usercell = '" & FReqHP & "' "
		End If

		If FSDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate > '" & FSDate & "' "
			vTotalSum = "o"
		End If

		If FEDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate < '" & DateAdd("d", 1, FEDate) & "' "
			vTotalSum = "o"
		End If


		sqlStr = "SELECT O.totalsum, count(O.idx) AS cnt, sum(O.subtotalprice) AS total " & _
				 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O " & _
				 "	INNER JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R ON O.giftOrderSerial = R.giftOrderSerial " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U ON R.userid = U.userid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	GROUP BY O.totalsum " & _
				 "	ORDER BY O.totalsum ASC "
'response.write sqlStr
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget.RecordCount

		If  not rsget.EOF  then
			FGiftStatisticShortList = rsget.getRows()
		Else
			ftotalcount = 0
		End If
		rsget.Close
	end Function

	' 밑에 함수를 수정할경우 반드시 FGiftStatisticList_notpaging 도 동일하게 수정해 주세요.
	public sub FGiftStatisticList
		Dim sqlStr, i, vSubQuery, vTotalSum
		vTotalSum = "x"

		If FGubun <> "" Then
			vSubQuery = " AND O.accountdiv = '" & FGubun & "' "
		Else
			vSubQuery = " AND O.accountdiv IN('550','560') "
		End If

		If FOrderSerial <> "" Then
			vSubQuery = vSubQuery & " AND O.giftOrderSerial = '" & FOrderSerial & "' "
		End If

		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND R.userid = '" & FUserID & "' "
		End If

		If FUSerName <> "" Then
			vSubQuery = vSubQuery & " AND U.username = '" & FUSerName & "' "
		End If

		If FReqHP <> "" Then
			vSubQuery = vSubQuery & " AND U.usercell = '" & FReqHP & "' "
		End If

		If FSDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate > '" & FSDate & "' "
			vTotalSum = "o"
		End If

		If FEDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate < '" & DateAdd("d", 1, FEDate) & "' "
			vTotalSum = "o"
		End If

		sqlStr = "SELECT COUNT(O.giftOrderSerial) " & _
				 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O with (nolock)" & _
				 "	INNER JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R with (nolock) ON O.giftOrderSerial = R.giftOrderSerial " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U with (nolock) ON R.userid = U.userid " & _
				 "	LEFT JOIN [db_order].[dbo].[tbl_mobile_gift] AS G with (nolock) ON O.masterCardCode = G.masterCardCode " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		ftotalcount = rsget(0)
		rsget.Close

		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & _ 
				 " IsNull(G.couponno,'') AS couponno, O.accountdiv, R.userid, isNull(U.username,'탈퇴or휴면회원') AS username, O.totalsum" &_
				 " , O.subtotalprice, R.cardStatus, O.regdate " & _
				 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O with (nolock)" & _
				 "	INNER JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R with (nolock) ON O.giftOrderSerial = R.giftOrderSerial " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U with (nolock) ON R.userid = U.userid " & _
				 "	LEFT JOIN [db_order].[dbo].[tbl_mobile_gift] AS G with (nolock) ON O.masterCardCode = G.masterCardCode " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY O.giftOrderSerial DESC "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		rsget.PageSize= FPageSize
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new cGift_oneitem

					FItemList(i).fcouponno		= rsget("couponno")
					FItemList(i).faccountdiv	= rsget("accountdiv")
					If rsget("accountdiv") = "550" Then
						FItemList(i).faccountname = "기프팅"
					ElseIf rsget("accountdiv") = "560" Then
						FItemList(i).faccountname = "기프티콘"
					End IF
					FItemList(i).fuserid		= rsget("userid")
					FItemList(i).fusername		= rsget("username")
					FItemList(i).ftotalsum		= rsget("totalsum")
					FItemList(i).fsubtotalprice	= rsget("subtotalprice")
					FItemList(i).fcardStatus	= rsget("cardStatus")
					FItemList(i).fregdate		= rsget("regdate")

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close

		If vTotalSum = "o" Then
			sqlStr = "SELECT isNull(SUM(subtotalprice),0) " & _
					 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O with (nolock)" & _
					 "	INNER JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R with (nolock) ON O.giftOrderSerial = R.giftOrderSerial " & _
					 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U with (nolock) ON R.userid = U.userid " & _
					 "	WHERE 1=1 " & _
					 "	" & vSubQuery & " "
			'response.write sqlStr & "<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalSum = rsget(0)
			rsget.Close
		End IF
	end sub

	' 밑에 함수를 수정할경우 반드시 FGiftStatisticList 의 검색 조건도 동일하게 수정해 주세요.
	public sub FGiftStatisticList_notpaging
		Dim sqlStr, i, vSubQuery, vTotalSum
		vTotalSum = "x"

		If FGubun <> "" Then
			vSubQuery = " AND O.accountdiv = '" & FGubun & "' "
		Else
			vSubQuery = " AND O.accountdiv IN('550','560') "
		End If

		If FOrderSerial <> "" Then
			vSubQuery = vSubQuery & " AND O.giftOrderSerial = '" & FOrderSerial & "' "
		End If

		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND R.userid = '" & FUserID & "' "
		End If

		If FUSerName <> "" Then
			vSubQuery = vSubQuery & " AND U.username = '" & FUSerName & "' "
		End If

		If FReqHP <> "" Then
			vSubQuery = vSubQuery & " AND U.usercell = '" & FReqHP & "' "
		End If

		If FSDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate > '" & FSDate & "' "
			vTotalSum = "o"
		End If

		If FEDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate < '" & DateAdd("d", 1, FEDate) & "' "
			vTotalSum = "o"
		End If

		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & _ 
				 " IsNull(G.couponno,'') AS couponno, O.accountdiv, R.userid, isNull(U.username,'탈퇴or휴면회원') AS username, O.totalsum" &_
				 " , O.subtotalprice, R.cardStatus, O.regdate " & _
				 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O with (nolock)" & _
				 "	INNER JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R with (nolock) ON O.giftOrderSerial = R.giftOrderSerial " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U with (nolock) ON R.userid = U.userid " & _
				 "	LEFT JOIN [db_order].[dbo].[tbl_mobile_gift] AS G with (nolock) ON O.masterCardCode = G.masterCardCode " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY O.giftOrderSerial DESC "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
    	dbget.CommandTimeout = 60*5   ' 5분
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		rsget.PageSize= FPageSize
		If not rsget.EOF  then
			fArrList = rsget.getrows()
		End If
		rsget.Close
	end sub

	public sub FCouponStatisticList
		Dim sqlStr, i, vSubQuery, vTotalSum
		vTotalSum = "x"

		If FGubun <> "" Then
			vSubQuery = " AND G.gubun = '" & FGubun & "' "
		End If

		If FTCouponNo <> "" Then
			vSubQuery = " AND G.couponno = '" & FTCouponNo & "' "
		End If

		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND G.userid = '" & FUserID & "' "
		End If

		If FUSerName <> "" Then
			vSubQuery = vSubQuery & " AND U.username = '" & FUSerName & "' "
		End If

		If FReqHP <> "" Then
			vSubQuery = vSubQuery & " AND U.usercell = '" & FReqHP & "' "
		End If

		If FNoCouponno <> "o" Then
			vSubQuery = vSubQuery & " AND G.couponidx <> 0 "
		End If

		If FSDate <> "" Then
			vSubQuery = vSubQuery & " AND G.regdate > '" & FSDate & "' "
			vTotalSum = "o"
		End If

		If FEDate <> "" Then
			vSubQuery = vSubQuery & " AND G.regdate < '" & DateAdd("d", 1, FEDate) & "' "
			vTotalSum = "o"
		End If

		sqlStr = "SELECT COUNT(G.idx) " & _
				 "			FROM [db_order].[dbo].[tbl_mobile_gift] AS G " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U ON G.userid = U.userid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close

		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " IsNull(G.couponno,'') AS couponno, G.couponidx, G.gubun, G.userid, isNull(U.username,'탈퇴or휴면회원') AS username, C.couponvalue, M.couponname, " & _
				 "				G.regdate, G.orderserial " & _
				 "			FROM [db_order].[dbo].[tbl_mobile_gift] AS G " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U ON G.userid = U.userid " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_coupon] AS C ON G.userid = C.userid AND G.couponidx = C.idx " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_user_coupon_master] AS M ON C.masteridx = M.idx " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY G.idx DESC "
'response.write sqlStr
		rsget.Open sqlStr, dbget ,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		rsget.PageSize= FPageSize
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new cGift_oneitem

					FItemList(i).fcouponno		= rsget("couponno")
					If rsget("gubun") = "giftting" Then
						FItemList(i).fgubun = "기프팅"
					ElseIf rsget("gubun") = "gifticon" Then
						FItemList(i).fgubun = "기프티콘"
					End IF
					FItemList(i).fuserid		= rsget("userid")
					FItemList(i).fusername		= rsget("username")
					FItemList(i).fcouponidx		= rsget("couponidx")
					FItemList(i).fcouponvalue	= rsget("couponvalue")
					FItemList(i).fcouponname	= rsget("couponname")
					FItemList(i).fregdate		= rsget("regdate")
					FItemList(i).forderserial	= rsget("orderserial")

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close

		If vTotalSum = "o" Then
			sqlStr = "SELECT isNull(SUM(subtotalprice),0) " & _
					 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O " & _
					 "	INNER JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R ON O.giftOrderSerial = R.giftOrderSerial " & _
					 "	LEFT JOIN [db_user].[dbo].[tbl_user_n] AS U ON R.userid = U.userid " & _
					 "	WHERE 1=1 " & _
					 "	" & vSubQuery & " "
			'rsget.Open sqlStr, dbget ,1
			'FTotalSum = rsget(0)
			'rsget.Close
		End IF
	end sub


	public Function FGiftStatisticNew
		Dim sqlStr, i, vSubQuery
		If FGubun = "10x10" Then
			vSubQuery = " AND O.accountdiv Not IN('550','560') "
		ElseIf FGubun = "550" Then
			vSubQuery = " AND O.accountdiv IN('" & FGubun & "') "
		ElseIf FGubun = "560" Then
			vSubQuery = " AND O.accountdiv IN('" & FGubun & "') "
		End If

		If FSDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate > '" & FSDate & "' "
		End If

		If FEDate <> "" Then
			vSubQuery = vSubQuery & " AND O.regdate < '" & DateAdd("d", 1, FEDate) & "' "
		End If


		sqlStr = "SELECT O.totalsum, count(O.idx), sum(isNull(O.subtotalprice,0)), count(R.idx), sum(isNull(R.cardPrice,0)) " & _
				 "			FROM [db_order].[dbo].[tbl_giftcard_order] AS O " & _
				 "	LEFT JOIN [db_user].[dbo].[tbl_giftcard_regList] AS R ON O.giftOrderSerial = R.giftOrderSerial " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	GROUP BY O.totalsum " & _
				 "	ORDER BY O.totalsum ASC "
'response.write sqlStr
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget.RecordCount
		FResultCount = ftotalcount

		redim preserve FItemList(FResultCount)

		If  not rsget.EOF  then
			Do Until rsget.Eof
				set FItemList(i) = new cGift_oneitem

					FItemList(i).ftotalsum		= rsget(0)
					FItemList(i).foidx			= rsget(1)
					FItemList(i).fsubtotalprice	= rsget(2)
					FItemList(i).fridx			= rsget(3)
					FItemList(i).fcardprice		= rsget(4)

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Function



	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class

function GetCardName(totalsum)
	select case totalsum
		case "5000"
			: GetCardName = "5천원권"
		case "10000"
			: GetCardName = "1만원권"
		case "20000"
			: GetCardName = "2만원권"
		case "30000"
			: GetCardName = "3만원권"
		case "50000"
			: GetCardName = "5만원권"
		case "80000"
			: GetCardName = "8만원권"
		case "100000"
			: GetCardName = "10만원권"
		case "150000"
			: GetCardName = "15만원권"
		case "200000"
			: GetCardName = "20만원권"
		case "300000"
			: GetCardName = "30만원권"
		case else
			: GetCardName = ""
	end select
end function

function GetCardStatusName(cardStatus)
	if IsNULL(cardStatus) then cardStatus = "0"
	cardStatus = Trim(cardStatus)

	select case cardStatus
		case "1"
			: GetCardStatusName = "등록완료"
		case "3"
			: GetCardStatusName = "등록취소"
		case "5"
			: GetCardStatusName = "카드만료"
		case "0"
			: GetCardStatusName = "등록이전"
		case else
			: GetCardStatusName = ""
	end select
end function

function GetCardStatusColor(cardStatus)
	if cardStatus="0" then
		GetCardStatusColor="#FF0000"
	elseif cardStatus="1" then
		GetCardStatusColor="#44BBBB"
	elseif cardStatus="2" then
		GetCardStatusColor="#000000"
	elseif cardStatus="3" then
		GetCardStatusColor="#000000"
	elseif cardStatus="4" then
		GetCardStatusColor="#0000FF"
	elseif cardStatus="5" then
		GetCardStatusColor="#CC9933"
	elseif cardStatus="6" then
		GetCardStatusColor="#FF00FF"
	elseif cardStatus="7" then
		GetCardStatusColor="#EE2222"
	elseif cardStatus="8" then
		GetCardStatusColor="#EE2222"
	elseif cardStatus="9" then
		GetCardStatusColor="#FF0000"
	end if
end function

%>