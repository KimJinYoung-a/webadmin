<%
Class COutmallItem
	Public FIdx
	Public FMallid
	Public FApiAction
	Public FItemid
	Public FItemoption
	Public FPriority
	Public FRegdate
	Public FReaddate
	Public FFindate
	Public FResultCode
	Public FLastErrMsg
	Public FOptioncnt
	Public FAccFailCnt
	Public FEtcMallSellyn
	Public FQuantity
	Public FRegedOptCnt
	Public FGSShopSellyn
	Public FLastUserid
	Public FMidx
	Public FOutmallGoodno
	Public FTitle
	Public FEtc1
	Public FStartdate
	Public FEnddate
End Class

Class COutmallItemSummaryItem
	public Fyyyymmdd
	public Fsellsite
	public FTTL
	public FmallActive
	public FmallWait
	public FmallAVailSellY
	public FregWiat
	public FregFail
	public FmallInActive
End Class

Class COutmallQueSummaryItem
	public Fyyyymmdd
	public Fsellsite
	public FapiAction
	public FTTL
	public FS_OK
	public FS_ERR
	public FS_DUPP
	public FS_BLANK
	public FS_NOREAD
	public FS_NULL

	public FTTL_H
	public FS_OK_H
	public FS_ERR_H
	public FS_DUPP_H
	public FS_BLANK_H
	public FS_NOREAD_H
	public FS_NULL_H
End Class


Class COutmallSummary
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectSellsite
	Public FRectSnapDate
	Public FRectIsSysUser
	public FRectApiAction

	public Sub getOutSailyQueSummary()
		Dim sqlStr
		if (FRectIsSysUser="") then FRectIsSysUser="NULL"

		sqlStr = "exec [db_etcmall].[dbo].[sp_Ten_OutMall_API_Que_SummaryList] '"&FRectSnapDate&"','"&FRectSellsite&"',"&FRectIsSysUser&",'"&FRectApiAction&"'"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		i=0
		Redim preserve FItemList(FResultCount)
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new COutmallQueSummaryItem
				if (FRectSellsite<>"") then
					FItemList(i).Fyyyymmdd			= rsget("yyyymmdd")
				else
					FItemList(i).Fsellsite			= rsget("mallid")
				end if

				FItemList(i).FapiAction		= rsget("apiAction")
				FItemList(i).FTTL			= rsget("TTL")
				FItemList(i).FS_OK			= rsget("S_OK")
				FItemList(i).FS_ERR			= rsget("S_ERR")
				FItemList(i).FS_DUPP		= rsget("S_DUPP")
				FItemList(i).FS_BLANK		= rsget("S_BLANK")
				FItemList(i).FS_NOREAD		= rsget("S_NOREAD")
				FItemList(i).FS_NULL		= rsget("S_NULL")

				FItemList(i).FTTL_H			= rsget("TTL_H")
				FItemList(i).FS_OK_H		= rsget("S_OK_H")
				FItemList(i).FS_ERR_H		= rsget("S_ERR_H")
				FItemList(i).FS_DUPP_H		= rsget("S_DUPP_H")
				FItemList(i).FS_BLANK_H		= rsget("S_BLANK_H")
				FItemList(i).FS_NOREAD_H	= rsget("S_NOREAD_H")
				FItemList(i).FS_NULL_H		= rsget("S_NULL_H")


				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	public Sub getOutItemSummaryList()
		Dim sqlStr
		sqlStr = "exec [db_etcmall].[dbo].[sp_Ten_OutMall_daily_SummaryList] '"&FRectSnapDate&"','"&FRectSellsite&"'"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		i=0
		Redim preserve FItemList(FResultCount)
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new COutmallItemSummaryItem
				if (FRectSellsite<>"") then
					FItemList(i).Fyyyymmdd			= rsget("yyyymmdd")
				else
					FItemList(i).Fsellsite			= rsget("mallid")
				end if
				FItemList(i).FTTL				= rsget("TTL")
				FItemList(i).FmallActive		= rsget("mallActive")
				FItemList(i).FmallWait			= rsget("mallWait")
				FItemList(i).FmallAVailSellY	= rsget("mallAVailSellY")
				FItemList(i).FregWiat			= rsget("regWiat")
				FItemList(i).FregFail			= rsget("regFail")
				FItemList(i).FmallInActive		= rsget("mallInActive")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
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
End Class

Class COutmall
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectMallid
	Public FRectItemid
	Public FRectApiAction
	Public FRectResultCode
	Public FRectLastUserid
	Public FRectGSShopSellyn
	Public FRectErrMsg

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Sub getQueLogList
		Dim sqlStr, addSql, i

		'텐바이텐 상품번호 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and Q.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and Q.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectMallid <> "" Then
			addSql = addSql & " and Q.mallid = '"&FRectMallid&"' "
		End If

		If FRectApiAction <> "" Then
			If FRectApiAction = "SOLDOUT" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EditSellYn') "
			ElseIf FRectApiAction = "EDITPOLICY" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EDITBATCH') "
			ElseIf FRectApiAction = "EditInfo" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EDITBATCH') "
			Else
				addSql = addSql & " and Q.apiAction = '"&FRectApiAction&"' "
			End If
		End If

		If FRectResultCode <> "" Then
			If FRectResultCode = "QNull" Then
				addSql = addSql & " and isnull(Q.readdate, '') = '' "
			ElseIf FRectResultCode = "ERR" Then
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
				addSql = addSql & " and G.accFailcnt > 0 "
			Else
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
			End If
		End If

		If FRectLastUserid <> "" Then
			If FRectLastUserid <> "system" Then
				addSql = addSql & " and Q.lastUserid <> 'system' "
			Else
				addSql = addSql & " and Q.lastUserid = 'system' "
			End If
		End If

		if FRectGSShopSellyn <> "" Then
			Select Case FRectMallid
				Case "lotteimall"
					addSql = addSql & " and G.ltimallSellyn = '" & FRectGSShopSellyn & "' "
				Case "gsshop"
					addSql = addSql & " and G.GSShopSellyn = '" & FRectGSShopSellyn & "'"
				Case "lotteCom"
					addSql = addSql & " and G.lotteSellyn  = '" & FRectGSShopSellyn & "' "
				Case "ezwel"
					addSql = addSql & " and G.ezwelSellyn  = '" & FRectGSShopSellyn & "' "
				Case "cjmall"
					addSql = addSql & " and G.cjmallSellyn  = '" & FRectGSShopSellyn & "' "
				Case "auction1010"
					addSql = addSql & " and G.AuctionSellyn  = '" & FRectGSShopSellyn & "' "
				Case "gmarket1010"
					addSql = addSql & " and G.GmarketSellyn  = '" & FRectGSShopSellyn & "' "
				Case "homeplus"
					addSql = addSql & " and G.HomeplusSellyn  = '" & FRectGSShopSellyn & "' "
				Case "interpark"
					addSql = addSql & " and G.mayiParkSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstorefarm"
					addSql = addSql & " and G.nvstorefarmSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstoremoonbangu"
					addSql = addSql & " and G.nvstoremoonbanguSellyn  = '" & FRectGSShopSellyn & "' "
				Case "Mylittlewhoopee"
					addSql = addSql & " and G.MylittlewhoopeeSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstorefarmclass"
					addSql = addSql & " and G.nvClassSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstoregift"
					addSql = addSql & " and G.nvstoregiftSellyn  = '" & FRectGSShopSellyn & "' "
				Case "11stmy"
					addSql = addSql & " and G.my11stSellyn  = '" & FRectGSShopSellyn & "' "
				Case "11st1010"
					addSql = addSql & " and G.st11Sellyn  = '" & FRectGSShopSellyn & "' "
				Case "sabangnet"
					addSql = addSql & " and G.sabangnetSellyn  = '" & FRectGSShopSellyn & "' "
				Case "ssg"
					addSql = addSql & " and G.ssgSellyn  = '" & FRectGSShopSellyn & "' "
				Case "halfclub"
					addSql = addSql & " and G.HalfClubSellyn  = '" & FRectGSShopSellyn & "' "
				Case "lfmall"
					addSql = addSql & " and G.lfmallSellyn  = '" & FRectGSShopSellyn & "' "
				Case "coupang"
					addSql = addSql & " and G.coupangSellyn  = '" & FRectGSShopSellyn & "' "
				Case "hmall1010"
					addSql = addSql & " and G.hmallSellyn  = '" & FRectGSShopSellyn & "' "
				Case "WMP"
					addSql = addSql & " and G.wemakeSellyn  = '" & FRectGSShopSellyn & "' "
				Case "wmpfashion"
					addSql = addSql & " and G.wfwemakeSellyn  = '" & FRectGSShopSellyn & "' "
				Case "kakaogift"
					addSql = addSql & " and G.kakaoGiftSellYn  = '" & FRectGSShopSellyn & "' "
				Case "lotteon"
					addSql = addSql & " and G.lotteonSellYn  = '" & FRectGSShopSellyn & "' "
				Case "shintvshopping"
					addSql = addSql & " and G.shintvshoppingSellYn  = '" & FRectGSShopSellyn & "' "
				Case "wetoo1300k"
					addSql = addSql & " and G.wetoo1300kSellYn  = '" & FRectGSShopSellyn & "' "
				Case "skstoa"
					addSql = addSql & " and G.skstoaSellYn  = '" & FRectGSShopSellyn & "' "
				Case "shopify"
					addSql = addSql & " and G.shopifySellYn  = '" & FRectGSShopSellyn & "' "
				Case "kakaostore"
					addSql = addSql & " and G.kakaostoreSellYn  = '" & FRectGSShopSellyn & "' "
				Case "boribori1010"
					addSql = addSql & " and G.boriboriSellYn  = '" & FRectGSShopSellyn & "' "
				Case "bindmall1010"
					addSql = addSql & " and G.bindmallSellYn  = '" & FRectGSShopSellyn & "' "
				Case "wconcept1010"
					addSql = addSql & " and G.wconceptSellYn  = '" & FRectGSShopSellyn & "' "
				Case "benepia1010"
					addSql = addSql & " and G.benepiaSellYn  = '" & FRectGSShopSellyn & "' "
			End Select
		end if

		If FRectErrMsg <> "" Then
			addSql = addSql & " and Q.lastErrMsg like '%"&FRectErrMsg&"%' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(Q.idx) as cnt, CEILING(CAST(Count(Q.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_API_Que] as Q with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i with (nolock) on Q.itemid = i.itemid "

		Select Case FRectMallid
			Case "lotteimall"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "gsshop"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_gsshop_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "lotteCom"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "ezwel"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "cjmall"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_cjmall_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "auction1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_auction_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "gmarket1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_gmarket_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "homeplus"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "interpark"		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_interpark_reg_Item as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarm"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarm_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoremoonbangu"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "Mylittlewhoopee"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_Mylittlewhoopee_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarmclass"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoregift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoregift_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "11stmy"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_my11st_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "sabangnet"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "ssg"				sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ssg_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "halfclub"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_halfclub_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "lfmall"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lfmall_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "hmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "WMP"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wmpfashion"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "coupang"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaogift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.[tbl_kakaoGift_regItem] as G with (nolock) on Q.itemid = G.itemid "
			Case "11st1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "lotteon"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shintvshopping"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "skstoa"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shopify"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaostore"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "boribori1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "bindmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wconcept1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "benepia1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wetoo1300k"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
		End Select
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		''rw sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " Q.idx, Q.mallid, Q.apiAction, Q.itemid, Q.priority, Q.regdate, Q.readdate, Q.findate, Q.resultCode, Q.lastErrMsg "
		sqlStr = sqlStr & " ,i.optioncnt,G.accFailCnt, G.regedOptCnt, Q.lastUserid  "

		Select Case FRectMallid
			Case "lotteimall"		sqlStr = sqlStr & ", G.ltimallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "gsshop"			sqlStr = sqlStr & ", G.GSShopSellyn, '' as outmallGoodno "
			Case "lotteCom"			sqlStr = sqlStr & ", G.lotteSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "ezwel"			sqlStr = sqlStr & ", G.ezwelSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "cjmall"			sqlStr = sqlStr & ", G.cjmallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "auction1010"		sqlStr = sqlStr & ", G.AuctionSellyn as GSShopSellyn, isNull(G.auctionGoodno, '') as outmallGoodno  "
			Case "gmarket1010"		sqlStr = sqlStr & ", G.GmarketSellyn as GSShopSellyn, isNull(G.gmarketGoodno, '') as outmallGoodno "
			Case "homeplus"			sqlStr = sqlStr & ", G.HomeplusSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "interpark"		sqlStr = sqlStr & ", G.mayiParkSellyn as GSShopSellyn, isNull(G.interparkprdno, '') as outmallGoodno "
			Case "nvstorefarm"		sqlStr = sqlStr & ", G.nvstorefarmSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "nvstorefarmclass"	sqlStr = sqlStr & ", G.nvclassSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "nvstoremoonbangu"	sqlStr = sqlStr & ", G.nvstoremoonbanguSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "Mylittlewhoopee"	sqlStr = sqlStr & ", G.MylittlewhoopeeSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "nvstoregift"		sqlStr = sqlStr & ", G.nvstoregiftSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "11stmy"			sqlStr = sqlStr & ", G.my11stSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "11st1010"			sqlStr = sqlStr & ", G.st11Sellyn as GSShopSellyn, isNull(G.st11GoodNo, '') as outmallGoodno "
			Case "sabangnet"		sqlStr = sqlStr & ", G.sabangnetSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "ssg"				sqlStr = sqlStr & ", G.ssgSellyn as GSShopSellyn, isNull(G.ssgGoodno, '') as outmallGoodno  "
			Case "halfclub"			sqlStr = sqlStr & ", G.HalfClubSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "lfmall"			sqlStr = sqlStr & ", G.lfmallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "coupang"			sqlStr = sqlStr & ", G.coupangSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "hmall1010"		sqlStr = sqlStr & ", G.hmallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "WMP"				sqlStr = sqlStr & ", G.wemakeSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "wmpfashion"		sqlStr = sqlStr & ", G.wfwemakeSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "kakaogift"		sqlStr = sqlStr & ", G.kakaoGiftSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "lotteon"			sqlStr = sqlStr & ", G.lotteonSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "shintvshopping"	sqlStr = sqlStr & ", G.shintvshoppingSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "wetoo1300k"		sqlStr = sqlStr & ", G.wetoo1300kSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "skstoa"			sqlStr = sqlStr & ", G.skstoaSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "shopify"			sqlStr = sqlStr & ", G.shopifySellYn as GSShopSellyn, '' as outmallGoodno "
			Case "kakaostore"		sqlStr = sqlStr & ", G.kakaostoreSellYn as GSShopSellyn, isNull(G.kakaostoreGoodno, '') as outmallGoodno "
			Case "boribori1010"		sqlStr = sqlStr & ", G.boriboriSellYn as GSShopSellyn, isNull(G.boriboriGoodno, '') as outmallGoodno "
			Case "bindmall1010"		sqlStr = sqlStr & ", G.bindmallSellYn as GSShopSellyn, isNull(G.bindmallGoodno, '') as outmallGoodno "
			Case "wconcept1010"		sqlStr = sqlStr & ", G.wconceptSellYn as GSShopSellyn, isNull(G.wconceptGoodno, '') as outmallGoodno "
			Case "benepia1010"		sqlStr = sqlStr & ", G.benepiaSellYn as GSShopSellyn, isNull(G.benepiaGoodno, '') as outmallGoodno "
		End Select

		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_API_Que] as Q with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i with (nolock) on Q.itemid = i.itemid "

		Select Case FRectMallid
			Case "lotteimall"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "gsshop"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_gsshop_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "lotteCom"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "ezwel"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "cjmall"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_cjmall_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "auction1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_auction_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "gmarket1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_gmarket_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "homeplus"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "interpark"		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_interpark_reg_Item as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarm"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarm_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarmclass"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoremoonbangu"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "Mylittlewhoopee"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_Mylittlewhoopee_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoregift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoregift_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "11stmy"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_my11st_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "sabangnet"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "ssg"				sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ssg_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "halfclub"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_halfclub_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "lfmall"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lfmall_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "hmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "WMP"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wmpfashion"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "coupang"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaogift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.[tbl_kakaoGift_regItem] as G with (nolock) on Q.itemid = G.itemid "
			Case "11st1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "lotteon"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shintvshopping"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wetoo1300k"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "skstoa"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shopify"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaostore"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "boribori1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "bindmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wconcept1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "benepia1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If

			Case "benepia1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
		End Select

		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY Q.findate DESC, Q.itemid DESC "
		'rw sqlStr

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COutmallItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FMallid		= rsget("mallid")
					FItemList(i).FApiAction		= rsget("apiAction")
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FPriority		= rsget("priority")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FReaddate		= rsget("readdate")
					FItemList(i).FFindate		= rsget("findate")
					FItemList(i).FResultCode	= rsget("resultCode")
					FItemList(i).FLastErrMsg	= db2html(rsget("lastErrMsg"))
					FItemList(i).FOptioncnt		= rsget("optioncnt")
					FItemList(i).FAccFailCnt	= rsget("accFailCnt")
					FItemList(i).FRegedOptCnt	= rsget("regedOptCnt")
					FItemList(i).FGSShopSellyn	= rsget("GSShopSellyn")
					FItemList(i).FLastUserid	= rsget("lastUserid")
					FItemList(i).FOutmallGoodno	= rsget("outmallGoodno")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getQueLogbackupList
		Dim sqlStr, addSql, i

		'텐바이텐 상품번호 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and Q.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and Q.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectMallid <> "" Then
			addSql = addSql & " and Q.mallid = '"&FRectMallid&"' "
		End If

		If FRectApiAction <> "" Then
			If FRectApiAction = "SOLDOUT" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EditSellYn') "
			ElseIf FRectApiAction = "EDITPOLICY" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EDITBATCH') "
			ElseIf FRectApiAction = "EditInfo" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EDITBATCH') "
			Else
				addSql = addSql & " and Q.apiAction = '"&FRectApiAction&"' "
			End If
		End If

		If FRectResultCode <> "" Then
			If FRectResultCode = "QNull" Then
				addSql = addSql & " and isnull(Q.readdate, '') = '' "
			ElseIf FRectResultCode = "ERR" Then
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
				addSql = addSql & " and G.accFailcnt > 0 "
			Else
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
			End If
		End If

		If FRectLastUserid <> "" Then
			If FRectLastUserid <> "system" Then
				addSql = addSql & " and Q.lastUserid <> 'system' "
			Else
				addSql = addSql & " and Q.lastUserid = 'system' "
			End If
		End If

		if FRectGSShopSellyn <> "" Then
			Select Case FRectMallid
				Case "lotteimall"
					addSql = addSql & " and G.ltimallSellyn = '" & FRectGSShopSellyn & "' "
				Case "gsshop"
					addSql = addSql & " and G.GSShopSellyn = '" & FRectGSShopSellyn & "'"
				Case "lotteCom"
					addSql = addSql & " and G.lotteSellyn  = '" & FRectGSShopSellyn & "' "
				Case "ezwel"
					addSql = addSql & " and G.ezwelSellyn  = '" & FRectGSShopSellyn & "' "
				Case "cjmall"
					addSql = addSql & " and G.cjmallSellyn  = '" & FRectGSShopSellyn & "' "
				Case "auction1010"
					addSql = addSql & " and G.AuctionSellyn  = '" & FRectGSShopSellyn & "' "
				Case "gmarket1010"
					addSql = addSql & " and G.GmarketSellyn  = '" & FRectGSShopSellyn & "' "
				Case "homeplus"
					addSql = addSql & " and G.HomeplusSellyn  = '" & FRectGSShopSellyn & "' "
				Case "interpark"
					addSql = addSql & " and G.mayiParkSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstorefarm"
					addSql = addSql & " and G.nvstorefarmSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstoremoonbangu"
					addSql = addSql & " and G.nvstoremoonbanguSellyn  = '" & FRectGSShopSellyn & "' "
				Case "Mylittlewhoopee"
					addSql = addSql & " and G.MylittlewhoopeeSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstorefarmclass"
					addSql = addSql & " and G.nvClassSellyn  = '" & FRectGSShopSellyn & "' "
				Case "nvstoregift"
					addSql = addSql & " and G.nvstoregiftSellyn  = '" & FRectGSShopSellyn & "' "
				Case "11stmy"
					addSql = addSql & " and G.my11stSellyn  = '" & FRectGSShopSellyn & "' "
				Case "11st1010"
					addSql = addSql & " and G.st11Sellyn  = '" & FRectGSShopSellyn & "' "
				Case "sabangnet"
					addSql = addSql & " and G.sabangnetSellyn  = '" & FRectGSShopSellyn & "' "
				Case "ssg"
					addSql = addSql & " and G.ssgSellyn  = '" & FRectGSShopSellyn & "' "
				Case "halfclub"
					addSql = addSql & " and G.HalfClubSellyn  = '" & FRectGSShopSellyn & "' "
				Case "lfmall"
					addSql = addSql & " and G.lfmallSellyn  = '" & FRectGSShopSellyn & "' "
				Case "coupang"
					addSql = addSql & " and G.coupangSellyn  = '" & FRectGSShopSellyn & "' "
				Case "hmall1010"
					addSql = addSql & " and G.hmallSellyn  = '" & FRectGSShopSellyn & "' "
				Case "WMP"
					addSql = addSql & " and G.wemakeSellyn  = '" & FRectGSShopSellyn & "' "
				Case "wmpfashion"
					addSql = addSql & " and G.wfwemakeSellyn  = '" & FRectGSShopSellyn & "' "
				Case "kakaogift"
					addSql = addSql & " and G.kakaoGiftSellYn  = '" & FRectGSShopSellyn & "' "
				Case "lotteon"
					addSql = addSql & " and G.lotteonSellYn  = '" & FRectGSShopSellyn & "' "
				Case "shintvshopping"
					addSql = addSql & " and G.shintvshoppingSellYn  = '" & FRectGSShopSellyn & "' "
				Case "wetoo1300k"
					addSql = addSql & " and G.wetoo1300kSellYn  = '" & FRectGSShopSellyn & "' "
				Case "skstoa"
					addSql = addSql & " and G.skstoaSellYn  = '" & FRectGSShopSellyn & "' "
				Case "shopify"
					addSql = addSql & " and G.shopifySellYn  = '" & FRectGSShopSellyn & "' "
				Case "kakaostore"
					addSql = addSql & " and G.kakaostoreSellYn  = '" & FRectGSShopSellyn & "' "
				Case "boribori1010"
					addSql = addSql & " and G.boriboriSellYn  = '" & FRectGSShopSellyn & "' "
			End Select
		end if

		If FRectErrMsg <> "" Then
			addSql = addSql & " and Q.lastErrMsg like '%"&FRectErrMsg&"%' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(Q.idx) as cnt, CEILING(CAST(Count(Q.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_API_Que_LOG] as Q with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i with (nolock) on Q.itemid = i.itemid "

		Select Case FRectMallid
			Case "lotteimall"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "gsshop"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_gsshop_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "lotteCom"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "ezwel"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "cjmall"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_cjmall_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "auction1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_auction_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "gmarket1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_gmarket_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "homeplus"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "interpark"		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_interpark_reg_Item as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarm"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarm_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoremoonbangu"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "Mylittlewhoopee"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_Mylittlewhoopee_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarmclass"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoregift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoregift_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "11stmy"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_my11st_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "sabangnet"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "ssg"				sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ssg_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "halfclub"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_halfclub_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "lfmall"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lfmall_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "hmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "WMP"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wmpfashion"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "coupang"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaogift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.[tbl_kakaoGift_regItem] as G with (nolock) on Q.itemid = G.itemid "
			Case "11st1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "lotteon"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shintvshopping"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "skstoa"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shopify"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaostore"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "boribori1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "bindmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wconcept1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "benepia1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wetoo1300k"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
		End Select
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		''rw sqlStr
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " Q.idx, Q.mallid, Q.apiAction, Q.itemid, Q.priority, Q.regdate, Q.readdate, Q.findate, Q.resultCode, Q.lastErrMsg "
		sqlStr = sqlStr & " ,i.optioncnt,G.accFailCnt, G.regedOptCnt, Q.lastUserid  "

		Select Case FRectMallid
			Case "lotteimall"		sqlStr = sqlStr & ", G.ltimallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "gsshop"			sqlStr = sqlStr & ", G.GSShopSellyn, '' as outmallGoodno "
			Case "lotteCom"			sqlStr = sqlStr & ", G.lotteSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "ezwel"			sqlStr = sqlStr & ", G.ezwelSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "cjmall"			sqlStr = sqlStr & ", G.cjmallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "auction1010"		sqlStr = sqlStr & ", G.AuctionSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "gmarket1010"		sqlStr = sqlStr & ", G.GmarketSellyn as GSShopSellyn, isNull(G.gmarketGoodno, '') as outmallGoodno "
			Case "homeplus"			sqlStr = sqlStr & ", G.HomeplusSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "interpark"		sqlStr = sqlStr & ", G.mayiParkSellyn as GSShopSellyn, isNull(G.interparkprdno, '') as outmallGoodno "
			Case "nvstorefarm"		sqlStr = sqlStr & ", G.nvstorefarmSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "nvstorefarmclass"	sqlStr = sqlStr & ", G.nvclassSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "nvstoremoonbangu"	sqlStr = sqlStr & ", G.nvstoremoonbanguSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "Mylittlewhoopee"	sqlStr = sqlStr & ", G.MylittlewhoopeeSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "nvstoregift"		sqlStr = sqlStr & ", G.nvstoregiftSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "11stmy"			sqlStr = sqlStr & ", G.my11stSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "11st1010"			sqlStr = sqlStr & ", G.st11Sellyn as GSShopSellyn, isNull(G.st11GoodNo, '') as outmallGoodno "
			Case "sabangnet"		sqlStr = sqlStr & ", G.sabangnetSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "ssg"				sqlStr = sqlStr & ", G.ssgSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "halfclub"			sqlStr = sqlStr & ", G.HalfClubSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "lfmall"			sqlStr = sqlStr & ", G.lfmallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "coupang"			sqlStr = sqlStr & ", G.coupangSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "hmall1010"		sqlStr = sqlStr & ", G.hmallSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "WMP"				sqlStr = sqlStr & ", G.wemakeSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "wmpfashion"		sqlStr = sqlStr & ", G.wfwemakeSellyn as GSShopSellyn, '' as outmallGoodno "
			Case "kakaogift"		sqlStr = sqlStr & ", G.kakaoGiftSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "lotteon"			sqlStr = sqlStr & ", G.lotteonSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "shintvshopping"	sqlStr = sqlStr & ", G.shintvshoppingSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "wetoo1300k"		sqlStr = sqlStr & ", G.wetoo1300kSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "skstoa"			sqlStr = sqlStr & ", G.skstoaSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "shopify"			sqlStr = sqlStr & ", G.shopifySellYn as GSShopSellyn, '' as outmallGoodno "
			Case "kakaostore"		sqlStr = sqlStr & ", G.kakaostoreSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "boribori1010"		sqlStr = sqlStr & ", G.boriboriSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "bindmall1010"		sqlStr = sqlStr & ", G.bindmallSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "wconcept1010"		sqlStr = sqlStr & ", G.wconceptSellYn as GSShopSellyn, '' as outmallGoodno "
			Case "benepia1010"		sqlStr = sqlStr & ", G.benepiaSellYn as GSShopSellyn, '' as outmallGoodno "
		End Select

		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_API_Que_LOG] as Q with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i with (nolock) on Q.itemid = i.itemid "

		Select Case FRectMallid
			Case "lotteimall"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_ltimall_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "gsshop"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_gsshop_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "lotteCom"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_lotte_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "ezwel"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "cjmall"			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_cjmall_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "auction1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_auction_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "gmarket1010"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_gmarket_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "homeplus"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as G with (nolock) on Q.itemid = G.itemid "
			Case "interpark"		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_interpark_reg_Item as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarm"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarm_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstorefarmclass"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoremoonbangu"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "Mylittlewhoopee"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_Mylittlewhoopee_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "nvstoregift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoregift_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "11stmy"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_my11st_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "sabangnet"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "ssg"				sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ssg_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "halfclub"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_halfclub_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "lfmall"			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lfmall_regItem as G with (nolock) on Q.itemid = G.itemid "
			Case "hmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_hmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "WMP"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wmpfashion"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wfwemake_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "coupang"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_coupang_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaogift"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.[tbl_kakaoGift_regItem] as G with (nolock) on Q.itemid = G.itemid "
			Case "11st1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_11st_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "lotteon"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteon_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shintvshopping"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wetoo1300k"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "skstoa"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "shopify"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_shopify_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "kakaostore"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_kakaostore_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "boribori1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_boribori_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "bindmall1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_bindmall_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "wconcept1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wconcept_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
			Case "benepia1010"
				If FRectApiAction = "REG" Then
					sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				Else
					sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_benepia_regItem as G with (nolock) on Q.itemid = G.itemid "
				End If
		End Select

		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY Q.findate DESC, Q.itemid DESC "
		'rw sqlStr

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COutmallItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FMallid		= rsget("mallid")
					FItemList(i).FApiAction		= rsget("apiAction")
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FPriority		= rsget("priority")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FReaddate		= rsget("readdate")
					FItemList(i).FFindate		= rsget("findate")
					FItemList(i).FResultCode	= rsget("resultCode")
					FItemList(i).FLastErrMsg	= db2html(rsget("lastErrMsg"))
					FItemList(i).FOptioncnt		= rsget("optioncnt")
					FItemList(i).FAccFailCnt	= rsget("accFailCnt")
					FItemList(i).FRegedOptCnt	= rsget("regedOptCnt")
					FItemList(i).FGSShopSellyn	= rsget("GSShopSellyn")
					FItemList(i).FLastUserid	= rsget("lastUserid")
					FItemList(i).FOutmallGoodno	= rsget("outmallGoodno")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub






	Public Sub getQueOptionLogList
		Dim sqlStr, addSql, i

		'텐바이텐 상품번호 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and M.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and M.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectMallid <> "" Then
			addSql = addSql & " and Q.mallid = '"&FRectMallid&"' "
			addSql = addSql & " and M.mallid = '"&FRectMallid&"' "
		End If

		If FRectApiAction <> "" Then
			If FRectApiAction = "SOLDOUT" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EditSellYn') "
			Else
				addSql = addSql & " and Q.apiAction = '"&FRectApiAction&"' "
			End If
		End If

		If FRectResultCode <> "" Then
			If FRectResultCode = "QNull" Then
				addSql = addSql & " and isnull(Q.readdate, '') = '' "
			ElseIf FRectResultCode = "ERR" Then
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
				addSql = addSql & " and R.accFailcnt > 0 "
			Else
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
			End If
		End If

		If FRectLastUserid <> "" Then
			If FRectLastUserid <> "system" Then
				addSql = addSql & " and Q.lastUserid <> 'system' "
			Else
				addSql = addSql & " and Q.lastUserid = 'system' "
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(Q.idx) as cnt, CEILING(CAST(Count(Q.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_Option_API_Que] as Q "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_Outmall_option_Manager as M on Q.midx = M.idx "
		Select Case FRectMallid
			Case "gsshop"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as R on M.idx = R.midx "
			Case "lotteCom"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteAddoption_regitem as R on M.idx = R.midx "
			Case "lotteimall"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as R on M.idx = R.midx "
		End Select
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on M.itemid = i.itemid  "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option as o on M.itemid = o.itemid and M.itemoption = o.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " Q.idx, Q.mallid, Q.apiAction, M.itemid, M.itemoption, Q.priority, Q.regdate, Q.readdate, Q.findate, Q.resultCode, Q.lastErrMsg "
		sqlStr = sqlStr & " ,R.accFailCnt, Q.lastUserid, Q.midx "
		Select Case FRectMallid
			Case "gsshop"			sqlStr = sqlStr & ", R.GSShopSellyn, R.GSShopGoodno as outmallGoodno "
			Case "lotteCom"			sqlStr = sqlStr & ", R.lotteSellyn as GSShopSellyn, R.lottegoodno as outmallGoodno "
			Case "lotteimall"		sqlStr = sqlStr & ", R.ltimallSellyn as GSShopSellyn, R.ltimallgoodno as outmallGoodno "
		End Select
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_Option_API_Que] as Q "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_Outmall_option_Manager as M on Q.midx = M.idx "
		Select Case FRectMallid
			Case "gsshop"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as R on M.idx = R.midx "
			Case "lotteCom"		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteAddoption_regitem as R on M.idx = R.midx "
			Case "lotteimall"	sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as R on M.idx = R.midx "
		End Select
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on M.itemid = i.itemid  "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option as o on M.itemid = o.itemid and M.itemoption = o.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY Q.findate DESC, M.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COutmallItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FMallid		= rsget("mallid")
					FItemList(i).FApiAction		= rsget("apiAction")
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FItemoption	= rsget("itemoption")
					FItemList(i).FPriority		= rsget("priority")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FReaddate		= rsget("readdate")
					FItemList(i).FFindate		= rsget("findate")
					FItemList(i).FResultCode	= rsget("resultCode")
					FItemList(i).FLastErrMsg	= db2html(rsget("lastErrMsg"))
					FItemList(i).FAccFailCnt	= rsget("accFailCnt")
					FItemList(i).FGSShopSellyn	= rsget("GSShopSellyn")
					FItemList(i).FLastUserid	= rsget("lastUserid")
					FItemList(i).FMidx			= rsget("midx")
					FItemList(i).FOutmallGoodno	= rsget("outmallGoodno")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function fnQueGroupCntList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mallid, COUNT(*) as cnt "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_API_Que "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and readdate is null "
		sqlStr = sqlStr & " and mallid <> 'appDTL' "
		sqlStr = sqlStr & " and LEFT(apiaction, 3) <> 'REG' "
		sqlStr = sqlStr & " GROUP BY mallid "
		sqlStr = sqlStr & " ORDER by mallid ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			fnQueGroupCntList = rsget.getRows()
		End If
		rsget.close
	End Function

	Public Function getMallActionList
		Dim sqlStr, addSql, i

		If FRectMallid <> "" Then
			addSql = addSql & " and mallid = '"&FRectMallid&"' "
		End If

		If FRectApiAction <> "" Then
			If FRectApiAction = "SOLDOUT" Then
				addSql = addSql & " and apiAction in ('"&FRectApiAction&"', 'EditSellYn') "
			ElseIf FRectApiAction = "CHKSTAT" Then
				addSql = addSql & " and apiAction in ('"&FRectApiAction&"', 'CONFIRM') "
			Else
				addSql = addSql & " and apiAction = '"&FRectApiAction&"' "
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 6000 itemid, apiaction, count(*) as cnt "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_API_Que q"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and readdate is null "
		sqlStr = sqlStr & " and LEFT(apiaction, 3) <> 'REG' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " GROUP BY itemid, apiaction "
		If FRectMallid = "11st1010" Then
			sqlStr = sqlStr & " ORDER BY apiaction DESC, (SELECT accfailcnt FROM db_etcmall.dbo.tbl_11st_regitem r WHERE r.itemid = q.itemid) ASC "
		Else
			sqlStr = sqlStr & " ORDER BY apiaction DESC, count(*) DESC "
		End If
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			getMallActionList = rsget.getRows()
		End If
		rsget.close
	End Function

	Public Sub getQueLogNewitemList
		Dim sqlStr, addSql, i

		'텐바이텐 상품번호 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and Q.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and Q.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectMallid <> "" Then
			addSql = addSql & " and Q.mallid = '"&FRectMallid&"' "
		End If

		If FRectApiAction <> "" Then
			If FRectApiAction = "SOLDOUT" Then
				addSql = addSql & " and Q.apiAction in ('"&FRectApiAction&"', 'EditSellYn') "
			Else
				addSql = addSql & " and Q.apiAction = '"&FRectApiAction&"' "
			End If
		End If

		If FRectResultCode <> "" Then
			If FRectResultCode = "QNull" Then
				addSql = addSql & " and isnull(Q.readdate, '') = '' "
			ElseIf FRectResultCode = "ERR" Then
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
				addSql = addSql & " and G.accFailcnt > 0 "
			Else
				addSql = addSql & " and Q.resultCode = '"&FRectResultCode&"' "
			End If
		End If

		If FRectLastUserid <> "" Then
			If FRectLastUserid <> "system" Then
				addSql = addSql & " and Q.lastUserid <> 'system' "
			Else
				addSql = addSql & " and Q.lastUserid = 'system' "
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(Q.idx) as cnt, CEILING(CAST(Count(Q.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i  "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		Select Case FRectMallid
			Case "zilingo"		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_zilingo_regItem] as G on G.itemid = i.itemid and G.itemoption = IsNULL(o.itemoption,'0000') "
		End Select
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_outmall_newItem_API_Que] as Q on Q.itemid = i.itemid and G.itemoption = IsNULL(Q.itemoption,'0000') "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " Q.idx, Q.mallid, Q.apiAction, Q.itemid, Q.itemoption, Q.priority, Q.regdate, Q.readdate, Q.findate, Q.resultCode, Q.lastErrMsg "
		sqlStr = sqlStr & " ,i.optioncnt,G.accFailCnt, Q.lastUserid  "
		Select Case FRectMallid
			Case "zilingo"		sqlStr = sqlStr & ", G.zilingoSellyn as etcMallSellyn, G.quantity "
		End Select
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i  "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		Select Case FRectMallid
			Case "zilingo"		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_zilingo_regItem] as G on G.itemid = i.itemid and G.itemoption = IsNULL(o.itemoption,'0000') "
		End Select
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_outmall_newItem_API_Que] as Q on Q.itemid = i.itemid and G.itemoption = IsNULL(Q.itemoption,'0000') "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY Q.findate DESC, Q.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COutmallItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FMallid		= rsget("mallid")
					FItemList(i).FApiAction		= rsget("apiAction")
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FItemoption	= rsget("itemoption")
					FItemList(i).FPriority		= rsget("priority")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FReaddate		= rsget("readdate")
					FItemList(i).FFindate		= rsget("findate")
					FItemList(i).FResultCode	= rsget("resultCode")
					FItemList(i).FLastErrMsg	= db2html(rsget("lastErrMsg"))
					FItemList(i).FOptioncnt		= rsget("optioncnt")
					FItemList(i).FAccFailCnt	= rsget("accFailCnt")
					FItemList(i).FEtcMallSellyn	= rsget("etcMallSellyn")
					FItemList(i).FQuantity		= rsget("quantity")
					FItemList(i).FLastUserid	= rsget("lastUserid")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getInboundNotScheduleitemList
		Dim sqlStr, addSql, i

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_tmp_ScheduleSplit "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, title, itemid, etc1, startdate, enddate, regdate "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_tmp_ScheduleSplit "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COutmallItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FTitle			= rsget("title")
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FEtc1			= rsget("etc1")
					FItemList(i).FStartdate		= rsget("startdate")
					FItemList(i).FEnddate		= rsget("enddate")
					FItemList(i).FRegdate		= rsget("regdate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
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
End Class
%>
