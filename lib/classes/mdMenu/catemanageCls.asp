<%
'###########################################################
' History : 2014.08.11 김진영 생성
'###########################################################

Class CMDCategory_item
	Public FUserid
	Public FUsername
	Public FCatename
	Public FTargetMoney
	Public FProfitMoney
	Public FYyyy
	Public FMm

	Public FCatecode
	Public FDepth
	Public FTarget1
	Public FTarget2
	Public FTarget3
	Public FTarget4
	Public FTarget5
	Public FTarget6
	Public FTarget7
	Public FTarget8
	Public FTarget9
	Public FTarget10
	Public FTarget11
	Public FTarget12
	Public FProfit1
	Public FProfit2
	Public FProfit3
	Public FProfit4
	Public FProfit5
	Public FProfit6
	Public FProfit7
	Public FProfit8
	Public FProfit9
	Public FProfit10
	Public FProfit11
	Public FProfit12

	Public  FItemcost1
	Public  FItemcost2
	Public  FItemcost3
	Public  FItemcost4
	Public  FItemcost5
	Public  FItemcost6
	Public  FItemcost7
	Public  FItemcost8
	Public  FItemcost9
	Public  FItemcost10
	Public  FItemcost11
	Public  FItemcost12

	Public  FMaechulProfit1
	Public  FMaechulProfit2
	Public  FMaechulProfit3
	Public  FMaechulProfit4
	Public  FMaechulProfit5
	Public  FMaechulProfit6
	Public  FMaechulProfit7
	Public  FMaechulProfit8
	Public  FMaechulProfit9
	Public  FMaechulProfit10
	Public  FMaechulProfit11
	Public  FMaechulProfit12

	Public FSumTartgetMoney
	Public FSumProfitMoney
	Public FSumitemcost
	Public FSummaechulProfit

	Public  FTotTarget1
	Public  FTotTarget2
	Public  FTotTarget3
	Public  FTotTarget4
	Public  FTotTarget5
	Public  FTotTarget6
	Public  FTotTarget7
	Public  FTotTarget8
	Public  FTotTarget9
	Public  FTotTarget10
	Public  FTotTarget11
	Public  FTotTarget12

	Public  FTotProfit1
	Public  FTotProfit2
	Public  FTotProfit3
	Public  FTotProfit4
	Public  FTotProfit5
	Public  FTotProfit6
	Public  FTotProfit7
	Public  FTotProfit8
	Public  FTotProfit9
	Public  FTotProfit10
	Public  FTotProfit11
	Public  FTotProfit12

	Public  FTotItemcost1
	Public  FTotItemcost2
	Public  FTotItemcost3
	Public  FTotItemcost4
	Public  FTotItemcost5
	Public  FTotItemcost6
	Public  FTotItemcost7
	Public  FTotItemcost8
	Public  FTotItemcost9
	Public  FTotItemcost10
	Public  FTotItemcost11
	Public  FTotItemcost12

	Public  FTotmaechulProfit1
	Public  FTotmaechulProfit2
	Public  FTotmaechulProfit3
	Public  FTotmaechulProfit4
	Public  FTotmaechulProfit5
	Public  FTotmaechulProfit6
	Public  FTotmaechulProfit7
	Public  FTotmaechulProfit8
	Public  FTotmaechulProfit9
	Public  FTotmaechulProfit10
	Public  FTotmaechulProfit11
	Public  FTotmaechulProfit12
End Class

Class CMDCategory
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public FRectCatecode
	Public FRectUserid
	Public FRectyyyy
	Public FRectGubun
	Public FRectmm
	Public FLastYearArray

	Public Sub getMDCategoryRegedList
		Dim sqlStr, i, addSql
		If FRectUserid <> "" Then
			addSql = addSql & " and MC.userid = '"&FRectUserid&"' "
		End If

		If FRectCatecode <> "" Then
			If Len(FRectCatecode) = 3 Then
				addSql = addSql & " and Left(DC.catecode,3) = '"&FRectCatecode&"'"
			Else
				addSql = addSql & " and DC.catecode = '"&FRectCatecode&"'"
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_display_cate as DC "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_mdmenu_category MC on DC.catecode = MC.catecode "
		sqlStr = sqlStr & " WHERE DC.useyn = 'Y' "
		sqlStr = sqlStr & " and DC.depth <= 2 " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " DC.catecode, DC.depth, DC.catename, isnull(MC.userid, '') as userid, isnull(ut.username, '') as username "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_display_cate as DC "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_mdmenu_category MC on DC.catecode = MC.catecode "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten as ut on MC.userid = ut.userid "
		sqlStr = sqlStr & " WHERE DC.useyn = 'Y' "
		sqlStr = sqlStr & " and DC.depth <= 2 " & addSql
		sqlStr = sqlStr & " order by left(DC.catecode, 3) ASC, DC.depth ASC, DC.sortNO ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CMDCategory_item
					FItemList(i).FCatecode	= rsget("catecode")
					FItemList(i).FDepth		= rsget("depth")
					FItemList(i).FCatename	= rsget("catename")
					FItemList(i).FUserid	= rsget("userid")
					FItemList(i).FUsername	= rsget("username")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getMDPurposeRegedList
		Dim sqlStr, i, addSql, addsql2, sumaddsql
		If FRectCatecode = "" Then
			addSql = addSql & " and DC.depth = 1 "
		Else
			addSql = addSql & " and DC.depth <= 2 and Left(DC.catecode,3) = '"&FRectCatecode&"' "
		End If

		If FRectyyyy <> "" Then
			addsql2 = addsql2 & " and PU.yyyy = '"&FRectyyyy&"' "
		End If

		If FRectGubun <> "" Then
			addSql = addSql & " and MC.gubun = '"&FRectGubun&"'"
		End If

	    sqlStr = "exec db_partner.[dbo].[sp_Ten_MDCategoryTotalPrice] '" & FRectCatecode & "', '" & FRectyyyy & "', '"&FRectGubun&"' "
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
    	rsget.Open sqlStr,dbget
		If not rsget.EOF Then
			FTotalCount	= rsget("cnt")
			Set FOneItem = new CMDCategory_item
				FOneItem.FTotTarget1		= rsget("totTarget1")
				FOneItem.FTotTarget2		= rsget("totTarget2")
				FOneItem.FTotTarget3		= rsget("totTarget3")
				FOneItem.FTotTarget4		= rsget("totTarget4")
				FOneItem.FTotTarget5		= rsget("totTarget5")
				FOneItem.FTotTarget6		= rsget("totTarget6")
				FOneItem.FTotTarget7		= rsget("totTarget7")
				FOneItem.FTotTarget8		= rsget("totTarget8")
				FOneItem.FTotTarget9		= rsget("totTarget9")
				FOneItem.FTotTarget10		= rsget("totTarget10")
				FOneItem.FTotTarget11		= rsget("totTarget11")
				FOneItem.FTotTarget12		= rsget("totTarget12")

				FOneItem.FTotProfit1		= rsget("totProfit1")
				FOneItem.FTotProfit2		= rsget("totProfit2")
				FOneItem.FTotProfit3		= rsget("totProfit3")
				FOneItem.FTotProfit4		= rsget("totProfit4")
				FOneItem.FTotProfit5		= rsget("totProfit5")
				FOneItem.FTotProfit6		= rsget("totProfit6")
				FOneItem.FTotProfit7		= rsget("totProfit7")
				FOneItem.FTotProfit8		= rsget("totProfit8")
				FOneItem.FTotProfit9		= rsget("totProfit9")
				FOneItem.FTotProfit10		= rsget("totProfit10")
				FOneItem.FTotProfit11		= rsget("totProfit11")
				FOneItem.FTotProfit12		= rsget("totProfit12")

				FOneItem.FTotItemcost1		= rsget("totitemcost1")
				FOneItem.FTotItemcost2		= rsget("totitemcost2")
				FOneItem.FTotItemcost3		= rsget("totitemcost3")
				FOneItem.FTotItemcost4		= rsget("totitemcost4")
				FOneItem.FTotItemcost5		= rsget("totitemcost5")
				FOneItem.FTotItemcost6		= rsget("totitemcost6")
				FOneItem.FTotItemcost7		= rsget("totitemcost7")
				FOneItem.FTotItemcost8		= rsget("totitemcost8")
				FOneItem.FTotItemcost9		= rsget("totitemcost9")
				FOneItem.FTotItemcost10		= rsget("totitemcost10")
				FOneItem.FTotItemcost11		= rsget("totitemcost11")
				FOneItem.FTotItemcost12		= rsget("totitemcost12")

				FOneItem.FTotmaechulProfit1	= rsget("totmaechulProfit1")
				FOneItem.FTotmaechulProfit2	= rsget("totmaechulProfit2")
				FOneItem.FTotmaechulProfit3	= rsget("totmaechulProfit3")
				FOneItem.FTotmaechulProfit4	= rsget("totmaechulProfit4")
				FOneItem.FTotmaechulProfit5	= rsget("totmaechulProfit5")
				FOneItem.FTotmaechulProfit6	= rsget("totmaechulProfit6")
				FOneItem.FTotmaechulProfit7	= rsget("totmaechulProfit7")
				FOneItem.FTotmaechulProfit8	= rsget("totmaechulProfit8")
				FOneItem.FTotmaechulProfit9	= rsget("totmaechulProfit9")
				FOneItem.FTotmaechulProfit10= rsget("totmaechulProfit10")
				FOneItem.FTotmaechulProfit11= rsget("totmaechulProfit11")
				FOneItem.FTotmaechulProfit12= rsget("totmaechulProfit12")

		End If
		rsget.Close
		If FTotalCount < 1 Then Exit Sub
		sqlStr = "exec db_partner.[dbo].[sp_Ten_MDCategoryPrice] '" & FRectCatecode & "', '" & FRectyyyy & "', '"&FRectGubun&"' "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FtotalPage =  CInt(FTotalCount\FPageSize)
		If (FTotalCount \ FPageSize) <> (FTotalCount / FPageSize) Then
			FtotalPage = FtotalPage + 1
		End If
		FResultCount = rsget.RecordCount - (FPageSize * (FCurrPage - 1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CMDCategory_item
					FItemList(i).FCatecode		= rsget("catecode")
					FItemList(i).FDepth			= rsget("depth")
					FItemList(i).FCatename		= rsget("catename")

					FItemList(i).FTarget1		= rsget("target1")
					FItemList(i).FTarget2		= rsget("target2")
					FItemList(i).FTarget3		= rsget("target3")
					FItemList(i).FTarget4		= rsget("target4")
					FItemList(i).FTarget5		= rsget("target5")
					FItemList(i).FTarget6		= rsget("target6")
					FItemList(i).FTarget7		= rsget("target7")
					FItemList(i).FTarget8		= rsget("target8")
					FItemList(i).FTarget9		= rsget("target9")
					FItemList(i).FTarget10		= rsget("target10")
					FItemList(i).FTarget11		= rsget("target11")
					FItemList(i).FTarget12		= rsget("target12")

					FItemList(i).FProfit1		= rsget("profit1")
					FItemList(i).FProfit2		= rsget("profit2")
					FItemList(i).FProfit3		= rsget("profit3")
					FItemList(i).FProfit4		= rsget("profit4")
					FItemList(i).FProfit5		= rsget("profit5")
					FItemList(i).FProfit6		= rsget("profit6")
					FItemList(i).FProfit7		= rsget("profit7")
					FItemList(i).FProfit8		= rsget("profit8")
					FItemList(i).FProfit9		= rsget("profit9")
					FItemList(i).FProfit10		= rsget("profit10")
					FItemList(i).FProfit11		= rsget("profit11")
					FItemList(i).FProfit12		= rsget("profit12")

					FItemList(i).FItemcost1		= rsget("itemcost1")
					FItemList(i).FItemcost2		= rsget("itemcost2")
					FItemList(i).FItemcost3		= rsget("itemcost3")
					FItemList(i).FItemcost4		= rsget("itemcost4")
					FItemList(i).FItemcost5		= rsget("itemcost5")
					FItemList(i).FItemcost6		= rsget("itemcost6")
					FItemList(i).FItemcost7		= rsget("itemcost7")
					FItemList(i).FItemcost8		= rsget("itemcost8")
					FItemList(i).FItemcost9		= rsget("itemcost9")
					FItemList(i).FItemcost10	= rsget("itemcost10")
					FItemList(i).FItemcost11	= rsget("itemcost11")
					FItemList(i).FItemcost12	= rsget("itemcost12")

					FItemList(i).FMaechulProfit1	= rsget("maechulProfit1")
					FItemList(i).FMaechulProfit2	= rsget("maechulProfit2")
					FItemList(i).FMaechulProfit3	= rsget("maechulProfit3")
					FItemList(i).FMaechulProfit4	= rsget("maechulProfit4")
					FItemList(i).FMaechulProfit5	= rsget("maechulProfit5")
					FItemList(i).FMaechulProfit6	= rsget("maechulProfit6")
					FItemList(i).FMaechulProfit7	= rsget("maechulProfit7")
					FItemList(i).FMaechulProfit8	= rsget("maechulProfit8")
					FItemList(i).FMaechulProfit9	= rsget("maechulProfit9")
					FItemList(i).FMaechulProfit10	= rsget("maechulProfit10")
					FItemList(i).FMaechulProfit11	= rsget("maechulProfit11")
					FItemList(i).FMaechulProfit12	= rsget("maechulProfit12")

					FItemList(i).FSumTartgetMoney	= rsget("SumTartgetMoney")
					FItemList(i).FSumProfitMoney	= rsget("SumProfitMoney")
					FItemList(i).FSumitemcost		= rsget("Sumitemcost")
					FItemList(i).FSummaechulProfit	= rsget("SummaechulProfit")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	Public Sub getMDPurposeRegedListNew
		Dim sqlStr, i, addSql, addsql2, sumaddsql
		If FRectCatecode = "" Then
			addSql = addSql & " and DC.depth = 1 "
		Else
			addSql = addSql & " and DC.depth <= 2 and Left(DC.catecode,3) = '"&FRectCatecode&"' "
		End If

		If FRectyyyy <> "" Then
			addsql2 = addsql2 & " and PU.yyyy = '"&FRectyyyy&"' "
		End If

		If FRectGubun <> "" Then
			addSql = addSql & " and MC.gubun = '"&FRectGubun&"'"
		End If


		sqlStr = "exec db_partner.[dbo].[sp_Ten_MDCategoryPrice] '" & FRectCatecode & "', '" & FRectyyyy & "', '"&FRectGubun&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FResultCount = rsget.RecordCount
		
		If FResultCount < 1 Then
			rsget.close
			Exit Sub
		End IF
			
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CMDCategory_item
					FItemList(i).FCatecode		= rsget("catecode")
					FItemList(i).FDepth			= rsget("depth")
					FItemList(i).FCatename		= rsget("catename")

					FItemList(i).FTarget1		= rsget("target1")
					FItemList(i).FTarget2		= rsget("target2")
					FItemList(i).FTarget3		= rsget("target3")
					FItemList(i).FTarget4		= rsget("target4")
					FItemList(i).FTarget5		= rsget("target5")
					FItemList(i).FTarget6		= rsget("target6")
					FItemList(i).FTarget7		= rsget("target7")
					FItemList(i).FTarget8		= rsget("target8")
					FItemList(i).FTarget9		= rsget("target9")
					FItemList(i).FTarget10		= rsget("target10")
					FItemList(i).FTarget11		= rsget("target11")
					FItemList(i).FTarget12		= rsget("target12")

					FItemList(i).FProfit1		= rsget("profit1")
					FItemList(i).FProfit2		= rsget("profit2")
					FItemList(i).FProfit3		= rsget("profit3")
					FItemList(i).FProfit4		= rsget("profit4")
					FItemList(i).FProfit5		= rsget("profit5")
					FItemList(i).FProfit6		= rsget("profit6")
					FItemList(i).FProfit7		= rsget("profit7")
					FItemList(i).FProfit8		= rsget("profit8")
					FItemList(i).FProfit9		= rsget("profit9")
					FItemList(i).FProfit10		= rsget("profit10")
					FItemList(i).FProfit11		= rsget("profit11")
					FItemList(i).FProfit12		= rsget("profit12")

					FItemList(i).FItemcost1		= rsget("itemcost1")
					FItemList(i).FItemcost2		= rsget("itemcost2")
					FItemList(i).FItemcost3		= rsget("itemcost3")
					FItemList(i).FItemcost4		= rsget("itemcost4")
					FItemList(i).FItemcost5		= rsget("itemcost5")
					FItemList(i).FItemcost6		= rsget("itemcost6")
					FItemList(i).FItemcost7		= rsget("itemcost7")
					FItemList(i).FItemcost8		= rsget("itemcost8")
					FItemList(i).FItemcost9		= rsget("itemcost9")
					FItemList(i).FItemcost10	= rsget("itemcost10")
					FItemList(i).FItemcost11	= rsget("itemcost11")
					FItemList(i).FItemcost12	= rsget("itemcost12")

					FItemList(i).FMaechulProfit1	= rsget("maechulProfit1")
					FItemList(i).FMaechulProfit2	= rsget("maechulProfit2")
					FItemList(i).FMaechulProfit3	= rsget("maechulProfit3")
					FItemList(i).FMaechulProfit4	= rsget("maechulProfit4")
					FItemList(i).FMaechulProfit5	= rsget("maechulProfit5")
					FItemList(i).FMaechulProfit6	= rsget("maechulProfit6")
					FItemList(i).FMaechulProfit7	= rsget("maechulProfit7")
					FItemList(i).FMaechulProfit8	= rsget("maechulProfit8")
					FItemList(i).FMaechulProfit9	= rsget("maechulProfit9")
					FItemList(i).FMaechulProfit10	= rsget("maechulProfit10")
					FItemList(i).FMaechulProfit11	= rsget("maechulProfit11")
					FItemList(i).FMaechulProfit12	= rsget("maechulProfit12")

					FItemList(i).FSumTartgetMoney	= rsget("SumTartgetMoney")
					FItemList(i).FSumProfitMoney	= rsget("SumProfitMoney")
					FItemList(i).FSumitemcost		= rsget("Sumitemcost")
					FItemList(i).FSummaechulProfit	= rsget("SummaechulProfit")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
		
		
		sqlStr = "exec db_partner.[dbo].[sp_Ten_MDCategoryPrice] '" & FRectCatecode & "', '" & (FRectyyyy-1) & "', '"&FRectGubun&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		If Not rsget.Eof Then
			FLastYearArray = rsget.getRows()
		End If
		rsget.Close
	End Sub

	Public Sub getMDPurpose1DepthList
		Dim sqlStr, i, addSql, addsql2
		If FRectCatecode <> "" Then
			addSql = addSql & " and Left(DC.catecode,3) = '"&LEFT(FRectCatecode,3)&"' "
		End If

		If FRectyyyy <> "" Then
			addsql2 = addsql2 & " and PU.yyyy = '"&FRectyyyy&"' "
		End If

		If FRectmm <> "" Then
			If Len(FRectmm) = 1 Then
				FRectmm = "0"&FRectmm
				addsql2 = addsql2 & " and PU.mm = '"&FRectmm&"' "
			Else
				addsql2 = addsql2 & " and PU.mm = '"&FRectmm&"' "
			End If
		End If
	
		If FRectGubun <> "" Then
			addsql2 = addsql2 & " and PU.gubun = '"&FRectGubun&"'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " DC.catecode, DC.depth, DC.catename, isnull(PU.targetMoney, '') as targetMoney, isnull(PU.profitMoney, '') as profitMoney, PU.yyyy, PU.mm"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_display_cate as DC "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_mdmenu_purpose PU on DC.catecode = PU.catecode " & addsql2
		sqlStr = sqlStr & " WHERE 1=1" ''DC.useyn = 'Y' " ''사용여부 관계없이 보이게.
		sqlStr = sqlStr & " and DC.depth <= 2 " & addSql
		sqlStr = sqlStr & " order by left(DC.catecode, 3) ASC, DC.depth ASC, DC.sortNO ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FtotalPage =  CInt(FTotalCount\FPageSize)
		If (FTotalCount \ FPageSize) <> (FTotalCount / FPageSize) Then
			FtotalPage = FtotalPage + 1
		End If
		FResultCount = rsget.RecordCount - (FPageSize * (FCurrPage - 1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CMDCategory_item
					FItemList(i).FCatecode		= rsget("catecode")
					FItemList(i).FDepth			= rsget("depth")
					FItemList(i).FCatename		= rsget("catename")
					FItemList(i).FTargetMoney	= rsget("targetMoney")
					FItemList(i).FProfitMoney	= rsget("profitMoney")
					FItemList(i).FYyyy			= rsget("yyyy")
					FItemList(i).FMm			= rsget("mm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close

	End Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function
End Class

Function DrawUserIdCombo(selectBoxName,selectedId)
	Dim tmp_str, strSql, hStr
	hStr = ""
	hStr = hStr & "<select name="&selectBoxName&" id='uid' class='select'>"
	hStr = hStr & "<option value='' selected>선택</option>"
	strSql = ""
	strSql = strSql & " select D.userid, D.username from "
	strSql = strSql & " [db_partner].[dbo].tbl_partner as A "
	strSql = strSql & " INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid "
	strSql = strSql & " where A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop'  "
	strSql = strSql & " AND A.part_sn IN (11,21) "
	strSql = strSql & " ORDER BY A.part_sn ASC, A.posit_sn ASC, A.regdate ASC "
	rsget.Open strSql,dbget,1

	If not rsget.EOF Then
		Do Until rsget.EOF
			If rsget("userid") = selectedId Then
				tmp_str = " selected"
			End If
			hStr = hStr & "<option value='"&rsget("userid")&"' "&tmp_str&">" + db2html(rsget("username")) + "</option>"
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close
	hStr = hStr & "</select>"
	DrawUserIdCombo = hStr
End Function

Function DrawUserIdOption
	Dim strSql, hStr
	hStr = ""
	strSql = ""
	strSql = strSql & " select D.userid, D.username from "
	strSql = strSql & " [db_partner].[dbo].tbl_partner as A "
	strSql = strSql & " INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid "
	strSql = strSql & " where A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop'  "
	strSql = strSql & " AND A.part_sn IN (11,21) "
	strSql = strSql & " ORDER BY A.part_sn ASC, A.posit_sn ASC, A.regdate ASC "
	rsget.Open strSql,dbget,1

	If not rsget.EOF Then
		Do Until rsget.EOF
			hStr = hStr & "<option value='"&rsget("userid")&"' >" + db2html(rsget("username")) + "</option>"
			rsget.MoveNext
		Loop
	End If
	rsget.close
	hStr = hStr & "</select>"
	DrawUserIdOption = hStr
End Function

Sub DrawPurposeDateBox(yyyy)
	Dim buf,i
	buf = "<select class='select' name='yyyy'>"
    for i=2013 to Year(now) + 1
		if (CStr(i)=CStr(yyyy)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"
    response.write buf
End Sub

Function NullOrCurrFormat(oval)
	If IsNULL(oval) then
		NullOrCurrFormat = " "
	Else
		NullOrCurrFormat = FormatNumber(oval,0)
	End If
End Function

Function fnGet1DepthCode(icode)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 catename FROM  db_item.dbo.tbl_display_cate WHERE catecode = '"&icode&"' "
	rsget.Open strSql,dbget,1
	If not rsget.EOF  then
   		fnGet1DepthCode = rsget("catename")
	End If
	rsget.Close
End Function
%>
