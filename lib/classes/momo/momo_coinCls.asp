<%
Class cmymomo_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fcoin
	public fcurrentcoin
	public fuserid
	public fgubuncd
	public fgubuntitle
	public fregcount
	public fregdate
	public fid
	public fidx
	public fuseyn
	public fgubun
	public fmng_idx
	public ftype
	public fitem
	public foption
	public fitemname
	public fitem_desc
	public fimagesmall
	public fitemoption
	public foptionname
	public foutputdate
	public fsongjangno
	public fetc
	public fsavecoin
	public fnowcoin
	public fcdcount
End Class


Class cmymomo_list

	public FItemList()
	public fuserid
	
End Class


Class ClsMomoCoin

	public FItemList()
	public FOneItem
	public FGubun
	public FGubun2
	public FUserID
	public FIdx
	public FItemID
	public FMngIdx
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount
	public FPlusMinus
	public FSDate
	public FEDate
	public FSort
	public FDeljikwon

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	public Sub FBonusCoinList
		Dim sqlStr, i, vSubQuery
					 
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND userid = '" & FUserID & "' "
		End If
		
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_log] " & _
				  "	WHERE gubuncd = 13 " & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " id, userid, coin, gubun, convert(varchar(10),regdate,120) AS regdate " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_log] " & _
				 "	WHERE gubuncd = 13 " & vSubQuery & " " & _
				 "	ORDER BY id DESC "
		
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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fid		= rsget("id")
					FItemList(i).fuserid	= rsget("userid")
					FItemList(i).fcoin		= rsget("coin")
					FItemList(i).fgubun		= rsget("gubun")
					FItemList(i).fregdate	= rsget("regdate")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub
	
	
	public Sub FCoinMngList
		Dim sqlStr, i, vSubQuery
		
		'### 나중에 생기면 쓰려고 임시로 만들어둠
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " "
		End If
		'### 나중에 생기면 쓰려고 임시로 만들어둠
		
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_manage] AS M " & _
				 "	WHERE 1=1 " & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " M.idx, M.coin, M.useyn, convert(varchar(10),M.regdate,120) AS regdate " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_manage] AS M " & _
				 "	WHERE 1=1 " & vSubQuery & " " & _
				 "	ORDER BY M.useyn DESC, M.coin ASC "
		
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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fidx		= rsget("idx")
					FItemList(i).fcoin		= rsget("coin")
					FItemList(i).fuseyn		= rsget("useyn")
					FItemList(i).fregdate	= rsget("regdate")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub


	public Sub FCoinMngView
		Dim sqlStr
		sqlStr = "SELECT M.idx, M.coin, M.useyn, M.regdate FROM [db_momo].[dbo].[tbl_coin_manage] AS M WHERE M.idx = '" & FIdx & "' "
        rsget.Open SqlStr, dbget, 1
        
        set FOneItem = new cmymomo_oneitem

        If Not rsget.Eof Then
			FOneItem.fcoin = rsget("coin")
			FOneItem.fuseyn = rsget("useyn")
			FOneItem.fregdate = rsget("regdate")
        End If
        rsget.Close
	end Sub
	
	
	public Sub FCoinMngItemList
		Dim sqlStr, i, vSubQuery
		
		'### 나중에 생기면 쓰려고 임시로 만들어둠
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " "
		End If
		'### 나중에 생기면 쓰려고 임시로 만들어둠
		
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_manage_prod] AS I " & _
				 "	WHERE I.mng_idx = '" & FMngIdx & "' " & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " I.idx, I.mng_idx, I.type, I.prod, I.prod_option, I.prod_desc, I.useyn " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_manage_prod] AS I " & _
				 "	WHERE I.mng_idx = '" & FMngIdx & "' " & vSubQuery & " " & _
				 "	ORDER BY M.idx DESC "
		
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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fidx = rsget("idx")
					FItemList(i).ftype = rsget("type")
					FItemList(i).fitem = rsget("prod")
					FItemList(i).foption = rsget("prod_option")
					FItemList(i).fitem_desc = rsget("prod_desc")
					FItemList(i).fuseyn = rsget("useyn")
					FItemList(i).fmng_idx = rsget("mng_idx")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub
	
	
	public Sub FCoinMngItemView
		Dim sqlStr
		sqlStr = "SELECT I.idx, I.mng_idx, I.type, I.prod, I.prod_option, I.prod_desc, I.useyn FROM [db_momo].[dbo].[tbl_coin_manage_prod] AS I WHERE I.idx = '" & FIdx & "' "
        rsget.Open SqlStr, dbget, 1
        
        set FOneItem = new cmymomo_oneitem

        If Not rsget.Eof Then
			FOneItem.fidx = rsget("idx")
			FOneItem.ftype = rsget("type")
			FOneItem.fitem = rsget("prod")
			FOneItem.foption = rsget("prod_option")
			FOneItem.fitem_desc = rsget("prod_desc")
			FOneItem.fuseyn = rsget("useyn")
			FOneItem.fmng_idx = rsget("mng_idx")
        End If
        rsget.Close
	end Sub


	public Sub FItem_List
		Dim sqlStr, i, vSubQuery
					 
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND userid = '" & FUserID & "' "
		End If
		
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_manage_item] AS C " & _
				 "	INNER JOIN [db_item].[dbo].[tbl_item] AS I ON C.itemid = I.itemid " & _
				 "	WHERE 1=1 " & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " C.itemid, C.useyn, I.itemname, I.smallimage " & _
				 "		FROM [db_momo].[dbo].[tbl_coin_manage_item] AS C " & _
				 "	INNER JOIN [db_item].[dbo].[tbl_item] AS I ON C.itemid = I.itemid " & _
				 "	WHERE 1=1 " & vSubQuery & " "
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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fitem	= rsget("itemid")
					FItemList(i).fuseyn	= rsget("useyn")
					FItemList(i).fitemname = rsget("itemname")
					FItemList(i).fimagesmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).fitem) + "/" + rsget("smallimage")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub
	
	
	public Sub FItemOption_List
		Dim sqlStr, i, vSubQuery
					 
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND userid = '" & FUserID & "' "
		End If
		
		sqlStr = "SELECT O.itemoption, O.optionname " & _
				 "		FROM [db_item].[dbo].[tbl_item_option] AS O " & _
				 "	WHERE O.isusing = 'Y' AND O.itemid = '" & FItemID & "' " & vSubQuery & " "
		
		rsget.Open sqlStr, dbget ,1
		rsget.pagesize = FPageSize
		ftotalcount = rsget.recordcount
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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fitemoption	= rsget("itemoption")
					FItemList(i).foptionname	= rsget("optionname")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub
	
	
	public Sub FCoinUseList
		Dim sqlStr, i, vSubQuery
					 
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND userid = '" & FUserID & "' "
		End If
		
		If FGubun = "c" Then
			sqlStr = "SELECT COUNT(*) " & _
					 "		FROM [db_momo].[dbo].[tbl_coin_log] AS L " & _
					 "	WHERE gubuncd = '11' " & vSubQuery & " "
			rsget.Open sqlStr, dbget ,1
			ftotalcount = rsget(0)
			rsget.Close
			
			sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " L.id, L.userid, L.coin, L.gubun AS itemname, L.regdate AS orderdate, L.etc " & _
					 "		FROM [db_momo].[dbo].[tbl_coin_log] AS L " & _
					 "	WHERE gubuncd = '11' " & vSubQuery & " " & _
					 "	ORDER BY L.id DESC "
			rsget.Open sqlStr, dbget ,1
		Else
			sqlStr = "SELECT COUNT(*) " & _
					 "		FROM [db_momo].[dbo].[tbl_momo_order] AS O " & _
					 "	WHERE 1=1 " & vSubQuery & " "
			rsget.Open sqlStr, dbget ,1
			ftotalcount = rsget(0)
			rsget.Close
			
			sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " O.orderid AS id, O.userid, O.coin, O.itemname, O.itemoption_name, O.orderdate, O.outputdate, O.songjangno " & _
					 "		FROM [db_momo].[dbo].[tbl_momo_order] AS O " & _
					 "	WHERE 1=1 " & vSubQuery & " " & _
					 "	ORDER BY O.orderid DESC "
			rsget.Open sqlStr, dbget ,1
		End If

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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fidx			= rsget("id")
					FItemList(i).fuserid		= rsget("userid")
					FItemList(i).fcoin			= rsget("coin")
					FItemList(i).fitemname 		= rsget("itemname")
					FItemList(i).fregdate		= rsget("orderdate")
					If FGubun = "c" Then
						FItemList(i).fetc			= rsget("etc")
					Else
						FItemList(i).foutputdate	= rsget("outputdate")
						FItemList(i).foptionname	= rsget("itemoption_name")
						FItemList(i).fsongjangno	= rsget("songjangno")
					End IF
					'FItemList(i).fimagesmall 	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).fitem) + "/" + rsget("smallimage")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub
	
	
	public Sub FCoinLogList
		Dim sqlStr, i, vSubQuery
					 
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND L.userid = '" & FUserID & "' "
		End If
		
		
		If FPlusMinus <> "" Then
			If FPlusMinus = "p" Then
				vSubQuery = vSubQuery & " AND Left(L.coin,1) <> '-' AND L.regdate Between '" & FSDate & "' AND '" & FEDate & "' "
			Else
				vSubQuery = vSubQuery & " AND Left(L.coin,1) = '-' AND L.regdate Between '" & FSDate & "' AND '" & FEDate & "' "
			End IF
		End If
		
		If FDeljikwon = "o" Then
			vSubQuery = vSubQuery & " AND L.userid NOT IN('tozzinet') "
		End If
		
		If FSort = "now" Then
			FSort = "(C.savecoin-C.currentcoin)"
		ElseIf FSort = "use" Then
			FSort = "C.currentcoin"
		ElseIf FSort = "save" Then
			FSort = "C.savecoin"
		End IF

		sqlStr = "	SELECT COUNT(*) FROM ( " & _
				 "			SELECT L.userid " & _
				 "					FROM [db_momo].[dbo].[tbl_coin_log] AS L " & _
				 "				INNER JOIN [db_momo].[dbo].[tbl_coin_current] AS C ON L.userid = C.userid " & _
				 "				WHERE 1=1 " & vSubQuery & " " & _
				 "				GROUP BY L.userid, C.savecoin, C.currentcoin " & _
				 "		) AS A "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " L.userid, C.savecoin, C.currentcoin " & _
				 "		 FROM [db_momo].[dbo].[tbl_coin_log] AS L " & _
				 "	INNER JOIN [db_momo].[dbo].[tbl_coin_current] AS C ON L.userid = C.userid " & _
				 "		WHERE 1=1 " & vSubQuery & " " & _
				 "	GROUP BY L.userid, C.savecoin, C.currentcoin " & _
				 "	ORDER BY " & FSort & " DESC "
		rsget.Open sqlStr, dbget ,1
'response.write sqlStr

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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fuserid		= rsget("userid")
					FItemList(i).fsavecoin		= rsget("savecoin")
					FItemList(i).fcurrentcoin	= rsget("currentcoin")
					FItemList(i).fnowcoin		= rsget("savecoin") - rsget("currentcoin")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub
	
	
	public Sub FUserCoinLogList
		Dim sqlStr, i, vSubQuery, vSubQuery2
		
		FPageSize = 100
		FCurrPage = 1
		
		If FGubun2 = "prodcoupon" Then
			vSubQuery2 = "L.gubuncd IN(11,12)"
		Else
			vSubQuery2 = "L.gubuncd NOT IN(11,12)"
		End If
		
		If FGubun = "corner" Then
			sqlStr = "SELECT " & _
					 "		 '' AS userid, L.gubuncd, COUNT(L.gubuncd) AS cdcount, SUM(L.coin) AS coinsum " & _
					 "	FROM [db_momo].[dbo].[tbl_coin_log] AS L " & _
					 "	WHERE 1=1 AND " & vSubQuery2 & " AND L.userid NOT IN('tozzinet') " & _			 
					 "	GROUP BY L.gubuncd " & _
					 "	ORDER BY L.gubuncd ASC "
		Else
			If FUserID <> "" Then
				vSubQuery = vSubQuery & " AND L.userid = '" & FUserID & "' "
			End If
			
			sqlStr = "SELECT " & _
					 "		 L.userid, L.gubuncd, COUNT(L.gubuncd) AS cdcount, SUM(L.coin) AS coinsum " & _
					 "	FROM [db_momo].[dbo].[tbl_coin_log] AS L " & _
					 "	WHERE 1=1 " & vSubQuery & " AND " & vSubQuery2 & " " & _
					 "	GROUP BY L.userid, L.gubuncd " & _
					 "	ORDER BY L.gubuncd ASC "
		End IF
		rsget.Open sqlStr, dbget ,1
		FTotalCount = rsget.RecordCount

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
				set FItemList(i) = new cmymomo_oneitem
					FItemList(i).fuserid		= rsget("userid")
					FItemList(i).fgubuncd		= rsget("gubuncd")
					FItemList(i).fgubuntitle	= GubunTitle(rsget("gubuncd"))
					FItemList(i).fcdcount		= rsget("cdcount")
					FItemList(i).fcoin			= rsget("coinsum")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub





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


Function OptionList(vItemID)
	Dim sqlStr, vTemp
	vTemp = ""
	sqlStr = "SELECT O.itemoption, O.optionname " & _
			 "		FROM [db_item].[dbo].[tbl_item_option] AS O " & _
			 "	WHERE O.isusing = 'Y' AND O.itemid = '" & vItemID & "' "
	rsget.Open sqlStr, dbget ,1
	If Not rsget.Eof Then
		Do Until rsget.Eof
			vTemp = vTemp & "[" & rsget(0) & "]" & rsget(1) & " , "
		rsget.MoveNext
		Loop
	End If
	rsget.Close
	If vTemp = "" Then
		OptionList = "옵션 없음"
	Else
		OptionList = vTemp
	End If
End Function


Function MomoItemType(vType)
	Dim vName
	Select Case vType
		Case "i"
			vName = "상품"
		Case "c"
			vName = "쿠폰"
		Case "s"
			vName = "Secret 선물"
		Case Else
			vName = ""
	End Select
	MomoItemType = vName
End Function


Function GubunTitle(gubun)
	Select Case gubun
		Case 1
			GubunTitle = "감성사전(25)"
		Case 2
			GubunTitle = "YES or NO(2)"
		Case 3
			GubunTitle = "감성포토(3)"
		Case 4
			GubunTitle = "타블로이드(5)"
		Case 5
			GubunTitle = "아이디어 파일(3)"
		Case 6
			GubunTitle = "우리끼리 공감질문(2)"
		Case 7
			GubunTitle = "나의 인기 북마크(1)"
		Case 8
			GubunTitle = "감성소설(10)"
		Case 9
			GubunTitle = "모모 다디어리 참여(10)"
		Case 10
			GubunTitle = "모모 다이어리 댓글(1)"
		Case 11
			GubunTitle = "쿠폰교환"
		Case 12
			GubunTitle = "상품교환"
		Case 13
			GubunTitle = "보너스 코인"
		Case 14
			GubunTitle = "감성예보(10)"
		Case 15
			GubunTitle = "감성예보 한주 참여코인"
		Case 16
			GubunTitle = "한줄낙서"
		Case 17
			GubunTitle = "감성랭킹(50)"
		Case 18
			GubunTitle = "감성놀이"
		Case Else
			GubunTitle = ""
	End Select
End Function


Function MomoTotalCount(gununcd,coin)
	Dim vTemp
	SELECT CASE gununcd
		CASE 1
			vTemp = CLng(coin)/2
		CASE 2
			vTemp = CLng(coin)/2
		CASE 3
			vTemp = CLng(coin)/3
		CASE 4
			vTemp = CLng(coin)/5
		CASE 5
			vTemp = CLng(coin)/3
		CASE 6
			vTemp = CLng(coin)/2
		CASE 7
			vTemp = CLng(coin)/1
		CASE 8
			vTemp = CLng(coin)/1
		CASE 10
			vTemp = CLng(coin)/1
		CASE ELSE
			vTemp = ""
	End SELECT
	MomoTotalCount = vTemp
End Function
%>