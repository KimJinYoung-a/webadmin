<%
Class cmomoranking_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public ftitle
	public fstartdate
	public fenddate
	public fregdate
	public fisusing
	public ftitle_img
	
	public fordernum
	public fitemid
	public fitemname
	public fitemdetail
	public fitemimg1
	public fitemimg2
	public ftotalvote
	public fupvote
	
	public FRankingDetail

End Class


Class ClsMomoRanking

	public FItemList()
	public FOneItem
	public FGubun
	public FUserID
	public FIdx
	public FDIdx
	public FItemID
	public FIsUsing
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount


	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	public Sub FRankingList
		Dim sqlStr, i, vSubQuery
					 
		If FIsUsing <> "" Then
			vSubQuery = vSubQuery & " AND m.isusing = '" & FIsUsing & "' "
		End If
		
		sqlStr = "SELECT COUNT(idx) " & _
				 "		FROM [db_momo].[dbo].[tbl_ranking_master] AS m " & _
				  "	WHERE 1=1 " & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " " & _
				 "			m.idx, m.title, m.title_img, convert(varchar(10),m.startdate,120) AS startdate, convert(varchar(10),m.enddate,120) AS enddate, m.isusing, convert(varchar(10),m.regdate,120) AS regdate " & _
				 "		FROM [db_momo].[dbo].[tbl_ranking_master] AS m " & _
				 "	WHERE 1=1 " & vSubQuery & " " & _
				 "	ORDER BY m.idx DESC "
		
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
				set FItemList(i) = new cmomoranking_item
					FItemList(i).fidx		= rsget("idx")
					FItemList(i).ftitle		= db2html(rsget("title"))
					FItemList(i).ftitle_img	= rsget("title_img")
					FItemList(i).fstartdate	= rsget("startdate")
					FItemList(i).fenddate	= rsget("enddate")
					FItemList(i).fisusing	= rsget("isusing")
					FItemList(i).fregdate	= rsget("regdate")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end Sub


	public Sub FRankingMasterView
		Dim sqlStr
		sqlStr = "SELECT m.idx, m.title, m.title_img, m.startdate, m.enddate, m.isusing, m.regdate FROM [db_momo].[dbo].[tbl_ranking_master] AS m WHERE m.idx = '" & FIdx & "' "
        rsget.Open SqlStr, dbget, 1
        
        set FOneItem = new cmomoranking_item

        If Not rsget.Eof Then
			FOneItem.fidx 		= rsget("idx")
			FOneItem.ftitle		= db2html(rsget("title"))
			FOneItem.ftitle_img	= rsget("title_img")
			FOneItem.fstartdate	= rsget("startdate")
			FOneItem.fenddate	= rsget("enddate")
			FOneItem.fisusing	= rsget("isusing")
			FOneItem.fregdate	= rsget("regdate")
        End If
        rsget.Close
	end Sub
	
	
	public Sub FRankingDetailView
		Dim sqlStr
		sqlStr = "SELECT d.idx, d.ordernum, d.itemid, d.itemname, d.itemdetail, d.itemimg1, d.itemimg2, d.isusing FROM [db_momo].[dbo].[tbl_ranking_detail] AS d WHERE d.idx = '" & FDIdx & "' "
        rsget.Open SqlStr, dbget, 1
        
        set FOneItem = new cmomoranking_item

        If Not rsget.Eof Then
			FOneItem.fidx 			= rsget("idx")
			FOneItem.fordernum		= rsget("ordernum")
			FOneItem.fitemid		= rsget("itemid")
			FOneItem.fitemname		= db2html(rsget("itemname"))
			FOneItem.fitemdetail	= db2html(rsget("itemdetail"))
			FOneItem.fitemimg1		= rsget("itemimg1")
			FOneItem.fitemimg2		= rsget("itemimg2")
			FOneItem.fisusing		= rsget("isusing")
        End If
        rsget.Close
	end Sub


	public function FRankingDetailViewList
		Dim sqlStr, i, vSubQuery

		sqlStr = "SELECT " & _
				 "			d.ordernum, d.itemid, d.itemname, d.itemimg1, d.itemimg2, d.totalvote, d.upvote, d.isusing, d.idx, d.itemdetail " & _
				 "		FROM [db_momo].[dbo].[tbl_ranking_detail] AS d " & _
				 "	WHERE 1=1 AND d.masteridx = '" & FIdx & "' " & _
				 "	ORDER BY d.ordernum ASC "
		
		rsget.Open sqlStr, dbget ,1
		
		ftotalcount = rsget.RecordCount
		
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FRankingDetailViewList = rsget.getRows()
		END IF
		rsget.Close
	end function
	


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