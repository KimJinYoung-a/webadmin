<%
Class cmddiary_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fmgzId
	public fmenuimg
	public fmenuimg_on
	public fmainimg
	public fusemap
	public fopendate
	public fuseyn
	public fregdate
	
End Class


Class Clsmddiary

	public FItemList()
	public FOneItem
	public FGubun
	public FUseYN
	public FMgzID
	public FMenuImg
	public FMenuImg_On
	public FMainImg
	public FOpenDate
	public FRegdate
	public FUseMap
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	
	public sub FmddiaryList
		Dim sqlStr, i, vSubQuery
		
		If FUseYN <> "" Then
			vSubQuery = vSubQuery & " AND D.useYN = '" & FUseYN & "' "
		End If
		
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_sitemaster].[dbo].[tbl_md_diary] AS D " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " D.mgzId, D.menuimg, D.menuimg_on, D.mainimg, D.usemap, D.opendate, D.useyn, Convert(varchar(10),D.regdate,120) AS regdate " & _
				 "		FROM [db_sitemaster].[dbo].[tbl_md_diary] AS D " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY D.mgzId DESC "
		
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
				set FItemList(i) = new cmddiary_oneitem

					FItemList(i).fmgzId			= rsget("mgzId")
					FItemList(i).fmenuimg		= rsget("menuimg")
					FItemList(i).fmenuimg_on	= rsget("menuimg_on")
					FItemList(i).fmainimg		= rsget("mainimg")
					FItemList(i).fusemap		= rsget("usemap")
					FItemList(i).fopendate		= rsget("opendate")
					FItemList(i).fuseyn			= rsget("useyn")
					FItemList(i).fregdate		= rsget("regdate")

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end sub
	
	
	public Function FmddiaryCont
	Dim strSql
		IF FMgzID = "" THEN Exit Function
		strSql = " SELECT D.mgzId, D.menuimg, D.menuimg_on, D.mainimg, D.usemap, D.opendate, D.useyn, D.regdate "&_
				 " 		FROM [db_sitemaster].[dbo].[tbl_md_diary] AS D "&_
				 " WHERE mgzId = '" & FMgzID & "' "
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FMenuImg 	= rsget("menuimg")
			FMenuImg_On	= rsget("menuimg_on")
			FMainImg	= rsget("mainimg")
			FOpenDate	= rsget("opendate")
			FRegdate	= rsget("regdate")
			FUseYN 		= rsget("useyn")
			FUseMap		= rsget("usemap")

		End IF
		rsget.Close
	End Function
	
	
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