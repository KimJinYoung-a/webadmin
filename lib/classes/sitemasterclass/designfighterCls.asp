<%
Class CDesignFighterItem

	public Fidx
	public Fitemid1
	public Fitemid2
	public Ficon1
	public Ficon2
	public Fisusing
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CDesignFighter
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectCD1
	public FRectidx

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetDesignFighterList()
		dim sqlStr,i

		sqlStr = "select count(idx) as cnt"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
		sqlStr = sqlStr + " where idx<>0" + vbcrlf
		if FRectidx <> "" then
		sqlStr = sqlStr + " and idx=" + Cstr(FRectidx) + "" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " idx, itemid1, itemid2, icon1, icon2, isusing, regdate" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
		sqlStr = sqlStr + " where idx<>0" + vbcrlf
		if FRectidx <> "" then
		sqlStr = sqlStr + " and idx=" + Cstr(FRectidx) + "" + vbcrlf
		end if
		sqlStr = sqlStr + " order by idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CDesignFighterItem

				FItemList(i).Fidx       = rsget("idx")
				FItemList(i).Fitemid1       = rsget("itemid1")
				FItemList(i).Fitemid2       = rsget("itemid2")
				FItemList(i).Ficon1 = "http://imgstatic.10x10.co.kr/contents/designfighter/icon1/" + rsget("icon1")
				FItemList(i).Ficon2 = "http://imgstatic.10x10.co.kr/contents/designfighter/icon2/" + rsget("icon2")
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).Fregdate       = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if

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

end Class

Class CDesignFighterDetail

	public Fidx
	public Fitemid1
	public Fitemid2
	Public Fitemname1
	Public Fitemname2
	public Ftitle
	public Ftitle1
	public Ftitle2
	public Ftitle3
	public Ftitle4
	Public Ftitleimg
	public Fmainimg1
	public Fmainimg2
	public Ficon1
	public Ficon2
	public Fsicon1
	public Fsicon2
	public Fbanimg
	public Fimg1_1
	public Fimg2_1
	public Fimg3_1
	public Fimg4_1
	public Fimg5_1
	public Fimg6_1
	public Fimg7_1
	public Fimg8_1
	public Fimg1_2
	public Fimg2_2
	public Fimg3_2
	public Fimg4_2
	public Fimg5_2
	public Fimg6_2
	public Fimg7_2
	public Fimg8_2
	public Fimg1_3
	public Fimg2_3
	public Fimg3_3
	public Fimg4_3
	public Fimg5_3
	public Fimg6_3
	public Fimg7_3
	public Fimg8_3
	public Fimg1_4
	public Fimg2_4
	public Fimg3_4
	public Fimg4_4
	public Fimg5_4
	public Fimg6_4
	public Fimg7_4
	public Fimg8_4
	public Fimg1_5
	public Fimg2_5
	public Fimg3_5
	public Fimg4_5
	public Fimg5_5
	public Fimg6_5
	public Fimg7_5
	public Fimg8_5
	public Fcontents1_1
	public Fcontents2_1
	public Fcontents3_1
	public Fcontents4_1
	public Fcontents5_1
	public Fcontents6_1
	public Fcontents7_1
	public Fcontents8_1
	public Fcontents1_2
	public Fcontents2_2
	public Fcontents3_2
	public Fcontents4_2
	public Fcontents5_2
	public Fcontents6_2
	public Fcontents7_2
	public Fcontents8_2
	public Fcontents1_3
	public Fcontents2_3
	public Fcontents3_3
	public Fcontents4_3
	public Fcontents5_3
	public Fcontents6_3
	public Fcontents7_3
	public Fcontents8_3
	public Fcontents1_4
	public Fcontents2_4
	public Fcontents3_4
	public Fcontents4_4
	public Fcontents5_4
	public Fcontents6_4
	public Fcontents7_4
	public Fcontents8_4
	public Fcontents1_5
	public Fcontents2_5
	public Fcontents3_5
	public Fcontents4_5
	public Fcontents5_5
	public Fcontents6_5
	public Fcontents7_5
	public Fcontents8_5
	public Fisusing
	public FRectidx
	Public Fwinyn

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetDesignFighterItem()
		dim sqlStr,i

		sqlStr = "select * from [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
		sqlStr = sqlStr + " where idx=" + Cstr(FRectidx) + "" + vbcrlf

		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
			Fidx = rsget("idx")
			Fitemid1 = rsget("itemid1")
			Fitemid2 = rsget("itemid2")
			Fitemname1 = db2html(rsget("itemname1"))
			Fitemname2 = db2html(rsget("itemname2"))			
			Ftitle1 = db2html(rsget("title1"))
			Ftitle2 = db2html(rsget("title2"))
			Ftitle3 = db2html(rsget("title3"))
			Ftitle4 = db2html(rsget("title4"))
			Ftitleimg = "/designfighter/" + rsget("titleimg")
			Fmainimg1 = "/designfighter/" + rsget("mainimg1")
			Fmainimg2 = "/designfighter/" + rsget("mainimg2")
			Ficon1 = "/designfighter/icon1/" + rsget("icon1")
			Ficon2 = "/designfighter/icon2/" + rsget("icon2")
			Fsicon1 = "/designfighter/icon1/" + rsget("sicon1")
			Fsicon2 = "/designfighter/icon2/" + rsget("sicon2")
		
			Fimg1_1 = "/designfighter/" + rsget("img1_1")
			Fimg1_2 = "/designfighter/" + rsget("img1_2")
			Fimg1_3 = "/designfighter/" + rsget("img1_3")
			Fimg1_4 = "/designfighter/" + rsget("img1_4")
			Fimg1_5 = "/designfighter/" + rsget("img1_5")
			Fimg2_1 = "/designfighter/" + rsget("img2_1")
			Fimg2_2 = "/designfighter/" + rsget("img2_2")
			Fimg2_3 = "/designfighter/" + rsget("img2_3")
			Fimg2_4 = "/designfighter/" + rsget("img2_4")
			Fimg2_5 = "/designfighter/" + rsget("img2_5")
			Fimg3_1 = "/designfighter/" + rsget("img3_1")
			Fimg3_2 = "/designfighter/" + rsget("img3_2")
			Fimg3_3 = "/designfighter/" + rsget("img3_3")
			Fimg3_4 = "/designfighter/" + rsget("img3_4")
			Fimg3_5 = "/designfighter/" + rsget("img3_5")
			Fimg4_1 = "/designfighter/" + rsget("img4_1")
			Fimg4_2 = "/designfighter/" + rsget("img4_2")
			Fimg4_3 = "/designfighter/" + rsget("img4_3")
			Fimg4_4 = "/designfighter/" + rsget("img4_4")
			Fimg4_5 = "/designfighter/" + rsget("img4_5")
			Fimg5_1 = "/designfighter/" + rsget("img5_1")
			Fimg5_2 = "/designfighter/" + rsget("img5_2")
			Fimg5_3 = "/designfighter/" + rsget("img5_3")
			Fimg5_4 = "/designfighter/" + rsget("img5_4")
			Fimg5_5 = "/designfighter/" + rsget("img5_5")
			Fimg6_1 = "/designfighter/" + rsget("img6_1")
			Fimg6_2 = "/designfighter/" + rsget("img6_2")
			Fimg6_3 = "/designfighter/" + rsget("img6_3")
			Fimg6_4 = "/designfighter/" + rsget("img6_4")
			Fimg6_5 = "/designfighter/" + rsget("img6_5")
			Fimg7_1 = "/designfighter/" + rsget("img7_1")
			Fimg7_2 = "/designfighter/" + rsget("img7_2")
			Fimg7_3 = "/designfighter/" + rsget("img7_3")
			Fimg7_4 = "/designfighter/" + rsget("img7_4")
			Fimg7_5 = "/designfighter/" + rsget("img7_5")
			Fimg8_1 = "/designfighter/" + rsget("img8_1")
			Fimg8_2 = "/designfighter/" + rsget("img8_2")
			Fimg8_3 = "/designfighter/" + rsget("img8_3")
			Fimg8_4 = "/designfighter/" + rsget("img8_4")
			Fimg8_5 = "/designfighter/" + rsget("img8_5")
			Fcontents1_1 = db2html(rsget("contents1_1"))
			Fcontents1_2 = db2html(rsget("contents1_2"))
			Fcontents1_3 = db2html(rsget("contents1_3"))
			Fcontents1_4 = db2html(rsget("contents1_4"))
			Fcontents1_5 = db2html(rsget("contents1_5"))
			Fcontents2_1 = db2html(rsget("contents2_1"))
			Fcontents2_2 = db2html(rsget("contents2_2"))
			Fcontents2_3 = db2html(rsget("contents2_3"))
			Fcontents2_4 = db2html(rsget("contents2_4"))
			Fcontents2_5 = db2html(rsget("contents2_5"))
			Fcontents3_1 = db2html(rsget("contents3_1"))
			Fcontents3_2 = db2html(rsget("contents3_2"))
			Fcontents3_3 = db2html(rsget("contents3_3"))
			Fcontents3_4 = db2html(rsget("contents3_4"))
			Fcontents3_5 = db2html(rsget("contents3_5"))
			Fcontents4_1 = db2html(rsget("contents4_1"))
			Fcontents4_2 = db2html(rsget("contents4_2"))
			Fcontents4_3 = db2html(rsget("contents4_3"))
			Fcontents4_4 = db2html(rsget("contents4_4"))
			Fcontents4_5 = db2html(rsget("contents4_5"))
			Fcontents5_1 = db2html(rsget("contents5_1"))
			Fcontents5_2 = db2html(rsget("contents5_2"))
			Fcontents5_3 = db2html(rsget("contents5_3"))
			Fcontents5_4 = db2html(rsget("contents5_4"))
			Fcontents5_5 = db2html(rsget("contents5_5"))
			Fcontents6_1 = db2html(rsget("contents6_1"))
			Fcontents6_2 = db2html(rsget("contents6_2"))
			Fcontents6_3 = db2html(rsget("contents6_3"))
			Fcontents6_4 = db2html(rsget("contents6_4"))
			Fcontents6_5 = db2html(rsget("contents6_5"))
			Fcontents7_1 = db2html(rsget("contents7_1"))
			Fcontents7_2 = db2html(rsget("contents7_2"))
			Fcontents7_3 = db2html(rsget("contents7_3"))
			Fcontents7_4 = db2html(rsget("contents7_4"))
			Fcontents7_5 = db2html(rsget("contents7_5"))
			Fcontents8_1 = db2html(rsget("contents8_1"))
			Fcontents8_2 = db2html(rsget("contents8_2"))
			Fcontents8_3 = db2html(rsget("contents8_3"))
			Fcontents8_4 = db2html(rsget("contents8_4"))
			Fcontents8_5 = db2html(rsget("contents8_5"))
			Fisusing = rsget("isusing")
			Fwinyn = rsget("winyn")
			
			Fbanimg = "/designfighter/" + rsget("banimg")
			Ftitle = db2html(rsget("title"))
		End if
		rsget.Close

	end function

end Class
%>