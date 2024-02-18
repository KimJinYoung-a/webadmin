<%
'###########################################################
' History : 2007.11.09 한용민 생성
'###########################################################

Class MDChoiceItem

	public Fidx
	public Fmakerid
	public Ftitleimgurl
	Public Fmodelitem
	Public FImgSmall
	Public Fcdl
	Public Fitemid
	Public Fitemname

	public Function GetCD1Name()
		if Fcdl = "010" then
			GetCD1Name = "디자인문구"
		elseif Fcdl = "020" then
			GetCD1Name = "오피스/개인소품"
		elseif Fcdl = "030" then
			GetCD1Name = "키덜트얼리/취미"
		elseif Fcdl = "040" then
			GetCD1Name = "가구/패브릭"
		elseif Fcdl = "050" then
			GetCD1Name = "조명/데코"
		elseif Fcdl = "060" then
			GetCD1Name = "주방/욕실"
		elseif Fcdl = "070" then
			GetCD1Name = "가방/슈즈/쥬얼리"
		elseif Fcdl = "080" then
			GetCD1Name = "Women"
		elseif Fcdl = "090" then
			GetCD1Name = "Men"
		elseif Fcdl = "100" then
			GetCD1Name = "베이비"
		elseif Fcdl = "110" then
			GetCD1Name = "감성채널"
		elseif Fcdl = "999" then
			GetCD1Name = "전시안함"						
			
		end if
	end Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class MDChoice
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	Public FRectCDL


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

	public function GetImageFolerName(byval itemid)
		if Not(itemid="" or isNull(itemid)) then
			'GetImageFolerName = "0" + CStr(Clng(itemid\10000))
			GetImageFolerName = GetImageSubFolderByItemid(itemid)
		end if
	end function

	public function GetImageFolerName2(byval i)
		'GetImageFolerName2 = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName2 = GetImageSubFolderByItemid(FItemList(i).FItemID)
	end function

	public Function GetMDChoiceBrand()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from [db_contents].[dbo].tbl_idx_mdchoice_brand"
		sqlStr = sqlStr + " where idx<>0"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.makerid, c.titleimgurl, c.modelitem, c.modelimg" + vbcrlf
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_idx_mdchoice_brand b,"
		sqlStr = sqlStr + " [db_user].dbo.tbl_user_c c" + vbcrlf
		sqlStr = sqlStr + " where b.makerid = c.userid" + vbcrlf
		sqlStr = sqlStr + " order by b.idx desc"
'response.write sqlStr
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
				set FItemList(i) = new MDChoiceItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fmakerid      = rsget("makerid")
				FItemList(i).Ftitleimgurl       = "http://webimage.10x10.co.kr/image/brandlogo/" + rsget("titleimgurl")
				FItemList(i).Fmodelitem       =  rsget("modelitem")
				FItemList(i).FImgSmall       = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(FItemList(i).Fmodelitem) + "/" + rsget("modelimg")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function
	
	public Function GetCategoryLeftBrandRank()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from [db_contents].[dbo].tbl_category_left_brand_rank"
		sqlStr = sqlStr + " where idx<>0"
		If FRectCDL <> "" then
		sqlStr = sqlStr + " and cdl = '" + CStr(FRectCDL) + "'"
		End If

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.makerid, c.titleimgurl, c.modelitem, c.modelimg" + vbcrlf
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_category_left_brand_rank b,"
		sqlStr = sqlStr + " [db_user].dbo.tbl_user_c c" + vbcrlf
		sqlStr = sqlStr + " where b.makerid = c.userid" + vbcrlf
		If FRectCDL <> "" then
		sqlStr = sqlStr + " and b.cdl = '" + CStr(FRectCDL) + "'"
		End If
		sqlStr = sqlStr + " order by b.idx desc"
'response.write sqlStr
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
				set FItemList(i) = new MDChoiceItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fcdl      = rsget("cdl")
				FItemList(i).Fmakerid      = rsget("makerid")
				FItemList(i).Ftitleimgurl       = "http://webimage.10x10.co.kr/image/brandlogo/" + rsget("titleimgurl")
				FItemList(i).Fmodelitem       =  rsget("modelitem")
				FItemList(i).FImgSmall       = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(FItemList(i).Fmodelitem) + "/" + rsget("modelimg")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function
	
	public Function GetCategoryLeftNewItemBrand()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from [db_contents].[dbo].tbl_category_left_newitem_brand"
		sqlStr = sqlStr + " where idx<>0"
		If FRectCDL <> "" then
		sqlStr = sqlStr + " and cdl = '" + CStr(FRectCDL) + "'"
		End If

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.itemid,b.itemname, b.makerid, b.imgsmall, c.titleimgurl, c.modelitem, c.modelimg" + vbcrlf
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_category_left_newitem_brand b,"
		sqlStr = sqlStr + " [db_user].dbo.tbl_user_c c" + vbcrlf
		sqlStr = sqlStr + " where b.makerid = c.userid" + vbcrlf
		If FRectCDL <> "" then
		sqlStr = sqlStr + " and b.cdl = '" + CStr(FRectCDL) + "'"
		End If
		sqlStr = sqlStr + " order by b.idx desc"
'response.write sqlStr
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
				set FItemList(i) = new MDChoiceItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fcdl      = rsget("cdl")
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fmakerid      = db2html(rsget("makerid"))
				FItemList(i).FImgSmall      = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName2(i) + "/" + rsget("imgsmall")
				FItemList(i).Ftitleimgurl       = "http://webimage.10x10.co.kr/image/brandlogo/" + rsget("titleimgurl")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>