<%

Class CMDChoiceItem

	public Fidx
	Public Fcdl
	Public Fcdm
	Public Fdisp1
	Public Fsubject
	Public Fregdate
	public Fitemid
	public Fisusing
	public Fcode_nm
	public Fmidcode_nm
	public FsortNo
	public FitemName
	public FImageSmall

	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold
	public FMakerid
	public Forgprice
	public Fsailyn
	public Fsailprice
	public Fitemcouponyn
	public Fitemcoupontype
	public Fitemcouponvalue

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CMDChoice
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Public FRectCDL
	Public FRectCDM
	Public FRectDisp1
	Public FRectIdx
	Public FRectIsUsing
	public FRectStyleSerail

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

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
	end function

	public Function GetMDChoiceList2015()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and c.theme_idx = '" + FRectIdx + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		if FRectDisp1 <> "" then
			sqlStr = sqlStr + " and c.dispcate1 = '" + FRectDisp1 + "'" + vbcrlf
		else
			sqlStr = sqlStr + " and c.dispcate1 is not null " + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.cdl, c.itemid, c.isusing, i.itemname, i.smallimage, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold, i.makerid, i.orgprice, i.sailyn, i.sailprice, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue " + vbcrlf
		sqlStr = sqlStr + " ,'' as code_nm " + vbcrlf
		sqlStr = sqlStr + " , isNull((select catename from db_item.dbo.tbl_display_cate where catecode = i.dispcate1 and depth = 1),'') as midcode_nm, c.cate_mid " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and c.theme_idx = '" + FRectIdx + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		if FRectDisp1 <> "" then
			sqlStr = sqlStr + " and c.dispcate1 = '" + FRectDisp1 + "'" + vbcrlf
		else
			sqlStr = sqlStr + " and c.dispcate1 is not null " + vbcrlf
		end if
		
		sqlStr = sqlStr + " order by c.sortNo, c.regdate desc"
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
				set FItemList(i) = new CMDChoiceItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cdl")
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fcdm		= rsget("cate_mid")
				FItemList(i).Fmidcode_nm	= rsget("midcode_nm")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FsortNo	= rsget("sortNo")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).Forgprice      = rsget("orgprice")
				FItemList(i).Fsailyn		= rsget("sailyn")
				FItemList(i).Fsailprice     = rsget("sailprice")
				FItemList(i).Fitemcouponyn  = rsget("itemcouponyn")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function GetMDChoiceList()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and c.theme_idx = '" + FRectIdx + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.cdl, c.itemid, c.isusing, i.itemname, i.smallimage, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold, i.makerid, i.orgprice, i.sailyn, i.sailprice, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue " + vbcrlf
		sqlStr = sqlStr + " ,'' as code_nm " + vbcrlf
		sqlStr = sqlStr + " , isNull((select catename from db_item.dbo.tbl_display_cate where catecode = i.dispcate1 and depth = 1),'') as midcode_nm, c.cate_mid " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and c.theme_idx = '" + FRectIdx + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		
		sqlStr = sqlStr + " order by c.sortNo, c.itemid desc"
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
				set FItemList(i) = new CMDChoiceItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cdl")
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fcdm		= rsget("cate_mid")
				FItemList(i).Fmidcode_nm	= rsget("midcode_nm")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FsortNo	= rsget("sortNo")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).Forgprice      = rsget("orgprice")
				FItemList(i).Fsailyn		= rsget("sailyn")
				FItemList(i).Fsailprice     = rsget("sailprice")
				FItemList(i).Fitemcouponyn  = rsget("itemcouponyn")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function
	
	
	public Function GetMDChoiceThemeList()
		dim sqlStr,i

		sqlStr = "select count(t.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice_theme t" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + vbcrlf
		
		if FRectDisp1<>"" then
			sqlStr = sqlStr + " and t.disp1 = '" + FRectDisp1 + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and t.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " t.idx, t.disp1, t.subject, t.sortno, t.isusing, t.regdate " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice_theme t " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + vbcrlf

		if FRectDisp1<>"" then
			sqlStr = sqlStr + " and t.disp1 = '" + FRectDisp1 + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and t.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		
		sqlStr = sqlStr + " order by t.sortNo asc, t.idx desc"
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
				set FItemList(i) = new CMDChoiceItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fdisp1		= rsget("disp1")
				FItemList(i).Fsubject	= db2html(rsget("subject"))
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FsortNo	= rsget("sortNo")
				FItemList(i).Fregdate	= rsget("regdate")

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
