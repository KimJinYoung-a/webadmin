<%

Class OnOffShopItem

	public Fidx
	public Fmasteridx
	public Fitemid
	public Fisusing
	public FitemName
	public FImageSmall
	public FGubun
	public Fregdate

	public FImageList
	public FImageMain
	public FImageBasic
	public FSellCash

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class COnOffShopMailzine
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectgubun
	public FRectMasteridx

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
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public Function GetBestitemList()
		dim sqlStr,i
		sqlStr = "select count(d.idx) as cnt"
		sqlStr = sqlStr + " from  [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_bestitem d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and d.masteridx='" + FRectMasteridx + "'"
		end if
		if FRectgubun<>"" then
			sqlStr = sqlStr + " and d.gubun='" + FRectgubun + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.masteridx, d.gubun, d.itemid, d.isusing, i.itemname,i.smallimage, m.regdate" + vbcrlf
		sqlStr = sqlStr + " from  [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_bestitem d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and d.masteridx='" + CStr(FRectMasteridx) + "'"
		end if
		if FRectgubun<>"" then
			sqlStr = sqlStr + " and d.gubun='" + CStr(FRectgubun) + "'"
		end if
		sqlStr = sqlStr + " order by d.idx desc"

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
				set FItemList(i) = new OnOffShopItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fmasteridx      = rsget("masteridx")
				FItemList(i).FGubun     = rsget("gubun")
				FItemList(i).Fitemid       = rsget("itemid")
				FItemList(i).Fisusing      = rsget("isusing")
				FItemList(i).FitemName   = db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")
				FItemList(i).Fregdate      = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function


	public sub GetOnOffBest()
        dim sqlStr, i

		sqlStr = "select top 20 n.itemid , i.sellcash, i.itemname, i.listimage, i.smallimage, i.mainimage, i.basicimage" +vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_mail_bestitem n, [db_item].[dbo].tbl_item i " +vbcrlf
		sqlStr = sqlStr + " where n.itemid = i.itemid"+vbcrlf

		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and n.masteridx='" + CStr(FRectMasteridx) + "'"
		end if
		if FRectgubun<>"" then
			sqlStr = sqlStr + " and n.gubun='" + CStr(FRectgubun) + "'"
		end if
		'response.write sqlStr
		'dbget.close()	:	response.End
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new OnOffShopItem

	            FItemList(i).FItemID       = rsget("itemid")
	            FItemList(i).FItemName     = db2html(rsget("itemname"))

	            FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerName(i) + "/" +rsget("listimage")

	            FItemList(i).FImageMain    = "http://webimage.10x10.co.kr/image/main/" + GetImageFolerName(i) + "/" +rsget("mainimage")
	            FItemList(i).FImageBasic   = "http://webimage.10x10.co.kr/image/basic/" + GetImageFolerName(i) + "/" +rsget("basicimage")
	            'FItemList(i).FImageMain    = "http://webimage.10x10.co.kr/image/main/" + GetImageFolerName(i) + "/" +rsget("mainimage")
	            'FItemList(i).FImageBasic   = "http://webimage.10x10.co.kr/image/basic/" + GetImageFolerName(i) + "/" +rsget("basicimage")

				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")
				FItemList(i).FSellCash = rsget("sellcash")

				rsget.movenext
				i=i+1
			loop
		end if
		'response.write FItemList(0).Fitemid
		'dbget.close()	:	response.End
		rsget.close

    end sub

	public Function GetMDitemList(byval idx)
		dim sqlStr,i
		sqlStr = "select top 6 " + vbcrlf
		sqlStr = sqlStr + " d.itemid, i.listimage, i.sellcash,i.itemname" + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_md_choice d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx='" + idx + "'"
		sqlStr = sqlStr + " order by d.idx desc"

		'response.write sqlStr
		'dbget.close()	:	response.End
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new OnOffShopItem

				FItemList(i).Fitemid       = rsget("itemid")
				FItemList(i).FitemName   = db2html(rsget("itemname"))
				FItemList(i).Fsellcash     = rsget("sellcash")
				FItemList(i).FImageList      = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerName(i) + "/" +rsget("listimage")

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

	function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function
end Class


%>