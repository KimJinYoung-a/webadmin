<%

Class CSpecialItem

	public Fidx
	public Fmasteridx
	public Fitemid
	public Fisusing
	public FitemName
	public FImageSmall
	public Fgubun
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class COffshopMailzine
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
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

	public Function GetNewitemList()
		dim sqlStr,i
		sqlStr = "select count(d.idx) as cnt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_md_choice d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and d.masteridx='" + FRectMasteridx + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.masteridx, d.itemid, d.isusing, i.itemname,i.smallimage, m.regdate" + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_md_choice d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and d.masteridx='" + FRectMasteridx + "'"
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
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fmasteridx      = rsget("masteridx")
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