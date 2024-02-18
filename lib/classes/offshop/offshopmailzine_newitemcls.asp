<%

Class CSpecialItem

	public Fidx
	public Fcd1
	public Fmasteridx
	public Fitemid
	public Fisusing
	public Fcode_nm
	public FitemName
	public FImageSmall
	public Fgubun
	public Fregdate
	public FImageList
	public FSellCash


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetMdGubun()
		if FGubun = "01" then
			GetMdGubun = "MD1"
		elseif FGubun = "02" then
			GetMdGubun = "MD2"
		elseif FGubun = "03" then
			GetMdGubun = "MD3"
		elseif FGubun = "04" then
			GetMdGubun = "MD4"
		end if
	end Function

end Class

Class COffshopMailzine
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectCD1
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
		sqlStr = sqlStr + " from  [db_item].[dbo].tbl_cate_large l, [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_newitem d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = i.cate_large" + vbcrlf
		sqlStr = sqlStr + " and i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and d.masteridx='" + FRectMasteridx + "'"
		end if
		if FRectCD1<>"" then
			sqlStr = sqlStr + " and d.cd1='" + FRectCD1 + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.masteridx, d.cd1, d.itemid, d.isusing, i.itemname,i.smallimage, l.code_nm, m.regdate" + vbcrlf
		sqlStr = sqlStr + " from  [db_item].[dbo].tbl_cate_large l, [db_item].[dbo].tbl_item i," + vbcrlf
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_mail_newitem d, [db_shop].[dbo].tbl_shopmaster_mail m" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = i.cate_large" + vbcrlf
		sqlStr = sqlStr + " and i.itemid=d.itemid" + vbcrlf
		sqlStr = sqlStr + " and d.masteridx=m.idx" + vbcrlf
		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and d.masteridx='" + FRectMasteridx + "'"
		end if
		if FRectCD1<>"" then
			sqlStr = sqlStr + " and d.cd1='" + FRectCD1 + "'"
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
				FItemList(i).Fcd1      = rsget("cd1")
				FItemList(i).Fitemid       = rsget("itemid")
				FItemList(i).Fisusing      = rsget("isusing")
				FItemList(i).FitemName   = db2html(rsget("itemname"))
				FItemList(i).Fcode_nm   = db2html(rsget("code_nm"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")
				FItemList(i).Fregdate      = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Sub GetPreNewItem()
		dim sqlStr,i

		sqlStr = "select n.itemid , i.sellcash, i.itemname, i.listimage, n.cd1" + vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_mail_newitem n, [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "where n.itemid=i.itemid " + vbcrlf

		if FRectMasteridx<>"" then
			sqlStr = sqlStr + " and n.masteridx='" + CStr(FRectMasteridx) + "'"
		end if
		sqlStr = sqlStr + " order by n.cd1 asc"

		'response.write sqlStr
		'dbget.close()	:	response.End

		rsget.open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FitemList(FResultCount)
		i=0
		if  not rsget.EOF  then

		do until rsget.eof
			set FItemList(i) = new CSpecialItem

				FItemList(i).FItemid       = rsget("itemid")
				FItemList(i).FItemName   = db2html(rsget("itemname"))
				FItemList(i).FSellCash     = rsget("sellcash")
				FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerName(i) + "/" +rsget("listimage")
				FItemList(i).Fcd1= rsget("cd1")
			i=i+1
			rsget.movenext

		loop
		end if

	rsget.close
	end Sub

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