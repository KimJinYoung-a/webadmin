<%
'###############################################
' PageName : CategoryCls.asp
' Discription : 카테고리 변경 관련 클래스
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'           2008.03.29 허진원 : 카테고리 관련 키워드 추가
'           2008.03.31 허진원 : 카테고리 탑키워드/빅찬스 추가
'           2008.04.02 허진원 : 카테고리 베스트브랜드 추가
'           2008.10.27 허진원 : 중카테고리 처리 추가
'			2009.04.16 허진원 : 관련이미지 처리 추가
'###############################################


Class COptionManagerItem
	public Fkeyword

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CCatemanageItem
	public Fcdlarge
	public Fcdmid
	public Fcdsmall
	public Fchannel
	public Fnmlarge
	public Fcatecnt
	public ForderNo

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CSimpleItem
	public FItemId
	public FItemName
	public Fmakerid
	public FImgSmall
	public FSellyn
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCatemanager
	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FtotalPage
	public FRectDispSailYN
	public FRectArrItemid
	public FRectMakerid

	public Sub GetCategoryKeyword(cdl,cdm,cds)
		dim sqlStr,i
		sqlStr = "select top 1 keyword from [db_item].dbo.tbl_Cate_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " and code_small='" + cds + "'"

		rsget.Open sqlStr, dbget, 1

		redim preserve FItemList(0)

		'값이 있던 없던 클랙스 배열 선언
		set FItemList(0) = new COptionManagerItem

		if not rsget.Eof then
			FItemList(0).Fkeyword = rsget("keyword")
		else
			FItemList(0).Fkeyword = ""
		end if

		rsget.close
	end sub

	public sub GetOrgCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_item].dbo.tbl_Cate_large"
		sqlStr = sqlStr + " order by code_large"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOrgCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm from [db_item].dbo.tbl_Cate_mid"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " order by code_mid"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOrgCateMasterSmall(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select s.code_large, s.code_mid, s.code_small, s.code_nm, IsNULL(T.cnt,0) as catecnt"
		sqlStr = sqlStr + " from [db_item].dbo.tbl_Cate_small s"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select cate_small, count(itemid) as cnt from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " 	where cate_large='" + cdl + "'"
		sqlStr = sqlStr + " 	and cate_mid='" + cdm + "'"
		sqlStr = sqlStr + "		group by cate_small"
		sqlStr = sqlStr + "	) as T on s.code_small=T.cate_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " order by code_small"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fcdsmall          = rsget("code_small")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))
				FItemList(i).Fcatecnt        = rsget("catecnt")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOrgCateItemList(cdl,cdm,cds)
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " where i.cate_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.cate_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.cate_small='" + cds + "'"

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallImage")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub

	public sub GetOrgCateNotMachItemList()
		dim sqlStr,i
		sqlStr = "select count(itemid), CEILING(CAST(Count(itemid) AS FLOAT)/" & FPageSize & ") from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FTotalPage	= rsget(1)
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + "  order by i.itemid desc "

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallImage")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub
	
	'//common/offshop/localeitem/popshopjumunitem_locale.asp
	public sub GetNewCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_item].dbo.tbl_Cate_large"
		sqlStr = sqlStr + " order by orderNo, code_large"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub
	
	'//common/offshop/localeitem/popshopjumunitem_locale.asp
	public sub GetNewCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm,orderNo from [db_item].dbo.tbl_Cate_mid"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " order by orderNo ,code_mid"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))
				FItemList(i).FOrderNo				=rsget("orderNo")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'//common/offshop/localeitem/popshopjumunitem_locale.asp
	public sub GetNewCateMasterSmall(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select s.code_large, s.code_mid, s.code_small, s.code_nm, orderNo ,IsNULL(T.cnt,0) as catecnt"
		sqlStr = sqlStr + " from [db_item].dbo.tbl_Cate_small s"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select cate_small, count(itemid) as cnt from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " 	where cate_large='" + cdl + "'"
		sqlStr = sqlStr + " 	and cate_mid='" + cdm + "'"
		sqlStr = sqlStr + "		group by cate_small"
		sqlStr = sqlStr + "	) as T on s.code_small=T.cate_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " order by orderNo, code_small"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fcdsmall          = rsget("code_small")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))
				FItemList(i).Fcatecnt        = rsget("catecnt")
				FItemList(i).FOrderNo        = rsget("orderNo")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public function GetNewCateCurrentPos(cdl,cdm,cds)
		dim sqlStr
		sqlStr = "select distinct top 1 code_nm "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_Cate_large"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		rsget.Open sqlStr, dbget, 1
		if not rsget.Eof then
			GetNewCateCurrentPos = db2html(rsget("code_nm"))
		end if
		rsget.close


		if cdm<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_item].dbo.tbl_Cate_mid"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			rsget.Open sqlStr, dbget, 1
			if not rsget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" +  db2html(rsget("code_nm"))
			end if
			rsget.close
		end if

		if cds<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_item].dbo.tbl_Cate_small"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			sqlStr = sqlStr + " and code_small='" + cds + "'"
			rsget.Open sqlStr, dbget, 1
			if not rsget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" + db2html(rsget("code_nm"))
			end if
			rsget.close
		end if

	end function

	public sub GetNewCateItemList(cdl,cdm,cds)
		dim sqlStr,addSql, i

		addSql = ""
		if FRectDispSailYN = "on" then
			addSql = addSql + " and i.sellyn='Y'"
		end if

		FRectArrItemid = trim(FRectArrItemid)
		if right(FRectArrItemid,1)="," then FRectArrItemid=left(FRectArrItemid,len(FRectArrItemid)-1)
		if FRectArrItemid<>"" then
			addSql = addSql + " and i.itemid in (" & FRectArrItemid & ")"
		end if
		if FRectMakerid<>"" then
			addSql = addSql + " and i.makerid='" & FRectMakerid & "'"
		end if

		sqlStr = "select count(i.itemid), CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " where i.cate_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.cate_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.cate_small='" + cds + "'" & addSql

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.close

		'지정페이지가 전체 페이지보다 크면 마지막 페이지 지정
		if Cint(FCurrPage)>Cint(FTotalPage) and Cint(FTotalPage)>0 then
			FCurrPage = FTotalPage
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " where i.cate_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.cate_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.cate_small='" + cds + "'" & addSql
		sqlStr = sqlStr + " order by i.itemid desc "


		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallImage")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub

	'// 추가 카테고리 상품 목록
	public sub GetAddCateItemList(cdl,cdm,cds)
		dim sqlStr,addSql, i

		addSql = ""
		if FRectDispSailYN = "on" then
			addSql = addSql + " and i.sellyn='Y'"
		end if

		FRectArrItemid = trim(FRectArrItemid)
		if right(FRectArrItemid,1)="," then FRectArrItemid=left(FRectArrItemid,len(FRectArrItemid)-1)
		if FRectArrItemid<>"" then
			addSql = addSql + " and i.itemid in (" & FRectArrItemid & ")"
		end if
		if FRectMakerid<>"" then
			addSql = addSql + " and i.makerid='" & FRectMakerid & "'"
		end if

		sqlStr = "select count(i.itemid), CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " 	join [db_item].dbo.tbl_item_Category c"
		sqlStr = sqlStr + " 		on i.itemid=c.itemid and c.code_div='A' "
		sqlStr = sqlStr + " where c.code_large='" + cdl + "'"
		sqlStr = sqlStr + " and c.code_mid='" + cdm + "'"
		sqlStr = sqlStr + " and c.code_small='" + cds + "'" & addSql

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.close

		'지정페이지가 전체 페이지보다 크면 마지막 페이지 지정
		if Cint(FCurrPage)>Cint(FTotalPage) and Cint(FTotalPage)>0 then
			FCurrPage = FTotalPage
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_item].dbo.tbl_item i"
		sqlStr = sqlStr + " 	join [db_item].dbo.tbl_item_Category c"
		sqlStr = sqlStr + " 		on i.itemid=c.itemid and c.code_div='A' "
		sqlStr = sqlStr + " where c.code_large='" + cdl + "'"
		sqlStr = sqlStr + " and c.code_mid='" + cdm + "'"
		sqlStr = sqlStr + " and c.code_small='" + cds + "'" & addSql
		sqlStr = sqlStr + " order by i.itemid desc "


		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallImage")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FtotalPage = 1
	End Sub

	Private Sub Class_Terminate()

	End Sub

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



'/// 카테고리 선택 클래스 ///
class CategoryItem

	public FCD1
	public FCD2
	public FCD3
	public FCDName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CCategory

	public FItemList()
	public FResultCount
	public FRectCD1
	public FRectCD2
	public FRectCD3

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FResultCount      = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
	

	public Sub CategoryCodeLarge()
		dim sql, i

		sql = " select code_large, code_nm from [db_item].dbo.tbl_Cate_large "
		sql = sql + " where display_yn = 'Y'"
		sql = sql + " order by code_large Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeMid()
		dim sql, i

		sql = " select code_large, code_mid, code_nm from [db_item].dbo.tbl_Cate_mid"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " order by code_mid Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeSmall()
		dim sql, i

		sql = " select code_large, code_mid, code_small, code_nm from [db_item].dbo.tbl_Cate_small"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " order by code_small Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCD3       = rsget("code_small")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeMid2()
		dim sql, i

		sql = " select code_large, code_mid, code_nm from [db_item].dbo.tbl_Cate_mid"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " order by code_mid Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeSmall2()
		dim sql, i

		sql = " select code_large, code_mid, code_small, code_nm from [db_item].dbo.tbl_Cate_small"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " and code_small = '" + Cstr(FRectCD3) + "'"
		sql = sql & " order by code_small Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCD3       = rsget("code_small")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

end Class


'/// 카테고리 관련 키워드(링크) 클래스 ///
Class CRelateListItem
	public FcdL
	public FcdL_nm
	public FcdM
	public FcdM_nm
	public FcdS
	public FcdS_nm
	public Flinkcode
	public FlinkKeyword
	public FlinkURL

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CRelateList
	'변수 선언
	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FtotalPage

	public FRectCDL
	public FRectCDM
	public FRectCDS
	
	public FRectsearchKey
	public FRectsearchString
	public FRectLinkCode

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'관련 키워드링크 리스트 접수
	public sub GetRelateLinkList()
		dim sqlStr, addSQL, i

		'검색 추가 쿼리 작성
		if FRectCDL<>"" then	addSQL = addSQL & " and code_large='" & FRectCDL & "' "
		if FRectCDM<>"" then	addSQL = addSQL & " and code_mid='" & FRectCDM & "' "
		if FRectCDS<>"" then	addSQL = addSQL & " and code_small='" & FRectCDS & "' "
		if FRectsearchString<>"" then	addSQL = addSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "

		sqlStr = "select count(linkCode), CEILING(CAST(Count(linkCode) AS FLOAT)/" & FPageSize & ") from [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " where 1=1 " & addSQL
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		sqlStr = "Select R.* " &_
				"	,(Select top 1 code_nm from [db_item].dbo.tbl_Cate_Large with (nolock) Where code_large=R.code_large) as CDL_nm " &_
				"	,(Select top 1 code_nm from [db_item].dbo.tbl_Cate_mid with (nolock) Where code_large=R.code_large and code_mid=R.code_mid) as CDM_nm " &_
				"	,(Select top 1 code_nm from [db_item].dbo.tbl_Cate_small with (nolock) Where code_large=R.code_large and code_mid=R.code_mid and code_small=R.code_small) as CDS_nm " &_
				"From " &_
				"	(select top " + CStr(FPageSize) + " * " &_
				"	from [db_item].dbo.tbl_Cate_RelateLink with (nolock)" &_
				"	where 1=1 " & addSQL &_
				"	) as R " &_
				"order by R.linkCode desc"
		'response.Write sqlStr
		rsget.Open sqlStr, dbget, 1

			FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)

		if  Not(rsget.EOF or rsget.BOF)  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CRelateListItem

				FItemList(i).Flinkcode		= rsget("linkCode")
				FItemList(i).FcdL			= rsget("code_large")
				FItemList(i).FcdL_nm		= rsget("CDL_nm")
				FItemList(i).FcdM			= rsget("code_mid")
				FItemList(i).FcdM_nm		= rsget("CDM_nm")
				FItemList(i).FcdS			= rsget("code_small")
				FItemList(i).FcdS_nm		= rsget("CDS_nm")
				FItemList(i).FlinkKeyword	= db2html(rsget("linkKeyword"))
				FItemList(i).FlinkURL		= db2html(rsget("linkURL"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub


	'관련 키워드링크 상세정보 접수
	public sub GetRelateLinkItem()
		dim sqlStr

		sqlStr = "select * " &_
				"	from [db_item].dbo.tbl_Cate_RelateLink with (nolock)" &_
				"	where linkcode=" & FRectLinkCode
		'response.Write sqlStr
		rsget.Open sqlStr, dbget, 1

		redim preserve FItemList(1)

		if  Not(rsget.EOF or rsget.BOF)  then
			set FItemList(1) = new CRelateListItem

			FItemList(1).Flinkcode		= rsget("linkCode")
			FItemList(1).FcdL			= rsget("code_large")
			FItemList(1).FcdM			= rsget("code_mid")
			FItemList(1).FcdS			= rsget("code_small")
			FItemList(1).FlinkKeyword	= db2html(rsget("linkKeyword"))
			FItemList(1).FlinkURL		= db2html(rsget("linkURL"))
		end if
		rsget.Close
	end Sub


	'페이지 함수
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

'/// 카테고리 탑 키워드 클래스 ///
Class CCategoryKeyWordItem

	public Fidx
	Public FCDL
	Public FCDL_Nm
	Public FCDM
	Public FCDM_Nm
	public Fkeyword
	public Fitemid
	public FImageSmall
	public Flinkinfo
	public Fisusing
	public FSortNo
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CCategoryKeyWord
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectgubun
	public FRectIdx
	Public FRectCDL
	Public FRectCDM
	Public FRectUsing
	Public FRectSearch

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

	public Function GetCaFavKeyWord()
		dim sqlStr, addSQL, i

		'추가조건 쿼리
		if FRectidx <> "" then
			addSQL = addSQL + " and K.idx=" + Cstr(FRectidx) + "" + vbcrlf
		end If
		if FRectCDL <> "" then
			addSQL = addSQL + " and K.cdl='" + Cstr(FRectCDL) + "'" + vbcrlf
		end if
		if FRectCDM <> "" then
			addSQL = addSQL + " and K.cdm='" + Cstr(FRectCDM) + "'" + vbcrlf
		end if
		if (FRectUsing="Y" or FRectUsing="N") then
			addSQL = addSQL + " and K.isusing='" + Cstr(FRectUsing) + "'" + vbcrlf
		end if		
		if FRectSearch <> "" then
			addSQL = addSQL + " and K.keyword like '%" + Cstr(FRectSearch) + "%'" + vbcrlf
		end if

		'목록카운트
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") "
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_keyword as K" + vbcrlf
		sqlStr = sqlStr + " where idx <> 0" + vbcrlf + addSQL

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + "	K.idx, K.cdl, K.cdm, K.keyword, K.linkinfo, K.isusing, K.sortNo, K.regdate, i.smallimage, i.itemid " + vbcrlf
		sqlStr = sqlStr + "	,cl.code_nm as cdl_nm, cm.code_nm as cdm_nm " + vbcrlf
		sqlStr = sqlStr + "from [db_sitemaster].[dbo].tbl_category_keyword as K " + vbcrlf
		sqlStr = sqlStr + "	left join db_item.dbo.tbl_cate_large cl " + vbcrlf
		sqlStr = sqlStr + "		on cl.code_large=K.cdl " + vbcrlf
		sqlStr = sqlStr + "	left join db_item.dbo.tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + "		on cm.code_large=K.cdl and cm.code_mid=K.cdm " + vbcrlf
		sqlStr = sqlStr + "	left join db_item.dbo.tbl_item as i " + vbcrlf
		sqlStr = sqlStr + "		on K.itemid=i.itemid " + vbcrlf
		sqlStr = sqlStr + "where K.idx <> 0" + vbcrlf + addSQL
		sqlStr = sqlStr + "order by K.cdl, K.sortNo"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCategoryKeyWordItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).FCDL		= rsget("cdl")
				FItemList(i).FCDL_Nm	= rsget("cdl_nm")
				FItemList(i).FCDM		= rsget("cdm")
				FItemList(i).FCDM_Nm	= rsget("cdm_nm")
				FItemList(i).Fkeyword	= db2html(rsget("keyword"))
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).FImageSmall= "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")
				FItemList(i).Flinkinfo	= db2html(rsget("linkinfo"))
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FsortNo	= rsget("sortNo")
				FItemList(i).Fregdate	= rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetImageFolerName(byval i)
		if FItemList(i).FItemID<>"" then
			GetImageFolerName = Num2Str(Clng(FItemList(i).FItemID\10000),2,"0","R")
		end if
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


'/// MD 빅찬스/베스트 브랜드 클래스 ///
Class CSpecialItem

	public Fidx
	public Fcdl
	public Fcdm
	public Fitemid
	public Fisusing
	public Fcode_nm
	public Fcdm_nm
	public FitemName
	public FImageSmall
	public Fgubun
	public FsellYn
	public FsailYn
	public ForgPrice
	public FsailPrice

	public FmakerId
	public FImage
	public Fregdate
	public FsortNo

	public Ftitleimgurl

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

Class CMDSRecommend
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectCDL
	public FRectCDM
	public FRectStyleSerail
	public FRectGubun
	public FRectIsUsing
	public FRectIdx

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
		if FItemList(i).FItemID<>"" then
			GetImageFolerName = Num2Str(Clng(FItemList(i).FItemID\10000),2,"0","R")
		end if
	end function

	public Function GetMDSRecommendList()
		dim sqlStr,i

		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
		sqlStr = sqlStr + " from [db_item].dbo.tbl_cate_large l, [db_contents].[dbo].tbl_md_special_recommend d," + vbcrlf
		sqlStr = sqlStr + " [db_item].dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = d.cd1" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and cd1 = '" + FRectCDL + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.cd1, d.itemid, d.isusing, i.itemname,i.smallimage, l.code_nm" + vbcrlf
		sqlStr = sqlStr + " from [db_item].dbo.tbl_cate_large l, [db_contents].[dbo].tbl_md_special_recommend d," + vbcrlf
		sqlStr = sqlStr + " [db_item].dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = d.cd1" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and cd1 = '" + FRectCDL + "'" + vbcrlf
		end if

		sqlStr = sqlStr + " order by d.idx desc"
'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fcdl      = rsget("cd1")
				FItemList(i).Fitemid       = rsget("itemid")
				FItemList(i).Fisusing      = rsget("isusing")
				FItemList(i).Fcode_nm      = rsget("code_nm")
				FItemList(i).FitemName   = db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function GetMainMDChoiceList()
		dim sqlStr,i
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") from [db_contents].[dbo].tbl_main_md_choice"
		sqlStr = sqlStr + " where itemid<>0"
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectGubun + "'"
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.gubun, d.itemid, d.isusing, i.itemname,i.smallimage" + vbcrlf
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_main_md_choice d," + vbcrlf
		sqlStr = sqlStr + " [db_item].dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where d.itemid=i.itemid" + vbcrlf
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and d.gubun = '" + FRectGubun + "'" + vbcrlf
		end if

		sqlStr = sqlStr + " order by d.idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fgubun      = rsget("gubun")
				FItemList(i).Fitemid       = rsget("itemid")
				FItemList(i).Fisusing      = rsget("isusing")
				FItemList(i).FitemName   = db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function

	'// 카테고리 빅찬스 목록
	public Function GetCategoryBigChanceList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectCDL<>"" then
			addSql = addSql + " and d.cdl = '" + FRectCDL + "'" + vbcrlf
		end if
		if FRectCDM<>"" then
			addSql = addSql + " and d.cdm = '" + FRectCDM + "'" + vbcrlf
		end if

		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_left_bigchance d," + vbcrlf
		sqlStr = sqlStr + " [db_item].dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where d.itemid=i.itemid" + vbcrlf + addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.cdl, d.cdm, d.itemid, d.sortNo, i.itemname,i.smallimage, i.sellyn " + vbcrlf
		sqlStr = sqlStr + " , i.sailyn, i.orgprice, i.sailprice " + vbcrlf
		sqlStr = sqlStr + " , (select code_nm From db_item.dbo.tbl_cate_large where code_large=d.cdl) as code_nm "  + vbcrlf
		sqlStr = sqlStr + " , (select code_nm From db_item.dbo.tbl_cate_mid where code_large=d.cdl and code_mid=d.cdm) as cdm_nm "  + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_left_bigchance d," + vbcrlf
		sqlStr = sqlStr + " [db_item].dbo.tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where d.itemid=i.itemid" + vbcrlf + addSql
		if FRectCDL<>"" then
			sqlStr = sqlStr + " order by d.sortNo asc"
		else
			sqlStr = sqlStr + " order by d.idx desc"
		end if

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fcdl			= rsget("cdl")
				FItemList(i).Fcode_nm		= rsget("code_nm")
				FItemList(i).Fcdm			= rsget("cdm")
				FItemList(i).Fcdm_nm		= rsget("cdm_nm")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).FitemName		= db2html(rsget("itemname"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")
				FItemList(i).FsellYn		= rsget("sellYn")		'품절여부
				FItemList(i).FsailYn		= rsget("sailYn")		'할인여부
				FItemList(i).ForgPrice		= rsget("orgPrice")		'원판매가
				FItemList(i).FsailPrice		= rsget("sailPrice")	'할인가
				FItemList(i).FsortNo		= rsget("sortNo")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function


	'// 베스트 브랜드 목록
	public Function GetBestBrandList()
		dim sqlStr, addSQL, i

		'추가 조건 쿼리
		if FRectCDL<>"" then
			addSQL = " and b.cdl='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			addSQL = " and b.cdm='" + FRectCDM + "'"
		end if
		if FRectIdx<>"" then
			addSQL = " and b.idx=" & FRectIdx
		end if
		if FRectisusing<>"" then
			addSQL = addSQL & " and isusing='" & FRectisusing & "'"
		end if

		'목록 카운트
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") from [db_sitemaster].[dbo].tbl_category_left_bestbrand as b "
		sqlStr = sqlStr + " where idx<>0" & addSQL

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		'목록 내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.cdm, b.makerid, b.imgfile, b.regdate, l.code_nm, b.isusing, b.sortNo " + vbcrlf
		sqlStr = sqlStr + "	,(Select code_nm from [db_item].[dbo].tbl_cate_mid Where code_large = b.cdl and code_mid = b.cdm) as midcode_nm " + vbCrLf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_cate_large l, [db_sitemaster].[dbo].tbl_category_left_bestbrand b" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = b.cdl" & addSQL
		sqlStr = sqlStr + " order by b.cdl, b.sortNo "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).FIdx		= rsget("idx")
				FItemList(i).FCdL		= rsget("cdl")
				FItemList(i).FCdM		= rsget("cdm")
				FItemList(i).FmakerId	= rsget("makerid")
				FitemList(i).FImage		= rsget("imgfile")
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fcdm_nm	= rsget("midcode_nm")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).FsortNo	= rsget("sortNo")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function


	'// 브랜드 포커스 목록
	public Function GetBrandFocusList()
		dim sqlStr, addSQL, i

		'추가 조건 쿼리
		if FRectCDL<>"" then
			addSQL = " and b.cdl='" + FRectCDL + "'"
		end if

		'목록 카운트
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") from [db_sitemaster].[dbo].tbl_category_left_brand_rank as b "
		sqlStr = sqlStr + " where idx<>0" & addSQL

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		'목록 내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.makerid, b.sortNo " + vbcrlf
		sqlStr = sqlStr + " , c.titleimgurl, c.modelitem, c.modelimg" + vbcrlf
		sqlStr = sqlStr + "	,(Select code_nm from [db_item].[dbo].tbl_cate_large Where code_large = b.cdl) as code_nm " + vbCrLf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_left_brand_rank b" + vbcrlf
		sqlStr = sqlStr + "		,[db_user].[dbo].tbl_user_c c" + vbcrlf
		sqlStr = sqlStr + " where b.makerid = c.userid " & addSQL
		sqlStr = sqlStr + " order by b.cdl, b.sortNo "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).FIdx			= rsget("idx")
				FItemList(i).FCdL			= rsget("cdl")
				FItemList(i).FmakerId		= db2html(rsget("makerid"))
				FItemList(i).Fcode_nm		= rsget("code_nm")
				FItemList(i).FsortNo		= rsget("sortNo")
				FItemList(i).Ftitleimgurl	= "http://webimage.10x10.co.kr/image/brandlogo/" + rsget("titleimgurl")
				FItemList(i).FitemId		=  rsget("modelitem")
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("modelimg")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function


	'// 페이지 관련 함수
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
%>
