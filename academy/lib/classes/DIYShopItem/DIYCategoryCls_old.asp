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

	public Sub GetCategoryKeyword(cdl,cdm,cds)
		dim sqlStr,i
		sqlStr = "select top 1 keyword from [db_academy].dbo.tbl_diy_item_cate_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " and code_small='" + cds + "'"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		redim preserve FItemList(0)

		'값이 있던 없던 클랙스 배열 선언
		set FItemList(0) = new COptionManagerItem

		if not rsACADEMYget.Eof then
			FItemList(0).Fkeyword = rsACADEMYget("keyword")
		else
			FItemList(0).Fkeyword = ""
		end if

		rsACADEMYget.close
	end sub

	public sub GetOrgCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_academy].dbo.tbl_diy_item_cate_large"
		sqlStr = sqlStr + " order by code_large"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsACADEMYget("code_large")
				FItemList(i).Fnmlarge        = db2html(rsACADEMYget("code_nm"))

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	public sub GetOrgCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm from [db_academy].dbo.tbl_diy_item_cate_mid"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " order by code_mid"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsACADEMYget("code_large")
				FItemList(i).Fcdmid          = rsACADEMYget("code_mid")
				FItemList(i).Fnmlarge        = db2html(rsACADEMYget("code_nm"))

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	public sub GetOrgCateMasterSmall(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select s.code_large, s.code_mid, s.code_small, s.code_nm, IsNULL(T.cnt,0) as catecnt"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_small s"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select cate_small, count(itemid) as cnt from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " 	where cate_large='" + cdl + "'"
		sqlStr = sqlStr + " 	and cate_mid='" + cdm + "'"
		sqlStr = sqlStr + "		group by cate_small"
		sqlStr = sqlStr + "	) as T on s.code_small=T.cate_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " order by code_small"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsACADEMYget("code_large")
				FItemList(i).Fcdmid          = rsACADEMYget("code_mid")
				FItemList(i).Fcdsmall          = rsACADEMYget("code_small")
				FItemList(i).Fnmlarge        = db2html(rsACADEMYget("code_nm"))
				FItemList(i).Fcatecnt        = rsACADEMYget("catecnt")
				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	public sub GetOrgCateItemList(cdl,cdm,cds)
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " where i.cate_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.cate_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.cate_small='" + cds + "'"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsACADEMYget("itemid")
				FItemList(i).FItemName  = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fmakerid   = rsACADEMYget("makerid")

				FItemList(i).FSellyn    = rsACADEMYget("sellyn")
				FItemList(i).Fisusing   = rsACADEMYget("isusing")

				FItemList(i).FImgSmall  = imgFingers & "/diyItem/waitimage/small/" & GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsACADEMYget("smallImage")

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.Close
	end sub

	public sub GetOrgCateNotMachItemList()
		dim sqlStr,i
		sqlStr = "select count(itemid), CEILING(CAST(Count(itemid) AS FLOAT)/" & FPageSize & ") from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget(0)
			FTotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + "  order by i.itemid desc "

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsACADEMYget("itemid")
				FItemList(i).FItemName  = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fmakerid   = rsACADEMYget("makerid")

				FItemList(i).FSellyn    = rsACADEMYget("sellyn")
				FItemList(i).Fisusing   = rsACADEMYget("isusing")

				FItemList(i).FImgSmall  = imgFingers & "/diyItem/waitimage/small/" & GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsACADEMYget("smallImage")

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.Close
	end sub

	public sub GetNewCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_academy].dbo.tbl_diy_item_cate_large"
		sqlStr = sqlStr + " order by code_large"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsACADEMYget("code_large")
				FItemList(i).Fnmlarge        = db2html(rsACADEMYget("code_nm"))

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	public sub GetNewCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm,orderNo from [db_academy].dbo.tbl_diy_item_cate_mid"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " order by orderNo ,code_mid"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsACADEMYget("code_large")
				FItemList(i).Fcdmid          = rsACADEMYget("code_mid")
				FItemList(i).Fnmlarge        = db2html(rsACADEMYget("code_nm"))
				FItemList(i).FOrderNo				=rsACADEMYget("orderNo")

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub


	public sub GetNewCateMasterSmall(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select s.code_large, s.code_mid, s.code_small, s.code_nm, orderNo ,IsNULL(T.cnt,0) as catecnt"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_small s"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select cate_small, count(itemid) as cnt from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " 	where cate_large='" + cdl + "'"
		sqlStr = sqlStr + " 	and cate_mid='" + cdm + "'"
		sqlStr = sqlStr + "		group by cate_small"
		sqlStr = sqlStr + "	) as T on s.code_small=T.cate_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " order by orderNo, code_small"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsACADEMYget("code_large")
				FItemList(i).Fcdmid          = rsACADEMYget("code_mid")
				FItemList(i).Fcdsmall          = rsACADEMYget("code_small")
				FItemList(i).Fnmlarge        = db2html(rsACADEMYget("code_nm"))
				FItemList(i).Fcatecnt        = rsACADEMYget("catecnt")
				FItemList(i).FOrderNo        = rsACADEMYget("orderNo")
				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	public function GetNewCateCurrentPos(cdl,cdm,cds)
		dim sqlStr
		sqlStr = "select distinct top 1 code_nm "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_large"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		if not rsACADEMYget.Eof then
			GetNewCateCurrentPos = db2html(rsACADEMYget("code_nm"))
		end if
		rsACADEMYget.close


		if cdm<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_mid"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			if not rsACADEMYget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" +  db2html(rsACADEMYget("code_nm"))
			end if
			rsACADEMYget.close
		end if

		if cds<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_small"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			sqlStr = sqlStr + " and code_small='" + cds + "'"
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			if not rsACADEMYget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" + db2html(rsACADEMYget("code_nm"))
			end if
			rsACADEMYget.close
		end if

	end function

	public sub GetNewCateItemList(cdl,cdm,cds)
		dim sqlStr,i

		sqlStr = "select count(i.itemid), CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " where i.cate_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.cate_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.cate_small='" + cds + "'"
		if FRectDispSailYN = "on" then
		sqlStr = sqlStr + " and i.sellyn='Y'"
		end if

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.isusing, i.smallImage "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " where i.cate_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.cate_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.cate_small='" + cds + "'"
		if FRectDispSailYN = "on" then
		sqlStr = sqlStr + " and i.sellyn='Y'"
		end if


		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsACADEMYget("itemid")
				FItemList(i).FItemName  = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fmakerid   = rsACADEMYget("makerid")

				FItemList(i).FSellyn    = rsACADEMYget("sellyn")
				FItemList(i).Fisusing   = rsACADEMYget("isusing")

				FItemList(i).FImgSmall  = imgFingers & "/diyItem/waitimage/small/" & GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsACADEMYget("smallImage")

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.Close
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

		sql = " select code_large, code_nm from [db_academy].dbo.tbl_diy_item_cate_large "
		sql = sql + " where display_yn = 'Y'"
		sql = sql + " order by code_large Asc"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
	        i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsACADEMYget("code_large")
				FItemList(i).FCDName      = rsACADEMYget("code_nm")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end sub

	public Sub CategoryCodeMid()
		dim sql, i

		sql = " select code_large, code_mid, code_nm from [db_academy].dbo.tbl_diy_item_cate_mid"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " order by code_mid Asc"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
	        i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsACADEMYget("code_large")
				FItemList(i).FCD2       = rsACADEMYget("code_mid")
				FItemList(i).FCDName      = rsACADEMYget("code_nm")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end sub

	public Sub CategoryCodeSmall()
		dim sql, i

		sql = " select code_large, code_mid, code_small, code_nm from [db_academy].dbo.tbl_diy_item_cate_small"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " order by code_small Asc"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
	        i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsACADEMYget("code_large")
				FItemList(i).FCD2       = rsACADEMYget("code_mid")
				FItemList(i).FCD3       = rsACADEMYget("code_small")
				FItemList(i).FCDName      = rsACADEMYget("code_nm")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end sub

	public Sub CategoryCodeMid2()
		dim sql, i

		sql = " select code_large, code_mid, code_nm from [db_academy].dbo.tbl_diy_item_cate_mid"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " order by code_mid Asc"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
	        i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsACADEMYget("code_large")
				FItemList(i).FCD2       = rsACADEMYget("code_mid")
				FItemList(i).FCDName      = rsACADEMYget("code_nm")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end sub

	public Sub CategoryCodeSmall2()
		dim sql, i

		sql = " select code_large, code_mid, code_small, code_nm from [db_academy].dbo.tbl_diy_item_cate_small"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " and code_small = '" + Cstr(FRectCD3) + "'"
		sql = sql & " order by code_small Asc"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
	        i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsACADEMYget("code_large")
				FItemList(i).FCD2       = rsACADEMYget("code_mid")
				FItemList(i).FCD3       = rsACADEMYget("code_small")
				FItemList(i).FCDName      = rsACADEMYget("code_nm")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
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

		sqlStr = "select count(linkCode), CEILING(CAST(Count(linkCode) AS FLOAT)/" & FPageSize & ") from [db_academy].dbo.tbl_diy_item_cate_RelateLink"
		sqlStr = sqlStr + " where 1=1 " & addSQL
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget(0)
			FtotalPage = rsACADEMYget(1)
		rsACADEMYget.Close

		sqlStr = "Select R.* " &_
				"	,(Select top 1 code_nm from [db_academy].dbo.tbl_diy_item_cate_Large Where code_large=R.code_large) as CDL_nm " &_
				"	,(Select top 1 code_nm from [db_academy].dbo.tbl_diy_item_cate_mid Where code_large=R.code_large and code_mid=R.code_mid) as CDM_nm " &_
				"	,(Select top 1 code_nm from [db_academy].dbo.tbl_diy_item_cate_small Where code_large=R.code_large and code_mid=R.code_mid and code_small=R.code_small) as CDS_nm " &_
				"From " &_
				"	(select top " + CStr(FPageSize) + " * " &_
				"	from [db_academy].dbo.tbl_diy_item_cate_RelateLink " &_
				"	where 1=1 " & addSQL &_
				"	) as R " &_
				"order by R.linkCode desc"
		'response.Write sqlStr
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

			FResultCount = rsACADEMYget.RecordCount
			redim preserve FItemList(FResultCount)

		if  Not(rsACADEMYget.EOF or rsACADEMYget.BOF)  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CRelateListItem

				FItemList(i).Flinkcode		= rsACADEMYget("linkCode")
				FItemList(i).FcdL			= rsACADEMYget("code_large")
				FItemList(i).FcdL_nm		= rsACADEMYget("CDL_nm")
				FItemList(i).FcdM			= rsACADEMYget("code_mid")
				FItemList(i).FcdM_nm		= rsACADEMYget("CDM_nm")
				FItemList(i).FcdS			= rsACADEMYget("code_small")
				FItemList(i).FcdS_nm		= rsACADEMYget("CDS_nm")
				FItemList(i).FlinkKeyword	= db2html(rsACADEMYget("linkKeyword"))
				FItemList(i).FlinkURL		= db2html(rsACADEMYget("linkURL"))

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.Close
	end sub


	'관련 키워드링크 상세정보 접수
	public sub GetRelateLinkItem()
		dim sqlStr

		sqlStr = "select * " &_
				"	from [db_academy].dbo.tbl_diy_item_cate_RelateLink " &_
				"	where linkcode=" & FRectLinkCode
		'response.Write sqlStr
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		redim preserve FItemList(1)

		if  Not(rsACADEMYget.EOF or rsACADEMYget.BOF)  then
			set FItemList(1) = new CRelateListItem

			FItemList(1).Flinkcode		= rsACADEMYget("linkCode")
			FItemList(1).FcdL			= rsACADEMYget("code_large")
			FItemList(1).FcdM			= rsACADEMYget("code_mid")
			FItemList(1).FcdS			= rsACADEMYget("code_small")
			FItemList(1).FlinkKeyword	= db2html(rsACADEMYget("linkKeyword"))
			FItemList(1).FlinkURL		= db2html(rsACADEMYget("linkURL"))
		end if
		rsACADEMYget.Close
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

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

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCategoryKeyWordItem

				FItemList(i).Fidx		= rsACADEMYget("idx")
				FItemList(i).FCDL		= rsACADEMYget("cdl")
				FItemList(i).FCDL_Nm	= rsACADEMYget("cdl_nm")
				FItemList(i).FCDM		= rsACADEMYget("cdm")
				FItemList(i).FCDM_Nm	= rsACADEMYget("cdm_nm")
				FItemList(i).Fkeyword	= db2html(rsACADEMYget("keyword"))
				FItemList(i).Fitemid	= rsACADEMYget("itemid")
				FItemList(i).FImageSmall= "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsACADEMYget("smallimage")
				FItemList(i).Flinkinfo	= db2html(rsACADEMYget("linkinfo"))
				FItemList(i).Fisusing	= rsACADEMYget("isusing")
				FItemList(i).FsortNo	= rsACADEMYget("sortNo")
				FItemList(i).Fregdate	= rsACADEMYget("regdate")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_large l, [db_contents].[dbo].tbl_md_special_recommend d," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item i" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = d.cd1" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and cd1 = '" + FRectCDL + "'" + vbcrlf
		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.cd1, d.itemid, d.isusing, i.itemname,i.smallimage, l.code_nm" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_cate_large l, [db_contents].[dbo].tbl_md_special_recommend d," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item i" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = d.cd1" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and cd1 = '" + FRectCDL + "'" + vbcrlf
		end if

		sqlStr = sqlStr + " order by d.idx desc"
'response.write sqlStr
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx     = rsACADEMYget("idx")
				FItemList(i).Fcdl      = rsACADEMYget("cd1")
				FItemList(i).Fitemid       = rsACADEMYget("itemid")
				FItemList(i).Fisusing      = rsACADEMYget("isusing")
				FItemList(i).Fcode_nm      = rsACADEMYget("code_nm")
				FItemList(i).FitemName   = db2html(rsACADEMYget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsACADEMYget("smallimage")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end function

	public Function GetMainMDChoiceList()
		dim sqlStr,i
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") from [db_contents].[dbo].tbl_main_md_choice"
		sqlStr = sqlStr + " where itemid<>0"
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectGubun + "'"
		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.gubun, d.itemid, d.isusing, i.itemname,i.smallimage" + vbcrlf
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_main_md_choice d," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item i" + vbcrlf
		sqlStr = sqlStr + " where d.itemid=i.itemid" + vbcrlf
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and d.gubun = '" + FRectGubun + "'" + vbcrlf
		end if

		sqlStr = sqlStr + " order by d.idx desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx     = rsACADEMYget("idx")
				FItemList(i).Fgubun      = rsACADEMYget("gubun")
				FItemList(i).Fitemid       = rsACADEMYget("itemid")
				FItemList(i).Fisusing      = rsACADEMYget("isusing")
				FItemList(i).FitemName   = db2html(rsACADEMYget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsACADEMYget("smallimage")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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
		sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item i" + vbcrlf
		sqlStr = sqlStr + " where d.itemid=i.itemid" + vbcrlf + addSql

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " d.idx, d.cdl, d.cdm, d.itemid, d.sortNo, i.itemname,i.smallimage, i.sellyn " + vbcrlf
		sqlStr = sqlStr + " , i.sailyn, i.orgprice, i.sailprice " + vbcrlf
		sqlStr = sqlStr + " , (select code_nm From db_item.dbo.tbl_cate_large where code_large=d.cdl) as code_nm "  + vbcrlf
		sqlStr = sqlStr + " , (select code_nm From db_item.dbo.tbl_cate_mid where code_large=d.cdl and code_mid=d.cdm) as cdm_nm "  + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_left_bigchance d," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item i" + vbcrlf
		sqlStr = sqlStr + " where d.itemid=i.itemid" + vbcrlf + addSql
		if FRectCDL<>"" then
			sqlStr = sqlStr + " order by d.sortNo asc"
		else
			sqlStr = sqlStr + " order by d.idx desc"
		end if

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx			= rsACADEMYget("idx")
				FItemList(i).Fcdl			= rsACADEMYget("cdl")
				FItemList(i).Fcode_nm		= rsACADEMYget("code_nm")
				FItemList(i).Fcdm			= rsACADEMYget("cdm")
				FItemList(i).Fcdm_nm		= rsACADEMYget("cdm_nm")
				FItemList(i).Fitemid		= rsACADEMYget("itemid")
				FItemList(i).FitemName		= db2html(rsACADEMYget("itemname"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsACADEMYget("smallimage")
				FItemList(i).FsellYn		= rsACADEMYget("sellYn")		'품절여부
				FItemList(i).FsailYn		= rsACADEMYget("sailYn")		'할인여부
				FItemList(i).ForgPrice		= rsACADEMYget("orgPrice")		'원판매가
				FItemList(i).FsailPrice		= rsACADEMYget("sailPrice")	'할인가
				FItemList(i).FsortNo		= rsACADEMYget("sortNo")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end function


	'// 베스트 브랜드 목록
	public Function GetBestBrandList()
		dim sqlStr, addSQL, i

		'추가 조건 쿼리
		if FRectCDL<>"" then
			addSQL = " and b.cdl='" + FRectCDL + "'"
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		'목록 내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.makerid, b.imgfile, b.regdate, l.code_nm, b.isusing, b.sortNo " + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_cate_large l, [db_sitemaster].[dbo].tbl_category_left_bestbrand b" + vbcrlf
		sqlStr = sqlStr + " where l.code_large = b.cdl" & addSQL
		sqlStr = sqlStr + " order by b.cdl, b.sortNo "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).FIdx		= rsACADEMYget("idx")
				FItemList(i).FCdL		= rsACADEMYget("cdl")
				FItemList(i).FmakerId	= rsACADEMYget("makerid")
				FitemList(i).FImage		= rsACADEMYget("imgfile")
				FItemList(i).Fcode_nm	= rsACADEMYget("code_nm")
				FItemList(i).Fisusing	= rsACADEMYget("isusing")
				FItemList(i).Fregdate	= rsACADEMYget("regdate")
				FItemList(i).FsortNo	= rsACADEMYget("sortNo")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		'목록 내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.makerid, b.sortNo " + vbcrlf
		sqlStr = sqlStr + " , c.titleimgurl, c.modelitem, c.modelimg" + vbcrlf
		sqlStr = sqlStr + "	,(Select code_nm from [db_item].[dbo].tbl_cate_large Where code_large = b.cdl) as code_nm " + vbCrLf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_left_brand_rank b" + vbcrlf
		sqlStr = sqlStr + "		,[db_user].[dbo].tbl_user_c c" + vbcrlf
		sqlStr = sqlStr + " where b.makerid = c.userid " & addSQL
		sqlStr = sqlStr + " order by b.cdl, b.sortNo "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).FIdx			= rsACADEMYget("idx")
				FItemList(i).FCdL			= rsACADEMYget("cdl")
				FItemList(i).FmakerId		= db2html(rsACADEMYget("makerid"))
				FItemList(i).Fcode_nm		= rsACADEMYget("code_nm")
				FItemList(i).FsortNo		= rsACADEMYget("sortNo")
				FItemList(i).Ftitleimgurl	= "http://webimage.10x10.co.kr/image/brandlogo/" + rsACADEMYget("titleimgurl")
				FItemList(i).FitemId		=  rsACADEMYget("modelitem")
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsACADEMYget("modelimg")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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
