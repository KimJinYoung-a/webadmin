<%
'###############################################
' PageName : DIYCategoryCls.asp
' Discription : 카테고리 관련 클래스
' History : 2010.09.16 허진원 생성
'			2010.11.10 한용민 수정
'###############################################

Class CCatemanageItem
	public Fcdlarge
	public Fcdmid
	public Fcdsmall
	public Fchannel
	public Fnmlarge
	public Fcatecnt
	public ForderNo
	public FIdx
	public FCdL
	public FMakerid
	public FImage
	public Fcode_nm
	public Fisusing
	public Fregdate
	public FsortNo
	
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
	public FRectCDL
	public FRectisUsing
	public FRectIdx
		
	'/academy/itemmaster/bestbrand/category_left_bestbrand.asp
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
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ")" 
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_category_left_bestbrand as b"
		sqlStr = sqlStr + " where idx<>0 " & addSQL

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget(0)
			FtotalPage	= rsACADEMYget(1)
		rsACADEMYget.Close

		'목록 내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.makerid, b.imgfile, b.regdate, l.code_nm, b.isusing, b.sortNo " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_Cate_large l"
		sqlStr = sqlStr + " join [db_academy].[dbo].tbl_category_left_bestbrand b" + vbcrlf
		sqlStr = sqlStr + " on l.code_large = b.cdl" + vbcrlf
		sqlStr = sqlStr + " where idx<>0 " & addSQL
		sqlStr = sqlStr + " order by b.cdl, b.sortNo "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

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

	public sub GetNewCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_academy].dbo.tbl_diy_item_Cate_large"
		sqlStr = sqlStr + " order by code_large"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge	= rsACADEMYget("code_large")
				FItemList(i).Fnmlarge	= db2html(rsACADEMYget("code_nm"))

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close
	end sub

	public sub GetNewCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm,orderNo from [db_academy].dbo.tbl_diy_item_Cate_mid"
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
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_Cate_small s"
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
End Class

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
End Class
%>