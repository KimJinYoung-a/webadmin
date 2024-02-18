<%

Class CateHotKeywordItem

	public FIdx
	public Fcdl
	public Fitemid
	public Fitemname
	public FimgFile
	public FIsusing
	public Fcode_nm
	public Fregdate
	public FdivCd
	public FdivName
	public FlinkURL
	public FSortNo
	public FKeyWord

	public FdivType
	public FimgWidth
	public FimgHeight

	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

end Class

Class CateHotKeyword
	public FIdx
	public Fcdl
	public Fitemid
	public Fitemname
	public FIsusing
	public FdivCd
	public FdivName

	public FdivType
	public FimgWidth
	public FimgHeight
	public FlinkURL
	public FimgFile
	public FDisp
	public FSortNo
	public FKeyWord

	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectcdl
	public FRectcdm
	public FRectisusing
	public FRectitemid
	public FRectdivCd

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public Function GetPageItemList()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from [db_sitemaster].[dbo].tbl_category_hotkeyword"
		sqlStr = sqlStr + " where idx<>0"
		sqlStr = sqlStr + " and isusing='" & FRectisusing & "'"
		if FDisp<>"" then
			sqlStr = sqlStr + " and disp='" + FDisp + "'"
		end if
		

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr =	"Select top " & CStr(FPageSize*FCurrPage) &_
					"	b.idx, b.disp, b.itemid, b.regdate, b.isusing, b.link,b.keyword, " &_
					"	isNull((db_item.dbo.getDisplayCateName(b.disp)),'') as code_nm, c.itemname, c.icon2image, " &_
					"	b.sortno " &_
					"	,c.sellyn, c.limityn, c.limitno, c.limitsold " &_
					"From [db_sitemaster].[dbo].tbl_category_hotkeyword b " &_
					"	Left Join [db_item].[dbo].tbl_item c " &_
					"		on b.itemid=c.itemid " &_
					"Where b.isusing='" & FRectisusing & "'"
		if FDisp<>"" then
			sqlStr = sqlStr + " and b.disp='" + FDisp + "'"
		end if

		sqlStr = sqlStr + " order by b.sortno asc, b.idx desc"

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
				set FItemList(i) = new CateHotKeywordItem

				FItemList(i).FIdx		= rsget("idx")
				FItemList(i).FitemId	= rsget("itemid")
				FItemList(i).Fitemname	= rsget("itemname")
				if Not(rsget("itemid")="" or isNull(rsget("itemid"))) Then
					FitemList(i).FimgFile	= "http://webimage.10x10.co.kr/image/icon2/" & GetImageSubFolderByItemid(FItemList(i).FitemId) & "/" & rsget("icon2image")
				end if
'				if Not(rsget("img1")="" or isNull(rsget("img1"))) then
'					FitemList(i).FimgFile	= "http://imgstatic.10x10.co.kr/main/categoryPage/" & rsget("img1")
'				end if
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FlinkURL	= rsget("link")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).FSortNo	= rsget("sortno")
				FItemList(i).FKeyWord	= rsget("keyword")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function
	
	public Sub GetOnePageItem(byval idx)
	dim sql
		sql="Select b.*, c.itemname " &_
			"From [db_sitemaster].[dbo].tbl_category_hotkeyword b " &_
			"	Left Join [db_item].[dbo].tbl_item c " &_
			"		on b.itemid=c.itemid " &_
			" where b.idx='" + CStr(idx) + "'" 

		rsget.open sql,dbget,1

		FitemId		= rsget("itemid")
		Fitemname	= rsget("itemname")
		FIsusing	= rsget("isusing")
		FlinkURL	= rsget("link")
'		FimgFile	= rsget("imgFile")
		FDisp		= rsget("disp")
		FSortNo		= rsget("sortno")
		FKeyWord	= rsget("keyword")

'		FimgWidth	= rsget("imgWidth")
'		FimgHeight	= rsget("imgHeight")
		
		rsget.close
	end Sub
		

	'// 항목구분 코드 목록 접수
	public Sub GetPageDivList()
		dim sqlStr,i
		sqlStr = "Select count(divCd), CEILING(CAST(Count(divCd) AS FLOAT)/" & Fpagesize & ") " &_
				"From [db_sitemaster].[dbo].tbl_category_mainItem_div "

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FTotalPage = rsget(1)
		rsget.Close

		sqlStr =	"Select top " & CStr(FPageSize*FCurrPage) & " * " &_
					"From [db_sitemaster].[dbo].tbl_category_mainItem_div " &_
					"order by divCd asc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CateMainPageItem

				FItemList(i).FdivCd		= rsget("divCd")
				FItemList(i).FdivName	= rsget("divName")
				FItemList(i).FimgWidth	= rsget("imgWidth")
				FItemList(i).FimgHeight	= rsget("imgHeight")
				FItemList(i).FIsusing	= rsget("isUsing")
				FItemList(i).Fregdate	= rsget("regdate")

				'형식 지정
				Select Case rsget("divType")
					Case "I"
						FItemList(i).FdivType = "상품지정"
					Case "M"
						FItemList(i).FdivType = "이미지 선택"
					Case "B"
						FItemList(i).FdivType = "상품지정 및 이미지추가"
				end Select

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub


	public Sub GetOnePageDivCd()
		dim sqlStr

		sqlStr =	"Select * " &_
					"From [db_sitemaster].[dbo].tbl_category_mainItem_div " &_
					"Where divCd=" & FRectdivCd

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		if  not rsget.EOF  then
			do until rsget.eof

				FdivCd		= rsget("divCd")
				FdivName	= rsget("divName")
				FimgWidth	= rsget("imgWidth")
				FimgHeight	= rsget("imgHeight")
				FdivType	= rsget("divType")
				FIsusing	= rsget("isUsing")
				
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub


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
