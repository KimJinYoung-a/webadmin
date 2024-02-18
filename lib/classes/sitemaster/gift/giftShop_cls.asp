<%
'###########################################################
' Description :  기프트 Shop 클래스
' History : 2014.04.01 허진원 생성
'###########################################################

'===============================================
'// 클래스 아이템 선언
'===============================================

Class CGiftShopItem
    public FthemeIdx
    public Fsubject
    public FsubDesc
    public FisOpen
    public Fuserid
	public FfrontItemid
    public FviewCount
    public FcommentCount
    public FitemCount
    public Fdevice
    public Fregdate
    public FsortNo
    public FisPick
    public FpickImage
    public FisUsing
    public Fadminid
    public Fadminname

    public FitemName
    public FbasicImage
    public Ftag

	'// 프론트 공개 여부
	Function IsOpend()
		if FisOpen="N" or FisUsing="N" then
			IsOpend = false
		else
			IsOpend = true
		end if
	End Function

	'// 관리구분
	Function getPickType()
		if FisPick="Y" then
			getPickType = "<span style=""color:darkred;font-weight:bold;"">10x10's Pick</span>"
		else
			getPickType = "User's Pick"
		end if
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CGiftShopPrdItem
    public Fitemid
    public Fmakerid
    public Fbrandname
    public Fitemname
    public FOrgprice
    public FSellCash
    public FLimitNo
    public FLimitSold
    public FisUsing
    public FSellYn
    public FSaleYn
    public FLimitYn
    public FItemCouponYN
    public FItemCouponType
    public FItemCouponValue
    public FRegdate
    public FBasicImage
    public FSmallImage

	Function isSoldOut()
		if FSellYn="N" or (FLimitYn="Y" and (FLimitNo-FLimitSold)<=0 and FisUsing="N") then
			isSoldOut = "<b>품절</b>"
		else
			isSoldOut = "판매중"
		end if
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


'===============================================
'// Gift Shop 클래스
'===============================================
Class CGiftShop
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx				'테마 번호
    public FRectIsOpen			'공개여부
    public FRectIsUsing			'사용여부
    public FRectIsPick			'관리글여부
    public FRectKeywordIdx		'키워드 (array)
    public FRectItemid			'상품코드
    public FRectSortMtd			'정렬방법
    public FRectIsMyItem		'내 테마 여부
    public FRectIsSoldOut		'품절 여부
    public FRectPackIdx			'포장상품 구분

    '# 페이지정보 목록
	public Sub GetThemeList()
		dim sqlStr, addSql, orderSql, i

		'사용여부
		if FRectIsUsing<>"" then
			addSql = " Where m.isUsing='" & FRectIsUsing & "'"
		else
			addSql = " Where m.isUsing='Y'"
		end if

		'공개여부
		if FRectIsOpen<>"A" then addSql = addSql & " and m.isOpen='" & FRectIsOpen & "'"

		'관리테마 여부
		if FRectIsPick<>"A" then addSql = addSql & " and m.isPick='" & FRectIsPick & "'"

		orderSql = " order by m.sortNo asc, m.themeIdx desc"


        '전체 카운트
        sqlStr = "select count(themeIdx), CEILING(CAST(Count(themeIdx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_theme as m " & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " m.*, isNull(u.username,m.userid) as adminname "
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_theme as m "
        sqlStr = sqlStr & "	left join db_partner.dbo.tbl_user_tenbyten as u "
		sqlStr = sqlStr & "		on m.adminid=u.userid " & addSql & orderSql
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CGiftShopItem

	            FItemList(i).FthemeIdx		= rsget("themeIdx")
	            FItemList(i).Fsubject		= rsget("subject")
	            FItemList(i).FsubDesc		= rsget("subDesc")
	            FItemList(i).FisOpen		= rsget("isOpen")
	            FItemList(i).Fuserid		= rsget("userid")
	            FItemList(i).FfrontItemid	= rsget("frontItemid")
	            FItemList(i).FviewCount		= rsget("viewCount")
	            FItemList(i).FcommentCount	= rsget("commentCount")
	            FItemList(i).FitemCount		= rsget("itemCount")
	            FItemList(i).Fdevice		= rsget("device")
	            FItemList(i).Fregdate		= rsget("regdate")
	            FItemList(i).FsortNo		= rsget("sortNo")
	            FItemList(i).FisPick		= rsget("isPick")
	            FItemList(i).FisUsing		= rsget("isUsing")
	            FItemList(i).Fadminid		= rsget("adminid")
	            FItemList(i).Fadminname		= rsget("adminname")

				if Not(rsget("pickImage")="" or isNull(rsget("pickImage"))) then
					FItemList(i).FpickImage = staticImgUrl & "/gift/shop/" + rsget("pickImage")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

    '# 테마 정보 접수
	public Sub GetThemeInfo()
		dim sqlStr, i

		'내용 접수
        sqlStr = "Select top 1 m.* "
        sqlStr = sqlStr & ",STUFF(( "
        sqlStr = sqlStr & "	SELECT ',' + convert(varchar(10),c.keywordidx) "
        sqlStr = sqlStr & "	FROM db_board.dbo.tbl_gift_keyword as c "
        sqlStr = sqlStr & "	JOIN db_board.dbo.tbl_giftShop_theme_keyword AS k "
        sqlStr = sqlStr & "		ON c.keywordidx = k.keywordidx "
        sqlStr = sqlStr & "		and c.isusing='Y' "
        sqlStr = sqlStr & "		and c.keywordtype=1 "
        sqlStr = sqlStr & "	WHERE k.themeIdx = m.themeIdx "
        sqlStr = sqlStr & "	order by c.sortno asc "
        sqlStr = sqlStr & "	FOR XML PATH('') "
        sqlStr = sqlStr & "), 1, 1, '') AS tag "
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_theme as m "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & "		on m.frontItemid=i.itemid "
        sqlStr = sqlStr & "Where m.isUsing='Y' "
        sqlStr = sqlStr & "	and themeIdx=" & FRectIdx
		rsget.Open sqlStr, dbget, 1

		if Not(rsget.EOF or rsget.BOF) then
			FResultCount = 1
			set FOneItem = new CGiftShopItem

            FOneItem.FthemeIdx		= rsget("themeIdx")
            FOneItem.Fsubject		= rsget("subject")
            FOneItem.FsubDesc		= rsget("subDesc")
            FOneItem.FisOpen		= rsget("isOpen")
            FOneItem.Fuserid		= rsget("userid")
            FOneItem.FfrontItemid	= rsget("frontItemid")
            FOneItem.FviewCount		= rsget("viewCount")
            FOneItem.FcommentCount	= rsget("commentCount")
            FOneItem.FitemCount		= rsget("itemCount")
            FOneItem.Fdevice		= rsget("device")
            FOneItem.Fregdate		= rsget("regdate")
            FOneItem.FsortNo		= rsget("sortNo")
            FOneItem.FisPick		= rsget("isPick")
            FOneItem.FisUsing		= rsget("isUsing")
            FOneItem.Ftag			= rsget("tag")

			if Not(rsget("pickImage")="" or isNull(rsget("pickImage"))) then
				FOneItem.FpickImage = rsget("pickImage")
			end if

		else
			FResultCount = 0
		end if
		rsget.close

	end Sub

    '# 테마 상품 목록
	public Sub GetThemeItemList()
		dim sqlStr, addSql, orderSql, i

		'품절 포함 여부
		if FRectIsSoldOut<>"Y" then
			addSql = addSql & " and i.sellyn in ('Y','S')"
		end if

		Select Case FRectSortMtd
			Case "ne"		'신상순
				orderSql = " order by i.itemid desc"
			Case "be"		'인기순
				orderSql = " order by i.itemscore desc, i.itemid desc"
			Case "lp"		'낮은 가격
				orderSql = " order by i.sellcash asc, i.itemid desc"
			Case "hp"		'높은 가격
				orderSql = " order by i.sellcash desc, i.itemid desc"
			Case "hs"		'높은 할인율
				orderSql = " order by ((i.orgprice-i.sellcash)/i.orgprice) desc, (i.orgprice-i.sellcash) desc, i.itemid desc"
			Case Else
				orderSql = " order by i.itemid desc"
		End Select

        '전체 카운트
        sqlStr = "select count(d.itemid), CEILING(CAST(Count(d.itemid) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "from db_board.dbo.tbl_giftShop_theme as m "
        sqlStr = sqlStr & "	join db_board.dbo.tbl_giftShop_theme_item as d "
        sqlStr = sqlStr & "		on m.themeIdx=d.themeIdx "
        sqlStr = sqlStr & "	join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on d.itemid=i.itemid "
        sqlStr = sqlStr & "Where m.themeIdx=" & FRectIdx & " " & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " i.*, d.regdate as tmItmRegDt "
        sqlStr = sqlStr & "from db_board.dbo.tbl_giftShop_theme as m "
        sqlStr = sqlStr & "	join db_board.dbo.tbl_giftShop_theme_item as d "
        sqlStr = sqlStr & "		on m.themeIdx=d.themeIdx "
        sqlStr = sqlStr & "	join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on d.itemid=i.itemid "
        sqlStr = sqlStr & "where m.themeIdx=" & FRectIdx & " " & addSql & orderSql
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CGiftShopPrdItem

				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fbrandname			= rsget("brandname")
				FItemList(i).Fitemname			= rsget("itemname")
				FItemList(i).FOrgprice			= rsget("orgprice")
				FItemList(i).FSellCash 			= rsget("sellcash")
				FItemList(i).FLimitNo			= rsget("limitno")
				FItemList(i).FLimitSold			= rsget("LimitSold")
				FItemList(i).FisUsing			= rsget("isUsing")
				FItemList(i).FSellYn			= rsget("sellyn")
				FItemList(i).FSaleYn			= rsget("sailyn")
				FItemList(i).FLimitYn 			= rsget("limityn")
				FItemList(i).FItemCouponYN		= rsget("itemcouponyn")
				FItemList(i).FItemCouponType 	= rsget("itemcoupontype")
				FItemList(i).FItemCouponValue	= rsget("itemcouponvalue")
				FItemList(i).FRegdate 			= rsget("tmItmRegDt")

				if Not(rsget("basicimage")="" or isNull(rsget("basicimage"))) then
					FItemList(i).FBasicImage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("basicimage")
				end if
				if Not(rsget("smallimage")="" or isNull(rsget("smallimage"))) then
					FItemList(i).FSmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				end if


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub



    '# 코멘트 목록
	public Sub GetThemeCommentList()
		dim sqlStr, addSql, i

		'내글 보기
		if FRectIsMyItem="Y" and IsUserLoginOK then
			addSql = addSql & " and m.userid='" & GetLoginUserID & "' "
		end if

        '전체 카운트
        sqlStr = "select count(commentIdx), CEILING(CAST(Count(commentIdx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_comment as m "
        sqlStr = sqlStr & "Where m.themeIdx=" & FRectIdx & " and m.isUsing='Y' " & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " m.* "
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftShop_comment as m "
        sqlStr = sqlStr & "Where m.themeIdx=" & FRectIdx & " and m.isUsing='Y' " & addSql
        sqlStr = sqlStr & "Order by M.commentIdx desc "
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CGiftShopCommentItem

	            FItemList(i).FthemeIdx		= rsget("themeIdx")
	            FItemList(i).FcommentIdx	= rsget("commentIdx")
	            FItemList(i).Fcomment		= rsget("comment")
	            FItemList(i).Fuserid		= rsget("userid")
	            FItemList(i).Fregdate		= rsget("regdate")
	            FItemList(i).FregKind		= rsget("regKind")
	            FItemList(i).FisUsing		= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	'# 선물포장 상품 목록
	public Sub GetPackageList()

		dim sqlStr, addSql, orderSql, i

		'품절 포함 여부
		if FRectIsSoldOut<>"Y" then
			addSql = addSql & " and i.sellyn in ('Y','S')"
		end if

		orderSql = " order by i.itemid desc"

        '전체 카운트
        sqlStr = "select count(d.itemid), CEILING(CAST(Count(d.itemid) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "from db_board.dbo.tbl_giftShop_packInfo as m "
        sqlStr = sqlStr & "	join db_board.dbo.tbl_giftShop_packItem as d "
        sqlStr = sqlStr & "		on m.packIdx=d.packIdx "
        sqlStr = sqlStr & "	join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on d.itemid=i.itemid "
        sqlStr = sqlStr & "Where m.packIdx=" & FRectPackIdx & " " & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " i.*, d.regdate as packRegDt "
        sqlStr = sqlStr & "from db_board.dbo.tbl_giftShop_packInfo as m "
        sqlStr = sqlStr & "	join db_board.dbo.tbl_giftShop_packItem as d "
        sqlStr = sqlStr & "		on m.packIdx=d.packIdx "
        sqlStr = sqlStr & "	join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on d.itemid=i.itemid "
        sqlStr = sqlStr & "where m.packIdx=" & FRectPackIdx & " " & addSql & orderSql
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CGiftShopPrdItem

				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fbrandname			= rsget("brandname")
				FItemList(i).Fitemname			= rsget("itemname")
				FItemList(i).FOrgprice			= rsget("orgprice")
				FItemList(i).FSellCash 			= rsget("sellcash")
				FItemList(i).FLimitNo			= rsget("limitno")
				FItemList(i).FLimitSold			= rsget("LimitSold")
				FItemList(i).FisUsing			= rsget("isUsing")
				FItemList(i).FSellYn			= rsget("sellyn")
				FItemList(i).FSaleYn			= rsget("sailyn")
				FItemList(i).FLimitYn 			= rsget("limityn")
				FItemList(i).FItemCouponYN		= rsget("itemcouponyn")
				FItemList(i).FItemCouponType 	= rsget("itemcoupontype")
				FItemList(i).FItemCouponValue	= rsget("itemcouponvalue")
				FItemList(i).FRegdate 			= rsget("packRegDt")

				if Not(rsget("basicimage")="" or isNull(rsget("basicimage"))) then
					FItemList(i).FBasicImage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("basicimage")
				end if
				if Not(rsget("smallimage")="" or isNull(rsget("smallimage"))) then
					FItemList(i).FSmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	End Sub

	'------------------------------------------------
	'-- 클래스 기본설정 및 기타 함수
	'------------------------------------------------

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
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

'// gift 키워드 항목 출력 (클릭함수,선택값)
Function getGiftKeyword(vClk,arrChk)
	Dim strRst, sqlStr, i
	sqlStr = "Select keywordIdx, keywordname "
	sqlStr = sqlStr & "From db_board.dbo.tbl_gift_keyword "
	sqlStr = sqlStr & "Where isUsing='Y' and keywordType=1 "
	sqlStr = sqlStr & "Order by sortNo asc, keywordIdx asc "
	rsget.Open sqlStr, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		strRst = "<span class=""chkBox"">"
		i=1
		Do Until rsget.EOF
			strRst = strRst & "<input id=""chkKwd" & i & """ type=""checkbox"" onclick=""" & vClk & """ value=""" & rsget("keywordIdx") & """ " & chkIIF(chkArrValue(arrChk,rsget("keywordIdx")),"checked","") & "><label for=""chkKwd" & i & """>" & rsget("keywordname") & "</label>"
			if (i mod 10)=0 then strRst = strRst & "<br>"
			rsget.MoveNext
			i=i+1
		Loop
		strRst = strRst & "</span>"
	end if
	rsget.Close

	getGiftKeyword = strRst
End Function

'// gift Shop 선물포장 상품 구분
Function getGiftPackName(pid)
	Select Case cStr(pid)
		Case "1"
			getGiftPackName = "플라워"
		Case "2"
			getGiftPackName = "카드"
		Case "3"
			getGiftPackName = "포장지"
		Case "4"
			getGiftPackName = "선물상자"
		Case "5"
			getGiftPackName = "리본"
		Case "6"
			getGiftPackName = "악세사리"
	End Select
end Function
%>