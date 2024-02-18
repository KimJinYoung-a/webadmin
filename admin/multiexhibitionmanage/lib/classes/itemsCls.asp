<%
'// 상품
Class ExhibitionItemsCls
    '// items
    public Fidx
    public Fgubun
    public Fcategory
    public Fitemid
    public Fpickitem
    public Fpicksorting
    public Fcategorysorting
	public Fitemname
	public FMakerid
	public Forgprice
	public Fsailprice
	public Fsailyn
	public Fitemcouponyn
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fsailsuplycash
	public Forgsuplycash
	public Fcouponbuyprice
	public FmwDiv
	public Fdeliverytype
	public Fsellcash
	public Fbuycash
	public Fitemscore
	public Foptcnt
	public Foptioncode
    	
    '// groupcode
    public Fgidx
    public Fgubuncode
    public Fmastercode
    public Fdetailcode
    public Ftypename
    public Fisusing
	public Ftype

    '// common
    public Fregdate
    public Flastupdate
    public Fadminid
    public Flastadminid
	public FImageList
	public Fsorting

	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FSellCash
	end Function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function
End Class

'// 이벤트
Class ExhibitionEventsCls
	public Fidx
	public Fevt_name
	public Fevt_code
	public Fmastercode
	public Fdetailcode
	public Fisusing
	public Fevtsorting
	public Fevt_subcopy
	public Fsquareimage '// PC 정사각 이미지
	public Frectangleimage '// mobile 직사각 이미지
	public Fsaleper '// 할인가
	public Fsalecper '// 쿠폰 할인가
	public Fstartdate '// 시작일
	public Fenddate '// 종료일
	public Fevt_startdate '// 이벤트 시작일
	public Fevt_enddate '// 이벤트 종료일
	public Fregdate
	public Flastupdate
	public Fadminid
	public Flastadminid

	public function IsEndDateExpired()
        IsEndDateExpired = Cdate(Left(now(),10))>Cdate(Left(Fenddate,10))
    end function
End Class

Class ExhibitionCls

	Public FItemList()
	Public FItem
	public FResultCount
	public FPageSize
	public FCurrPage
	public Ftotalcount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FOneItem
	public Frectidx
	public FrectIsusing
	public FrectGcode
	public FrectCate
	public FrectMakerid
	public FrectArrItemid
	public Frectpick
    public FrectCategory
	public FrectFlagDate
	public FrectEvt_Code
	public FrectMasterCode
	public FrectDetailCode
	public FRectSelDate
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

    '// 기획전 그룹
    public Sub getOneGroupItem()
        dim SqlStr
        SqlStr = "select"
		sqlStr = sqlStr & " gidx , gubuncode , mastercode , detailcode , typename , type , regdate , isusing" & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_exhibitionevent_groupcode" & vbcrlf
		sqlStr = sqlStr & " where gidx=" + CStr(FrectGcode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ExhibitionItemsCls
        if Not rsget.Eof then

            FOneItem.Fgidx 			= rsget("gidx")
            FOneItem.Fgubuncode 	= rsget("gubuncode")
            FOneItem.Fmastercode	= rsget("mastercode")
            FOneItem.Fdetailcode 	= rsget("detailcode")
            FOneItem.Ftypename 		= rsget("typename")
			FOneItem.Ftype 			= rsget("type")
            FOneItem.Fregdate 		= rsget("regdate")
            FOneItem.Fisusing	 	= rsget("isusing")

        end if
        rsget.close
    end Sub

    '// 기획전 그룹 리스트
	public sub getGroupList()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" & vbcrlf
		sqlStr = sqlStr & " count(gidx) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_exhibitionevent_groupcode" & vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " gidx , gubuncode , mastercode , detailcode , typename , regdate , isusing , type" & vbcrlf
		sqlStr = sqlStr & " , DENSE_RANK() OVER (PARTITION BY gubuncode ORDER BY type ASC) as rankname " & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_exhibitionevent_groupcode" & vbcrlf
		sqlStr = sqlStr & " order by mastercode desc , gubuncode asc , rankname asc , typename asc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ExhibitionItemsCls

				FItemList(i).Fgidx 		= rsget("gidx")
				FItemList(i).Fgubuncode = rsget("gubuncode")
				FItemList(i).Fmastercode = rsget("mastercode")
				FItemList(i).Fdetailcode = rsget("detailcode")
				FItemList(i).Ftypename 	= rsget("typename")
				FItemList(i).Fregdate 	= rsget("regdate")
				FItemList(i).Fisusing 	= rsget("isusing")
				FItemList(i).Ftype 		= rsget("type")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// items list - 상품별
	public Sub getItemsList()
		dim strSQL,i
		dim addSql , itemAddSql , countSql , itemSql
		dim havingCnt : havingCnt = 0

		'' -----------------------------------------------------------------------------------------------------
		'' option
		'' -----------------------------------------------------------------------------------------------------
		'' 기획전 검색
		if FrectMasterCode > 0 then
		 	addSql = addSql & " and d.mastercode= '"& FrectMasterCode &"'" & vbcrlf
		end if
		
		'' 상품코드 검색
        if (FRectArrItemid <> "") then
            if right(trim(FRectArrItemid),1)="," then
            	FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	itemAddSql = itemAddSql & " and m.itemid in (" + Left(FRectArrItemid,Len(FRectArrItemid)-1) + ")"
            else
				FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	itemAddSql = itemAddSql & " and m.itemid in (" + FRectArrItemid + ")"
            end if
        end if

		'' 브랜드 검색
		IF FrectMakerid<>"" Then
			itemSql = itemSql & " and i.makerid= '"& FrectMakerid &"' " & vbcrlf
		End IF

		'' 옵션코드 검색
		if FrectDetailCode <> "" then
			addSql = addSql & " and d.detailcode in ("& replace(FrectDetailCode,chr(32),"") &")" & vbcrlf
			havingCnt = split(replace(FrectDetailCode,chr(32),""),",")
			havingCnt = ubound(havingCnt)
			'havingCnt = 0
		end if
		'' -----------------------------------------------------------------------------------------------------
		
		'갯수
		strSQL =" SELECT count(*) AS cnt " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_exhibition_item_master AS m WITH(NOLOCK) " & vbcrlf
		strSQL = strSQL & " WHERE EXISTS ( " & vbcrlf
		strSQL = strSQL & " 	SELECT 1 FROM " & vbcrlf
		strSQL = strSQL & "			( " & vbcrlf
		strSQL = strSQL & " 			SELECT d.itemid , d.detailcode FROM db_item.dbo.tbl_exhibition_item_detail AS d WITH(NOLOCK) " & vbcrlf
		strSQL = strSQL & " 			WHERE d.itemid = m.itemid " & addSql & itemSql & vbcrlf
		strSQL = strSQL & " 			GROUP BY d.itemid , d.detailcode "  & vbcrlf
		strSQL = strSQL & " 		) AS td "  & vbcrlf
		strSQL = strSQL & " 		INNER JOIN db_item.dbo.tbl_item as i WITH(NOLOCK) " & vbcrlf
		strSQL = strSQL & " 		ON td.itemid = i.itemid " & vbcrlf
		strSQL = strSQL & " 		HAVING count(*) > "& havingCnt & vbcrlf
		strSQL = strSQL & " )" & itemAddSql

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = ""
		if FCurrPage = 1 then 
		strSQL = "SELECT TOP " & Cstr(FPageSize * FCurrPage) & vbcrlf
		else
		strSQL = "SELECT "& vbcrlf
		end if 
		strSQL = strSQL & " m.idx , m.mastercode , m.itemid , m.itemscore , cd.* "& vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_exhibition_item_master AS m WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & " INNER JOIN "& vbcrlf
		strSQL = strSQL & " 	( "& vbcrlf
		strSQL = strSQL & "  		SELECT td.itemid , i.itemname , i.makerid , i.smallimage "& vbcrlf
		strSQL = strSQL & " 			,isnull(i.orgprice,0) as orgprice , isnull(i.sailprice,0) as sailprice , i.sailyn "& vbcrlf
		strSQL = strSQL & " 			,i.itemcouponyn , i.itemcoupontype , isnull(i.sailsuplycash,0) as sailsuplycash , isnull(i.orgsuplycash ,0) as orgsuplycash "& vbcrlf
		strSQL = strSQL & "  			,CASE i.itemCouponyn When 'Y' then ( Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=td.itemid ) end as couponbuyprice "& vbcrlf
		strSQL = strSQL & "  			,i.mwdiv ,i.deliverytype , isnull(i.sellcash,0) as sellcash ,isnull(i.buycash,0) as buycash , i.itemcouponvalue , sum(td.optcnt) as optcnt "& vbcrlf
		strSQL = strSQL & "  		FROM "& vbcrlf
		strSQL = strSQL & "  		( "& vbcrlf
		strSQL = strSQL & "  			SELECT d.itemid , d.detailcode , count(*) as optcnt FROM db_item.dbo.tbl_exhibition_item_detail AS d WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & "  			WHERE 1=1 "& addSql & vbcrlf
		strSQL = strSQL & "  			GROUP BY d.itemid , d.detailcode"& vbcrlf
		strSQL = strSQL & "  		) AS td"& vbcrlf
		strSQL = strSQL & "  		INNER JOIN db_item.dbo.tbl_item AS i WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & "  		ON td.itemid = i.itemid"& vbcrlf
		strSQL = strSQL & "  		WHERE 1=1 "& itemSql & vbcrlf
		strSQL = strSQL & " 		GROUP BY td.itemid , i.itemname ,i.makerid ,i.smallimage "& vbcrlf
		strSQL = strSQL & "					,i.orgprice , i.sailprice ,i.sailyn ,i.itemcouponyn ,i.itemcoupontype ,i.sailsuplycash "& vbcrlf
		strSQL = strSQL & " 				,i.orgsuplycash ,i.itemcouponyn ,i.mwdiv ,i.deliverytype ,i.sellcash ,i.buycash ,i.itemcouponvalue ,i.curritemcouponidx "& vbcrlf
		strSQL = strSQL & "  		HAVING count(*) > "& havingCnt & vbcrlf
		strSQL = strSQL & "  	) AS cd "& vbcrlf
		strSQL = strSQL & " ON m.itemid = cd.itemid "& vbcrlf
		strSQL = strSQL & " WHERE 1=1"& itemAddSql &vbcrlf
		strSQL = strSQL & " ORDER BY m.idx DESC "& vbcrlf
		if FCurrPage > 1 then 
		strSQL = strSQL & " OFFSET "& (FCurrPage-1) * FPageSize &" ROWS FETCH NEXT "& FPageSize &" ROWS ONLY"& vbcrlf
		end if 

		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		rsget.PageSize= FPageSize
		i=0
		if  not rsget.EOF  then
			'rsget.absolutepage = FCurrPage '// 페이징 쿼리를 바꿔서 필요 없음
			do until rsget.EOF
				set FItemList(i) = new ExhibitionItemsCls

				FItemList(i).Fidx 				= rsget("idx")
				FItemList(i).Fmastercode		= rsget("mastercode")
				FItemList(i).Fitemid 			= rsget("itemid")
				FItemList(i).Fitemscore			= rsget("itemscore")
				FItemLIst(i).Fitemname 			= db2html(rsget("itemname"))
				FItemLIst(i).FMakerid 			= db2html(rsget("makerid"))
				FItemLIst(i).FImageList 		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemLIst(i).Forgprice 			= rsget("orgprice")
				FItemLIst(i).Fsailprice 		= rsget("sailprice")
				FItemLIst(i).Fsailyn 			= rsget("sailyn")
				FItemLIst(i).Fitemcouponyn  	= rsget("itemcouponyn")
				FItemLIst(i).Fitemcoupontype	= rsget("itemcoupontype")
				FItemLIst(i).Fsailsuplycash 	= rsget("sailsuplycash")
				FItemLIst(i).Forgsuplycash 		= rsget("orgsuplycash")
				FItemLIst(i).Fcouponbuyprice	= rsget("couponbuyprice")
				FItemLIst(i).FmwDiv 			= rsget("mwDiv")
				FItemLIst(i).Fdeliverytype 		= rsget("deliverytype")
				FItemLIst(i).Fsellcash 			= rsget("sellcash")
				FItemLIst(i).Fbuycash 			= rsget("buycash")
				FItemLIst(i).Fitemcouponvalue 	= rsget("itemcouponvalue")
				FItemLIst(i).Foptcnt 			= rsget("optcnt")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

	'// items list - 옵션별
	public Sub getOptionItemsList()
		dim strSQL,i
		dim addSql , itemAddSql , countSql , itemSql
		dim havingCnt : havingCnt = 0

		'' -----------------------------------------------------------------------------------------------------
		'' option
		'' -----------------------------------------------------------------------------------------------------
		'' 기획전 검색
		if FrectMasterCode > 0 then
		 	addSql = addSql & " and d.mastercode= '"& FrectMasterCode &"'" & vbcrlf
		end if
		
		'' 상품코드 검색
        if (FRectArrItemid <> "") then
            if right(trim(FRectArrItemid),1)="," then
            	FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	itemAddSql = itemAddSql & " and m.itemid in (" + Left(FRectArrItemid,Len(FRectArrItemid)-1) + ")"
            else
				FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	itemAddSql = itemAddSql & " and m.itemid in (" + FRectArrItemid + ")"
            end if
        end if

		'' 브랜드 검색
		IF FrectMakerid<>"" Then
			itemSql = itemSql & " and i.makerid= '"& FrectMakerid &"' " & vbcrlf
		End IF

		'' 옵션코드 검색
		if FrectDetailCode <> "" then
			addSql = addSql & " and d.detailcode in ("& replace(FrectDetailCode,chr(32),"") &")" & vbcrlf
			havingCnt = split(replace(FrectDetailCode,chr(32),""),",")
			havingCnt = ubound(havingCnt)
			'havingCnt = 0
		end if
		'' -----------------------------------------------------------------------------------------------------
		
		'갯수
		strSQL =" SELECT count(*) AS cnt " & vbcrlf
		strSQL = strSQL & " FROM " & vbcrlf
		strSQL = strSQL & " (" & vbcrlf
		strSQL = strSQL & "  	SELECT id.itemid FROM" & vbcrlf
		strSQL = strSQL & "  	(" & vbcrlf
		strSQL = strSQL & "  		SELECT d.itemid , d.detailcode FROM db_item.dbo.tbl_exhibition_item_detail AS d WITH(NOLOCK)" & vbcrlf
		strSQL = strSQL & "  		WHERE 1=1 "& addSql & vbcrlf
		strSQL = strSQL & "  		GROUP BY d.itemid , d.detailcode " & vbcrlf
		strSQL = strSQL & "  	) AS id" & vbcrlf
		strSQL = strSQL & "  	GROUP BY id.itemid" & vbcrlf
		strSQL = strSQL & "  	HAVING count(*) >"& havingCnt & vbcrlf
		strSQL = strSQL & " ) AS td" & vbcrlf
		strSQL = strSQL & " CROSS APPLY ( " & vbcrlf
		strSQL = strSQL & " 	SELECT d.itemid FROM db_item.dbo.tbl_exhibition_item_detail AS d WITH(NOLOCK) " & vbcrlf
		strSQL = strSQL & "  	WHERE d.itemid = td.itemid "& addSql & vbcrlf
		strSQL = strSQL & " ) as dt " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_exhibition_item_master AS m WITH(NOLOCK) " & vbcrlf
		strSQL = strSQL & " ON td.itemid = m.itemid "& itemAddSql & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item AS i WITH(NOLOCK) " & vbcrlf
		strSQL = strSQL & " ON td.itemid = i.itemid "& itemSql & vbcrlf

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'rw strSQL

		'데이터 리스트
		strSQL = ""
		if FCurrPage = 1 then 
			strSQL = "SELECT TOP " & Cstr(FPageSize * FCurrPage) & vbcrlf
		else
			strSQL = "SELECT "& vbcrlf
		end if 
		strSQL = strSQL & " 	dt.idx , m.mastercode , dt.detailcode , m.itemscore , dt.gubuncode , td.itemid , dt.optioncode , i.itemname , i.makerid , i.smallimage "& vbcrlf
		strSQL = strSQL & " 	,isnull(i.orgprice,0) as orgprice , isnull(i.sailprice,0) as sailprice , i.sailyn "& vbcrlf
		strSQL = strSQL & " 	,i.itemcouponyn , i.itemcoupontype , isnull(i.sailsuplycash,0) as sailsuplycash , isnull(i.orgsuplycash ,0) as orgsuplycash "& vbcrlf
		strSQL = strSQL & " 	,CASE i.itemCouponyn When 'Y' then ( Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=td.itemid ) end as couponbuyprice "& vbcrlf
		strSQL = strSQL & " 	,i.mwdiv ,i.deliverytype , isnull(i.sellcash,0) as sellcash ,isnull(i.buycash,0) as buycash , i.itemcouponvalue , iopt.optaddprice "& vbcrlf
		strSQL = strSQL & " FROM "& vbcrlf
		strSQL = strSQL & " ("& vbcrlf
		strSQL = strSQL & " 	SELECT id.itemid FROM"& vbcrlf
		strSQL = strSQL & " 	("& vbcrlf
		strSQL = strSQL & " 		SELECT d.itemid , d.detailcode FROM db_item.dbo.tbl_exhibition_item_detail AS d WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & " 		WHERE 1=1 "& addSql  & vbcrlf
		strSQL = strSQL & " 		GROUP BY d.itemid , d.detailcode "& vbcrlf
		strSQL = strSQL & " 	) AS id"& vbcrlf
		strSQL = strSQL & " 	GROUP BY id.itemid "& vbcrlf
		strSQL = strSQL & " 	HAVING count(*) > "& havingCnt & vbcrlf
		strSQL = strSQL & " ) AS TD "& vbcrlf
		strSQL = strSQL & " CROSS APPLY ("& vbcrlf
		strSQL = strSQL & " 		SELECT d.idx, d.detailcode , d.gubuncode , d.optioncode FROM db_item.dbo.tbl_exhibition_item_detail AS d WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & " 		WHERE d.itemid = td.itemid "& addSql & vbcrlf
		strSQL = strSQL & " ) AS dt"& vbcrlf
		strSQL = strSQL & " OUTER APPLY ("& vbcrlf
		strSQL = strSQL & " 		SELECT optaddprice FROM db_item.dbo.tbl_item_option WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & " 		WHERE itemid = td.itemid and itemoption = dt.optioncode "& vbcrlf
		strSQL = strSQL & " ) AS iopt "& vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_exhibition_item_master AS m WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & " ON td.itemid = m.itemid "& itemAddSql & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item AS i WITH(NOLOCK) "& vbcrlf
		strSQL = strSQL & " ON td.itemid = i.itemid "& itemSql & vbcrlf
		strSQL = strSQL & " ORDER BY m.idx DESC "& vbcrlf
		if FCurrPage > 1 then 
		strSQL = strSQL & " OFFSET "& (FCurrPage-1) * FPageSize &" ROWS FETCH NEXT "& FPageSize &" ROWS ONLY"& vbcrlf
		end if 

		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		rsget.PageSize= FPageSize
		i=0
		if  not rsget.EOF  then
			'rsget.absolutepage = FCurrPage '// 페이징 쿼리를 바꿔서 필요 없음
			do until rsget.EOF
				set FItemList(i) = new ExhibitionItemsCls

				FItemList(i).Fidx 				= rsget("idx")
				FItemList(i).Fmastercode		= rsget("mastercode")
				FItemList(i).Fdetailcode		= rsget("detailcode")
				FItemList(i).Fgubuncode			= rsget("gubuncode")
				FItemList(i).Fitemid 			= rsget("itemid")
				FItemList(i).Foptioncode		= rsget("optioncode")
				FItemList(i).Fitemscore			= rsget("itemscore")
				FItemLIst(i).Fitemname 			= db2html(rsget("itemname"))
				FItemLIst(i).FMakerid 			= db2html(rsget("makerid"))
				FItemLIst(i).FImageList 		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemLIst(i).Forgprice 			= rsget("orgprice")
				FItemLIst(i).Fsailprice 		= rsget("sailprice")
				FItemLIst(i).Fsailyn 			= rsget("sailyn")
				FItemLIst(i).Fitemcouponyn  	= rsget("itemcouponyn")
				FItemLIst(i).Fitemcoupontype	= rsget("itemcoupontype")
				FItemLIst(i).Fsailsuplycash 	= rsget("sailsuplycash")
				FItemLIst(i).Forgsuplycash 		= rsget("orgsuplycash")
				FItemLIst(i).Fcouponbuyprice	= rsget("couponbuyprice")
				FItemLIst(i).FmwDiv 			= rsget("mwDiv")
				FItemLIst(i).Fdeliverytype 		= rsget("deliverytype")
				FItemLIst(i).Fsellcash 			= rsget("sellcash")
				FItemLIst(i).Fbuycash 			= rsget("buycash")
				FItemLIst(i).Fitemcouponvalue 	= rsget("itemcouponvalue")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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

'---------------------------------------------------------------------------------------------------------------------------------------------------------
' function 
'---------------------------------------------------------------------------------------------------------------------------------------------------------
function DrawSelectAllView(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%=chkiif(changeFlag<>"","onchange='"&changeFlag&"(this.value);'","") %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select mastercode , typename from db_item.dbo.tbl_exhibitionevent_groupcode where detailcode = 0 order by typename asc "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("mastercode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("mastercode")&"' "&tmp_str&">" + db2html(rsget("typename")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select mastercode , typename from db_item.dbo.tbl_exhibitionevent_groupcode where isusing = 1 and detailcode = 0 order by typename asc "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("mastercode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("mastercode")&"' "&tmp_str&">" + db2html(rsget("typename")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

function getMasterCodeName(mastercode)
	dim codename,query1
	query1 = " select top 1 typename from db_item.dbo.tbl_exhibitionevent_groupcode where isusing = 1 and mastercode = "& mastercode &" and detailcode = 0 "
   	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		codename = rsget("typename")
	end if
	rsget.close

	getMasterCodeName = codename
end function

function getDetailCodeName(mastercode,detailcode)
	dim codename,query1
	query1 = " select top 1 typename from db_item.dbo.tbl_exhibitionevent_groupcode where isusing = 1 and mastercode = "& mastercode &" and detailcode = "& detailcode &""
   	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		codename = rsget("typename")
	end if
	rsget.close

	getDetailCodeName = codename
end function

function DrawDetailButtons(mastercode,jscall,menutitle)
	dim query
	query = " select mastercode , detailcode , typename from db_item.dbo.tbl_exhibitionevent_groupcode where isusing = 1 and mastercode = "& mastercode &" order by detailcode asc "
	rsget.Open query,dbget,1
	if  not rsget.EOF  then
       do until rsget.EOF
			if rsget("detailcode") > 0 then 
			response.write "<input type=""button"" value="""& rsget("typename") &""" onclick="""&jscall&"('"& rsget("mastercode") &"','"& rsget("detailcode") &"','"& menutitle &"');"" class=""button""> "
			end if 
        	rsget.MoveNext
		loop
	end if
	rsget.close
end function

'// array 중복 제거
Function DuplValRemove(ByVal varArr)
	Dim dic, items, rtnVal

	Set dic = CreateObject("Scripting.Dictionary")
	dic.removeall
	dic.CompareMode = 0

	For Each items In varArr
		If not dic.Exists(items) Then dic.Add items, items
	Next

	rtnVal = dic.keys
	Set dic = Nothing
	DuplValRemove = rtnVal
End Function
%>