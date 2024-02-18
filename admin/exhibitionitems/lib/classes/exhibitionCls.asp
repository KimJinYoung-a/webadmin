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
    	
    '// groupcode
    public Fgidx
    public Fgubuncode
    public Fmastercode
    public Fdetailcode
    public Ftitle
    public Fisusing

    '// common
    public Fregdate
    public Flastupdate
    public Fadminid
    public Flastadminid
	public FImageList
	public Fsorting
	public Faddtext1
	public Faddtext2
	public Foptioncode

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
	public Fetc_itemid
	public Fevt_kind
	public Ftitle
	public FbannerImage

	public function IsEndDateExpired()
        IsEndDateExpired = Cdate(Left(now(),10))>Cdate(Left(Fenddate,10))
    end function
End Class

'// 브랜드
Class ExhibitionBrandsCls
	public Fidx
	public Fmakerid
	public FsocName
	public FsocNameKor
	public Fmastercode
	public Fdetailcode
	public Fisusing
	public FsortNo
	public FbrandUsing
	public FmodelItem
	public FmodelImg
	public Fstartdate '// 시작일
	public Fenddate '// 종료일
	public Fregdate
	public FbannerImage

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
	public FRectEventKind
	
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
		sqlStr = sqlStr & " gidx , gubuncode , mastercode , detailcode , title , regdate , isusing" & vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_groupcode" & vbcrlf
		sqlStr = sqlStr & " where gidx=" & CStr(FrectGcode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ExhibitionItemsCls
        if Not rsget.Eof then

            FOneItem.Fgidx 			= rsget("gidx")
            FOneItem.Fgubuncode 	= rsget("gubuncode")
            FOneItem.Fmastercode	= rsget("mastercode")
            FOneItem.Fdetailcode 	= rsget("detailcode")
            FOneItem.Ftitle 		= rsget("title")
            FOneItem.Fregdate 		= rsget("regdate")
            FOneItem.Fisusing	 	= rsget("isusing")

        end if
        rsget.close
    end Sub

    '// 기획전 그룹 리스트
	public sub getGroupList()
		dim sqlStr,i, addSql

		if FrectIsusing<>"a" and FrectIsusing<>"" then
			addSql = addSql & " and isusing=" & FrectIsusing
		end if
		if FrectMasterCode<>"" then
			addSql = addSql & " and mastercode=" & FrectMasterCode
		end if

		'총 갯수 구하기
		sqlStr = "select" & vbcrlf
		sqlStr = sqlStr & " count(gidx) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_groupcode" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql
		

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " gidx , gubuncode , mastercode , detailcode , title , regdate , isusing" & vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_groupcode" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql
		sqlStr = sqlStr & " order by mastercode desc , detailcode asc " & vbcrlf

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
				FItemList(i).Ftitle 	= rsget("title")
				FItemList(i).Fregdate 	= rsget("regdate")
				FItemList(i).Fisusing 	= rsget("isusing")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Public Sub getExhibitionItem

		dim strSQL

		strSQL ="SELECT * FROM db_event.dbo.tbl_exhibition_items where idx =" & Frectidx

		'response.write strSQL
		rsget.open strSQL,dbget,1
		if not rsget.eof then
			set FItem = new ExhibitionItemsCls

			FItem.FIdx 			= rsget("idx")
			FItem.Fmastercode	= rsget("mastercode")
			FItem.Fdetailcode	= rsget("detailcode")
			FItem.Fitemid		= rsget("itemid")
			FItem.Fpickitem		= rsget("pickitem")

		end if
		rsget.close

	End Sub

	public Sub getItemsList()
		dim strSQL,i

		'갯수
		strSQL =" SELECT count(a.idx) as cnt" & vbcrlf
		strSQL = strSQL & " FROM db_event.dbo.tbl_exhibition_items a" & vbcrlf
		strSQL = strSQL & " left join db_item.dbo.tbl_item b" & vbcrlf
		strSQL = strSQL & " on a.itemid = b.itemid" & vbcrlf
		strSQL = strSQL & " where 1=1" & vbcrlf

		if FrectMasterCode > 0 then
		strSQL = strSQL & " and a.mastercode= '"& FrectMasterCode &"'" & vbcrlf
		end if

		if FrectDetailCode > 0 then
		strSQL = strSQL & " and a.detailcode= '"& FrectDetailCode &"'" & vbcrlf
		end if

		IF FrectMakerid<>"" Then
			strSQL = strSQL & " and b.makerid= '"& FrectMakerid &"' " & vbcrlf
		End IF

		IF Frectpick <>"" Then
			strSQL = strSQL & " and a.pickitem = '"& Frectpick &"' " & vbcrlf
		End IF

		''상품코드 검색
        if (FRectArrItemid <> "") then
            if right(trim(FRectArrItemid),1)="," then
            	FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" & Left(FRectArrItemid,Len(FRectArrItemid)-1) & ")"
            else
				FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" & FRectArrItemid & ")"
            end if
        end if

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = ""
		strSQL = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		strSQL = strSQL & " a.idx, a.mastercode, a.detailcode, a.ItemID, a.RegDate " & vbcrlf
		strSQL = strSQL & " ,b.itemname, b.makerid, b.smallimage, a.pickitem " & vbcrlf
		strSQL = strSQL & " ,isnull(b.orgprice,0) as orgprice , isnull(b.sailprice,0) as sailprice , b.sailyn , b.itemcouponyn , b.itemcoupontype , isnull(b.sailsuplycash,0) as sailsuplycash , isnull(b.orgsuplycash ,0) as orgsuplycash , " & vbcrlf
		strSQL = strSQL & " Case b.itemCouponyn When 'Y' then ( Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=b.curritemcouponidx and itemid=b.itemid ) end as couponbuyprice , b.mwdiv , b.deliverytype , isnull(b.sellcash,0) as sellcash  , isnull(b.buycash,0) as buycash , b.itemcouponvalue"
		strSQL = strSQL & " ,a.addtext1 , a.addtext2 " & vbcrlf
		strSQL = strSQL & " FROM db_event.dbo.tbl_exhibition_items a" & vbcrlf
		strSQL = strSQL & " left join db_item.dbo.tbl_item b" & vbcrlf
		strSQL = strSQL & " on a.itemid = b.itemid" & vbcrlf

		strSQL = strSQL & " where 1=1" & vbcrlf

		if FrectMasterCode > 0 then
		strSQL = strSQL & " and a.mastercode= '"& FrectMasterCode &"'" & vbcrlf
		end if

		if FrectDetailCode > 0 then
		strSQL = strSQL & " and a.detailcode= '"& FrectDetailCode &"'" & vbcrlf
		end if

		IF FrectMakerid<>"" Then
			strSQL = strSQL & " and b.makerid= '"& FrectMakerid &"' " & vbcrlf
		End IF

		IF Frectpick <>"" Then
			strSQL = strSQL & " and a.pickitem = '"& Frectpick &"' " & vbcrlf
		End IF

		''상품코드 검색 기능 수정 2015-09-15 유태욱
        if (FRectArrItemid <> "") then
            if right(trim(FRectArrItemid),1)="," then
            	FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" & Left(FRectArrItemid,Len(FRectArrItemid)-1) & ")"
            else
				FRectArrItemid = Replace(FRectArrItemid,",,",",")
            	strSQL = strSQL & " and a.itemid in (" & FRectArrItemid & ")"
            end if
        end if

		strSQL = strSQL & " order by a.idx desc" & vbcrlf

		'response.write strSQL
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
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ExhibitionItemsCls

				FItemList(i).Fidx 			= rsget("idx")
				FItemList(i).Fgubun 		= rsget("mastercode")
				FItemList(i).Fcategory 		= rsget("detailcode")
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).FRegDate 		= rsget("regdate")
				
				FItemLIst(i).FImageList 	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallimage")
				
				FItemLIst(i).Fitemname 		= db2html(rsget("itemname"))
				FItemLIst(i).FMakerid 		= db2html(rsget("makerid"))
				FItemLIst(i).Fpickitem 		= rsget("pickitem")

				FItemLIst(i).Forgprice 		= rsget("orgprice")
				FItemLIst(i).Fsailprice 	= rsget("sailprice")
				FItemLIst(i).Fsailyn 		= rsget("sailyn")
				FItemLIst(i).Fitemcouponyn  = rsget("itemcouponyn")
				FItemLIst(i).Fitemcoupontype= rsget("itemcoupontype")
				FItemLIst(i).Fsailsuplycash = rsget("sailsuplycash")
				FItemLIst(i).Forgsuplycash 	= rsget("orgsuplycash")
				FItemLIst(i).Fcouponbuyprice= rsget("couponbuyprice")
				FItemLIst(i).FmwDiv 		= rsget("mwDiv")
				FItemLIst(i).Fdeliverytype 	= rsget("deliverytype")
				FItemLIst(i).Fsellcash 		= rsget("sellcash")
				FItemLIst(i).Fbuycash 		= rsget("buycash")
				FItemLIst(i).Fmastercode 	= rsget("mastercode")
				FItemLIst(i).Fdetailcode 	= rsget("detailcode")
				FItemLIst(i).Fitemcouponvalue 	= rsget("itemcouponvalue")
				FItemLIst(i).Faddtext1 		= rsget("addtext1")
				FItemLIst(i).Faddtext2	 	= rsget("addtext2")
	
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

    '// Pick 아이템 및 정렬 popup
	public Sub getExhibitionBestItemList()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "SELECT count(*) as cnt FROM db_event.dbo.tbl_exhibition_items AS a " & vbcrlf
		strSQL = strSQL & " LEFT OUTER JOIN db_item.dbo.tbl_item AS b " & vbcrlf
		strSQL = strSQL & " ON a.itemid = b.itemid " & vbcrlf
		if FrectDetailCode = 0 then 
			strSQL = strSQL & " WHERE a.pickitem = 1 and mastercode = "& FrectMasterCode &" " & vbcrlf
		else
			strSQL = strSQL & " WHERE mastercode = "& FrectMasterCode &" and detailcode = "& FrectDetailcode &" " & vbcrlf
		end if 

		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = "SELECT a.idx , a.ItemID , b.itemname , b.basicimage " & vbcrlf
		if FrectDetailCode = 0 then 
			strSQL = strSQL & ", a.picksorting as sorting " & vbcrlf
		else
			strSQL = strSQL & ", a.categorysorting as sorting " & vbcrlf
		end if 

		strSQL = strSQL & ", a.pickitem ,a.RegDate , a.optioncode "
		strSQL = strSQL & " FROM db_event.dbo.tbl_exhibition_items AS a " & vbcrlf
		strSQL = strSQL & " LEFT OUTER JOIN db_item.dbo.tbl_item AS b " & vbcrlf
		strSQL = strSQL & " ON a.itemid = b.itemid " & vbcrlf
		if FrectDetailCode = 0 then 
			strSQL = strSQL & " WHERE a.pickitem = 1 and mastercode = "& FrectMasterCode &" " & vbcrlf
		else
			strSQL = strSQL & " WHERE mastercode = "& FrectMasterCode &" and detailcode = "& FrectDetailcode &" " & vbcrlf
		end if 

		if FrectDetailCode = 0 then
			strSQL = strSQL & " ORDER BY picksorting ASC " & vbcrlf
		else
			strSQL = strSQL & " ORDER BY categorysorting ASC " & vbcrlf
		end if 

		'response.write strSQL
		rsget.open strSQL,dbget,1

		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new ExhibitionItemsCls

				FItemList(i).Fidx 	    = rsget("idx")
				FItemList(i).Fitemid 	= rsget("ItemID")
				FItemLIst(i).fitemname 	= db2html(rsget("itemname"))
				FItemLIst(i).FImageList = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("basicimage")
				FItemLIst(i).Fsorting 	= rsget("sorting")
				FItemList(i).FisUsing 	= rsget("pickitem")
				FItemList(i).FRegDate 	= rsget("RegDate")
				FItemList(i).Foptioncode= rsget("optioncode")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------
' event
'---------------------------------------------------------------------------------------------------------------------------------------------------------
	'// one event
	public Sub getOneEventContents()
        dim sqlStr
        sqlStr = "select a.idx, a.evt_code, e.evt_name, e.evt_subcopyk, ed.etc_itemimg as squareimage"
		sqlStr = sqlStr & " , ed.evt_mo_listbanner as rectangleimage, ed.saleper, ed.salecper"
		sqlStr = sqlStr & " , a.startdate, a.enddate, e.evt_startdate, e.evt_enddate, a.evtsorting, a.isusing, a.banner_image"
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_eventgroup as a"
		sqlStr = sqlStr & " INNER JOIN db_event.dbo.tbl_event as e"
		sqlStr = sqlStr & " on a.evt_code = e.evt_code"
		sqlStr = sqlStr & " LEFT JOIN db_event.dbo.tbl_event_display ed"
		sqlStr = sqlStr & " on ed.evt_code = a.evt_code"
        sqlStr = sqlStr & " where a.idx=" & CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ExhibitionEventsCls

        if Not rsget.Eof then
    		FOneItem.Fidx			= rsget("idx")
			FOneItem.Fevt_code		= rsget("evt_code")
			FOneItem.Fevt_name		= rsget("evt_name")
			FOneItem.Fevt_subcopy	= rsget("evt_subcopyk")
			FOneItem.Fsquareimage	= rsget("squareimage")
			FOneItem.Frectangleimage= rsget("rectangleimage")
			FOneItem.Fsaleper		= rsget("saleper")
			FOneItem.Fsalecper		= rsget("salecper")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fevt_startdate	= rsget("evt_startdate")
			FOneItem.Fevt_enddate	= rsget("evt_enddate")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fevtsorting	= rsget("evtsorting")
			FOneItem.FbannerImage	= rsget("banner_image")
        end if
        rsget.Close
    end Sub

	'// event list
	public Sub getEventList()
        dim sqlStr, addSql, i

		if FRectEventKind <> "" then
			addSql = addSql & " and e.evt_kind = " & CStr(FRectEventKind)
		end if 

        if FRectmastercode > 0 then
            addSql = addSql & " and a.mastercode=" & CStr(FRectmastercode)
        end if

        if FRectIsusing<>"" then
            addSql = addSql & " and a.isusing=" & CStr(FRectIsusing) & ""
        end if

        if FRectSelDate<>"" then
            addSql = addSql & " and '" & FRectSelDate & "' between convert(varchar(10),a.startdate,120) and convert(varchar(10),a.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from db_event.dbo.tbl_exhibition_eventgroup as a WITH(NOLOCK) "
		sqlStr = sqlStr & " INNER JOIN db_event.dbo.tbl_event as e WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.evt_code = e.evt_code"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & addSql

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

       	sqlStr = "select a.idx , a.evt_code , e.evt_name , e.evt_subcopyk , ed.etc_itemimg as squareimage  "
		sqlStr = sqlStr & " , ed.evt_mo_listbanner as rectangleimage , ed.saleper , ed.salecper"
		sqlStr = sqlStr & " , a.startdate , a.enddate , e.evt_startdate , e.evt_enddate , a.evtsorting , a.isusing , ed.etc_itemid , e.evt_kind, a.banner_image"
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_eventgroup as a WITH(NOLOCK)"
		sqlStr = sqlStr & " INNER JOIN db_event.dbo.tbl_event as e WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.evt_code = e.evt_code"
		sqlStr = sqlStr & " LEFT JOIN db_event.dbo.tbl_event_display ed WITH(NOLOCK)"
		sqlStr = sqlStr & " on ed.evt_code = a.evt_code"
		sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
   		sqlStr = sqlStr & " order by a.evtsorting asc, a.idx desc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new ExhibitionEventsCls

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fevt_code		= rsget("evt_code")
				FItemList(i).Fevt_name		= rsget("evt_name")
				FItemList(i).Fevt_subcopy	= rsget("evt_subcopyk")
				FItemList(i).Fsquareimage	= rsget("squareimage")
				FItemList(i).Frectangleimage= rsget("rectangleimage")
				FItemList(i).Fsaleper		= rsget("saleper")
				FItemList(i).Fsalecper		= rsget("salecper")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).Fevt_startdate	= rsget("evt_startdate")
				FItemList(i).Fevt_enddate	= rsget("evt_enddate")
				FItemList(i).Fevtsorting 	= rsget("evtsorting")
				FItemList(i).Fisusing	 	= rsget("isusing")
				FItemList(i).Fetc_itemid	= rsget("etc_itemid")
				FItemList(i).Fevt_kind		= rsget("evt_kind")
				FItemList(i).FbannerImage		= rsget("banner_image")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public Sub getEventLinkList()
        dim sqlStr, addSql, i

        if FRectmastercode > 0 then
            addSql = addSql & " and a.mastercode=" & CStr(FRectmastercode)
        end if

        if FRectIsusing<>"" then
            addSql = addSql & " and a.isusing=" & CStr(FRectIsusing) & ""
        end if

        if FRectSelDate<>"" then
            addSql = addSql & " and '" & FRectSelDate & "' between convert(varchar(10),a.startdate,120) and convert(varchar(10),a.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from db_event.dbo.tbl_exhibition_event_link as a WITH(NOLOCK) "
		sqlStr = sqlStr & " INNER JOIN db_event.dbo.tbl_event as e WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.evt_code = e.evt_code"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & addSql

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

       	sqlStr = "select a.idx, a.evt_code, a.title, ed.etc_itemimg as squareimage"
		sqlStr = sqlStr & " , ed.evt_mo_listbanner as rectangleimage, ed.saleper, ed.salecper"
		sqlStr = sqlStr & " , a.startdate, a.enddate, e.evt_startdate, e.evt_enddate, a.sorting, a.isusing, ed.etc_itemid, e.evt_kind"
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_event_link as a WITH(NOLOCK)"
		sqlStr = sqlStr & " INNER JOIN db_event.dbo.tbl_event as e WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.evt_code = e.evt_code"
		sqlStr = sqlStr & " LEFT JOIN db_event.dbo.tbl_event_display ed WITH(NOLOCK)"
		sqlStr = sqlStr & " on ed.evt_code = a.evt_code"
		sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
   		sqlStr = sqlStr & " order by a.sorting asc, a.idx desc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new ExhibitionEventsCls

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fevt_code		= rsget("evt_code")
				FItemList(i).Ftitle			= rsget("title")
				FItemList(i).Fsquareimage	= rsget("squareimage")
				FItemList(i).Frectangleimage= rsget("rectangleimage")
				FItemList(i).Fsaleper		= rsget("saleper")
				FItemList(i).Fsalecper		= rsget("salecper")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).Fevt_startdate	= rsget("evt_startdate")
				FItemList(i).Fevt_enddate	= rsget("evt_enddate")
				FItemList(i).Fevtsorting 	= rsget("sorting")
				FItemList(i).Fisusing	 	= rsget("isusing")
				FItemList(i).Fetc_itemid	= rsget("etc_itemid")
				FItemList(i).Fevt_kind		= rsget("evt_kind")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public Sub getOneEventLinkContents()
        dim sqlStr
        sqlStr = "select a.idx, a.evt_code, a.title, e.evt_name, ed.etc_itemimg as squareimage,"
		sqlStr = sqlStr & " ed.evt_mo_listbanner as rectangleimage, ed.saleper, ed.salecper,"
		sqlStr = sqlStr & " a.startdate, a.enddate, e.evt_startdate, e.evt_enddate, a.sorting, a.isusing"
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_event_link as a"
		sqlStr = sqlStr & " INNER JOIN db_event.dbo.tbl_event as e"
		sqlStr = sqlStr & " on a.evt_code = e.evt_code"
		sqlStr = sqlStr & " LEFT JOIN db_event.dbo.tbl_event_display ed"
		sqlStr = sqlStr & " on ed.evt_code = a.evt_code"
        sqlStr = sqlStr & " where a.idx=" & CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ExhibitionEventsCls

        if Not rsget.Eof then
    		FOneItem.Fidx			= rsget("idx")
			FOneItem.Fevt_code		= rsget("evt_code")
			FOneItem.Ftitle			= rsget("title")
			FOneItem.Fevt_name		= rsget("evt_name")
			FOneItem.Fsquareimage	= rsget("squareimage")
			FOneItem.Frectangleimage= rsget("rectangleimage")
			FOneItem.Fsaleper		= rsget("saleper")
			FOneItem.Fsalecper		= rsget("salecper")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fevt_startdate	= rsget("evt_startdate")
			FOneItem.Fevt_enddate	= rsget("evt_enddate")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fevtsorting	= rsget("sorting")
        end if
        rsget.Close
    end Sub

	'// brand list
	public Sub getBrandList()
        dim sqlStr, addSql, i

        if FRectmastercode > 0 then
            addSql = addSql & " and a.mastercode=" & CStr(FRectmastercode)
        end if

        if FRectIsusing<>"" then
            addSql = addSql & " and a.isusing=" & CStr(FRectIsusing) & ""
        end if

        if FRectSelDate<>"" then
            addSql = addSql & " and '" & FRectSelDate & "' between convert(varchar(10),a.startdate,120) and convert(varchar(10),a.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from db_event.dbo.tbl_exhibition_brandgroup as a WITH(NOLOCK) "
		sqlStr = sqlStr & " INNER JOIN db_user.dbo.tbl_user_c as c WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.makerid = c.userid "
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & addSql

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

       	sqlStr = "select a.idx , a.makerid , c.socname , c.socname_kor, c.modelItem , c.modelBImg, c.isUsing as brandUsing "
		sqlStr = sqlStr & " , a.startdate , a.enddate , a.sortNo , a.isusing, a.banner_image "
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_brandgroup as a WITH(NOLOCK)"
		sqlStr = sqlStr & " INNER JOIN db_user.dbo.tbl_user_c as c WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.makerid = c.userid "
		sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
   		sqlStr = sqlStr & " order by a.sortNo asc, a.idx desc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new ExhibitionBrandsCls

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fmakerid		= rsget("makerid")
				FItemList(i).Fsocname		= rsget("socname")
				FItemList(i).FsocNameKor	= rsget("socname_kor")
				FItemList(i).FmodelItem		= rsget("modelItem")
				FItemList(i).FmodelImg		= rsget("modelBImg")
				FItemList(i).FbrandUsing	= rsget("brandUsing")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).FsortNo	 	= rsget("sortNo")
				FItemList(i).Fisusing	 	= rsget("isusing")

				if not (FItemList(i).FmodelImg="" or isNull(FItemList(i).FmodelImg)) then
					FItemList(i).FmodelImg = webImgUrl & "/image/list/" & GetImageSubFolderByItemid(FItemList(i).FmodelItem) & "/" & FItemList(i).FmodelImg
				end if

				FItemList(i).FbannerImage		= rsget("banner_image")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'// one brand
	public Sub getOneBrandContents()
        dim sqlStr
       	sqlStr = "select a.idx , a.makerid , c.socname , c.socname_kor, c.modelItem , c.modelBImg, c.isUsing as brandUsing "
		sqlStr = sqlStr & " , a.startdate , a.enddate , a.sortNo , a.isusing, a.banner_image"
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_brandgroup as a"
		sqlStr = sqlStr & " INNER JOIN db_user.dbo.tbl_user_c as c WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.makerid = c.userid "
        sqlStr = sqlStr & " where a.idx=" & CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ExhibitionBrandsCls

        if Not rsget.Eof then
			FOneItem.Fidx			= rsget("idx")
			FOneItem.Fmakerid		= rsget("makerid")
			FOneItem.Fsocname		= rsget("socname")
			FOneItem.FsocNameKor	= rsget("socname_kor")
			FOneItem.FmodelItem		= rsget("modelItem")
			FOneItem.FmodelImg		= rsget("modelBImg")
			FOneItem.FbrandUsing	= rsget("brandUsing")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.FsortNo	 	= rsget("sortNo")
			FOneItem.Fisusing	 	= rsget("isusing")
			FOneItem.FbannerImage	= rsget("banner_image")
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

End Class



'---------------------------------------------------------------------------------------------------------------------------------------------------------
' function 
'---------------------------------------------------------------------------------------------------------------------------------------------------------
function DrawSelectAllView(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select mastercode , title from db_event.dbo.tbl_exhibition_groupcode where detailcode = 0 order by title asc "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("mastercode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("mastercode")&"' "&tmp_str&">" & db2html(rsget("title")) & "</option>")
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
   query1 = " select mastercode , title from db_event.dbo.tbl_exhibition_groupcode where isusing = 1 and detailcode = 0 order by mastercode DESC "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("mastercode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("mastercode")&"' "&tmp_str&">" & db2html(rsget("title")) & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

function getMasterCodeName(mastercode)
	dim codename,query1
	query1 = " select top 1 title from db_event.dbo.tbl_exhibition_groupcode where isusing = 1 and mastercode = "& mastercode &" and detailcode = 0 "
   	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		codename = rsget("title")
	end if
	rsget.close

	getMasterCodeName = codename
end function

function getDetailCodeName(mastercode,detailcode)
	dim codename,query1
	query1 = " select top 1 title from db_event.dbo.tbl_exhibition_groupcode where isusing = 1 and mastercode = "& mastercode &" and detailcode = "& detailcode &""
   	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		codename = rsget("title")
	end if
	rsget.close

	getDetailCodeName = codename
end function

function DrawDetailButtons(mastercode,jscall,menutitle)
	dim query
	query = " select mastercode , detailcode , title from db_event.dbo.tbl_exhibition_groupcode where isusing = 1 and mastercode = "& mastercode &" order by detailcode asc "
	rsget.Open query,dbget,1
	if  not rsget.EOF  then
       do until rsget.EOF
			if rsget("detailcode") > 0 then 
			response.write "<input type=""button"" value="""& rsget("title") &""" onclick="""&jscall&"('"& rsget("mastercode") &"','"& rsget("detailcode") &"','"& menutitle &"');"" class=""button""> "
			end if 
        	rsget.MoveNext
		loop
	end if
	rsget.close
end function

function DrawDetailSelectBox(selectBoxName,selectedId,masterCode)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select detailcode , title from db_event.dbo.tbl_exhibition_groupcode where isusing = 1 and mastercode = '"& mastercode &"' and detailcode <> 0 order by detailcode asc "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("detailcode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("detailcode")&"' "&tmp_str&">" & db2html(rsget("title")) & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function
%>