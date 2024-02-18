<%
'####################################################
' Description :  사은품 관리
' History : 2008.04.01 정윤정 생성
'			2013.11.11 한용민 수정
'####################################################

Class CGiftKindOption
    public Fgift_kind_code
    public Fgift_kind_option
    public Fgift_kind_optionName
    public Fgift_kind_LimitYN
    public Fgift_kind_Limit
    public Fgift_kind_LimitSold
    public Fgift_kind_optionUsing

    public Fprd_itemgubun
    public Fprd_itemid
    public Fprd_itemoption

	public forderserial
	public freqname
	public fuserid
	public fgiftkind_cnt
	public fipkumdivname
	public fgift_name
	public fEventConditionStr
	public fgiftexistsyn
	public fgiftserviceyn

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CGift
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FGCode
	public FECode
	public FTotCnt
	public FCPage
	public FPSize
	public FSearchTxt
	public FSearchType
	public FGiftName
	public FBrand
	public FDateType
	public FSDate
	public FEDate
	public FGStatus
	public FGUsing
	public FGName
	public FGScope
	public FEGroupCode
	public FGType
	public FGRange1
	public FGRange2
	public FGKindCode
	public FGKindType
	public FGKindCnt
	public FGKindlimit
	public FRegdate
	public FAdminid
	public FGKindName
    public FGKindImg
	public FGDelivery
	public FItemid
	public FOpenDate
	public FCloseDate
	public FOldKindName
	public FSiteScope
	public FPartnerID

	''201004추가
	public Fimage120
	public Fimage400List
	public Fgiftkind_givecnt
	public Fgiftkind_linkGbn
	public Fbcouponidx
	public Fprd_itemgubun
	public Fprd_itemid
	public Fprd_itemoption
	
	public frectCategory
	public frectCategoryMid
	public frectDispCategory
	public frectrunoutrate90up
	public frectorderserial
	public frectuserid
	public frectreqname
	public FGiftIsusing
	public FGiftImage1
	public FGiftText1
	public FGiftImage2
	public FGiftText2
	public FGiftImage3
	public FGiftText3
	public FGiftInfoText
	public frectgift_code
	public frectgiftexistsyn
	public frectipkumdiv

	'== 사은품 리스트 가져오기		'//admin/shopmaster/gift/giftList.asp
	public Function fnGetGiftList
		Dim strSqlCnt, strSql, strSearch,iDelCnt

		strSearch = ""
		IF FECode <> "" THEN
			strSearch = " and a.evt_code ="&FECode
		END IF

		IF FSearchTxt <> "" THEN
			IF FSearchType = 1 THEN
				strSearch = strSearch & " and gift_code = "&FSearchTxt
			ELSE
				strSearch = strSearch & " and a.evt_code = "&FSearchTxt
			END IF
		END IF

		IF FGiftName <> "" THEN
				strSearch = strSearch & " and gift_name like '%"&FGiftName&"%'"
		END IF

		IF FBrand <> "" THEN
				strSearch = strSearch & " and makerid = '"&FBrand&"'"
		END IF

		If FItemid<>"" then
			strSearch = strSearch & " and gift_code in (select distinct gift_code from [db_event].[dbo].tbl_giftItem where itemid="&FItemid&" and giftitem_using=1)"
		end if

		IF FSDate <> "" AND FEDate <> "" THEN
			if CStr(FDateType) = "S" THEN
				strSearch  = strSearch & " and  datediff(day, '"&FSDate&"', gift_startdate) >= 0 and  datediff(day,'"&FEDate&"', gift_startdate) <=0  "
			elseif CStr(FDateType) = "E" THEN
				strSearch  = strSearch & " and  datediff(day,'"&FSDate&"',gift_enddate) >= 0 and  datediff(day,'"&FEDate&"',gift_enddate) <=0  "
			end if
		END IF

		IF FGStatus <> "" THEN
			IF FGStatus = 9 THEN
				strSearch = strSearch & " and ( gift_status = "&FGStatus&" or  datediff(day,getdate(),gift_enddate)< 0 ) "
			ELSEIF FGStatus = 6 THEN	'오픈예정
				strSearch  = strSearch & " and   gift_status = 7 and  datediff(day,getdate(),gift_startdate)<= 0 and datediff(day,getdate(),gift_enddate) >= 0  "
			ELSEIF FGStatus = 7 THEN	'오픈진행중
				strSearch  = strSearch & " and   gift_status = 7 and  datediff(day,getdate(),gift_startdate)> 0 and  datediff(day,getdate(),gift_enddate)>=0 "
			ELSE
				strSearch = strSearch & " and  gift_status = "&FGStatus&" AND  datediff(day,getdate(),gift_enddate)>=0  "
			END IF
		END IF

		IF FGDelivery <> "" THEN
			strSearch = strSearch & " and gift_delivery = '"&FGDelivery&"'"
		END IF
		
		if frectrunoutrate90up<>"" then
			strSearch = strSearch & " and isnull(giftkind_givecnt,0)<>0 and isnull(giftkind_limit,0)<>0"
			strSearch = strSearch & " and (case"
			strSearch = strSearch & " 		when isnull(giftkind_givecnt,0)<>0 and isnull(giftkind_limit,0)<>0"
			strSearch = strSearch & " 			then round( (convert(float,giftkind_givecnt)/convert(float,giftkind_limit))*100 ,2)"
			strSearch = strSearch & " 		else 0"
			strSearch = strSearch & " 		end) >= 90"
		end if
		
		if frectCategory<>"" then
			if frectCategory="110" then
				strSearch = strSearch & " and ed.evt_category='"& frectCategory &"'"
				
				if frectCategoryMid<>"" then
					strSearch = strSearch & " and ed.evt_catemid='"& frectCategoryMid &"'"
				end if
			else
				strSearch = strSearch & " and ed.evt_category='"& frectCategory &"'"
			end if
		end if
		
		if frectDispCategory <> "" then
			strSearch = strSearch & " and ed.evt_dispcate='"& frectDispCategory &"'"
		end if
		
		strSqlCnt = " SELECT COUNT(gift_code) "&_
					" FROM [db_event].[dbo].[tbl_gift] AS A "&_
					" left outer join [db_event].[dbo].[tbl_giftkind] AS B "&_
					" 		ON A.giftkind_code = B.giftkind_code "&_
					" left join db_event.dbo.tbl_event_display ed "&_
					" 		on a.evt_code=ed.evt_code "&_
					" WHERE 1=1 " & strSearch
		
		'response.write strSqlCnt & "<Br>"
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT TOP "&FPSize&"  [gift_code], [gift_name], [gift_scope], a.evt_code, [evtgroup_code], [makerid], [gift_type]"&_
					"		, [gift_range1], [gift_range2], A.[giftkind_code]"&_
					"  		, [giftkind_type], [giftkind_cnt], [giftkind_limit], [gift_startdate], [gift_enddate]"&_
					"		, [gift_status] = Case When DateDiff(day,getdate(),gift_enddate) < 0 Then 9 "&_
					 "							When A.gift_status = 7 and DateDiff(day,getdate(),gift_startdate) <= 0 Then 6 "&_
					"							ELSE gift_status end "&_
					"		, A.[regdate], [gift_using], [adminid], B.giftkind_name "&_
					"		, gift_cnt = Case gift_scope when 2 then (select count(itemid) from [db_event].[dbo].[tbl_eventitem] WHERE evt_code = A.evt_code)"&_
					"									when 4 then (select count(itemid) from [db_event].[dbo].[tbl_eventitem] WHERE evt_code = A.evt_code AND evtgroup_code = A.evtgroup_code)"&_
					"									when 5 then (select count(itemid) from [db_event].[dbo].[tbl_giftitem] WHERE gift_code = A.gift_code)	"&_
					"									else 0 end "&_
					"		,gift_delivery, opendate, closedate, giftkind_givecnt "&_
					" 		,(case  "&_
					"			when isnull(giftkind_givecnt,0)<>0 and isnull(giftkind_limit,0)<>0 "&_
					"				then round( (convert(float,giftkind_givecnt)/convert(float,giftkind_limit))*100 ,2) "&_
					"			else 0 "&_
					"			end) as runoutrate, b.giftkind_linkgbn,b.prd_itemgubun, b.prd_itemid, b.prd_itemoption, b.bcouponidx "&_
					" FROM [db_event].[dbo].[tbl_gift] AS A "&_
					" left outer join [db_event].[dbo].[tbl_giftkind] AS B "&_
					" 		ON A.giftkind_code = B.giftkind_code "&_
					" left join db_event.dbo.tbl_event_display ed "&_
					" 		on a.evt_code=ed.evt_code "&_
					" WHERE gift_code<=(  "&_
					" 						SELECT Min(gift_code) "&_ 
					" 						FROM ( "&_ 
					" 							SELECT TOP "&iDelCnt&" gift_code "&_
					" 							FROM [db_event].[dbo].[tbl_gift] AS A "&_
					" 							left outer join [db_event].[dbo].[tbl_giftkind] AS B "&_
					" 								ON A.giftkind_code = B.giftkind_code "&_
					" 							left join db_event.dbo.tbl_event_display ed "&_
					" 								on a.evt_code=ed.evt_code "&_					
					" 							WHERE 1=1 "&strSearch&" ORDER BY gift_code DESC "&_
					" 						) as T "&_
					" 					) "&strSearch&" ORDER BY gift_code DESC "

			'response.write strSql & "<Br>"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetGiftList = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function

	'== 사은품 종류 검색하기  top 100 으로 막음. search like '%[blabla]%' ==> 정규표현식으로 검색..?? []포함된건 어떻게 검색..? ==> [ => [[], % => [%]
	public Function fnGetGiftKind
		Dim strSql
		Dim isStringConverted, realSearchText

		if (FPSize = "") then
			FPSize = 100
		end if

		if IsNull(FSearchTxt) then
			FSearchTxt = ""
		end if

		isStringConverted = False
		if (InStr(FSearchTxt, "[[]") > 0) or (InStr(FSearchTxt, "[%]") > 0) then
			isStringConverted = True
			realSearchText = FSearchTxt
		end if

		if (isStringConverted = False) then
			realSearchText = Replace(FSearchTxt, "[", "[[]")
			realSearchText = Replace(realSearchText, "%", "[%]")
		end if

		strSql = " SELECT top " + CStr(FPSize) + " giftkind_code, giftkind_name, giftkind_img, itemid, regdate, giftkind_linkGbn, bcouponidx, reguserid, org_gift_code "
		strSql = strSql & " FROM  [db_event].[dbo].[tbl_giftkind] "
		strSql = strSql & "  WHERE  giftkind_name like '%" & realSearchText & "%' "
		strSql = strSql & "  order by giftkind_code desc "
		rsget.Open strSql,dbget

		IF not rsget.EOF THEN
			fnGetGiftKind = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'== 사은품 종류 내용보기
	public Function fnGetGiftKindConts
		Dim strSql, i
		strSql = " SELECT giftkind_name, giftkind_img, itemid, image120, regdate, giftkind_linkGbn, bcouponidx, prd_itemgubun, prd_itemid, prd_itemoption FROM  [db_event].[dbo].[tbl_giftkind] "&_
				" WHERE giftkind_code = " & FGKindCode
		
		'response.write strSql & "<Br>"
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FGKindName	   		= rsget("giftkind_name")
			FGKindImg      		= rsget("giftkind_img")
			FItemid        		= rsget("itemid")
			Fimage120      		= rsget("image120")
			Fgiftkind_linkGbn 	= rsget("giftkind_linkGbn")
			Fbcouponidx     	= rsget("bcouponidx")

			Fprd_itemgubun     	= rsget("prd_itemgubun")
			Fprd_itemid     	= rsget("prd_itemid")
			Fprd_itemoption     = rsget("prd_itemoption")
		End IF
		rsget.Close

	End Function

	'== 사은품 추가이미지
	public Function fnGetGiftKindAddImage
	    Dim strSql, i
		strSql = " SELECT * FROM  [db_event].[dbo].[tbl_giftkind_addimage] WHERE gift_kind_code=" & FGKindCode
	    rsget.Open strSql,dbget , 1
	    i=0
	    FResultCount = rsget.RecordCount
		IF not rsget.EOF THEN
		    ReDim preserve Fimage400List(FResultCount)
		    Do until rsget.EOF
		        Fimage400List(i)=rsget("gift_kind_addimage")
    			i=i+1
    			rsget.moveNext
			loop
		End IF
		rsget.Close
	END Function

	''==사은품 옵션
	public Function fnGetGiftKindOptions
	    Dim strSql, i
		strSql = " SELECT * FROM  [db_event].[dbo].[tbl_giftkind_option] "&_
				" WHERE gift_kind_code = " & FGKindCode &_
				" Order by gift_kind_code"
		rsget.Open strSql,dbget , 1
		FTotCnt      = rsget.recordCount
		FResultCount = FTotCnt
		ReDim FItemList(FResultCount)

		i=0
		IF not rsget.EOF THEN
		    do until rsget.eof

				set FItemList(i) = new CGiftKindOption

				FItemList(i).Fgift_kind_code        = rsget("gift_kind_code")
    			FItemList(i).Fgift_kind_option      = rsget("gift_kind_option")
    			FItemList(i).Fgift_kind_optionName  = db2Html(rsget("gift_kind_optionName"))
    			FItemList(i).Fgift_kind_LimitYN     = rsget("gift_kind_LimitYN")
    			FItemList(i).Fgift_kind_Limit       = rsget("gift_kind_Limit")
    			FItemList(i).Fgift_kind_LimitSold   = rsget("gift_kind_LimitSold")
    			FItemList(i).Fgift_kind_optionUsing = rsget("gift_kind_optionUsing")

				FItemList(i).Fprd_itemgubun 		= rsget("prd_itemgubun")
				FItemList(i).Fprd_itemid 			= rsget("prd_itemid")
				FItemList(i).Fprd_itemoption 		= rsget("prd_itemoption")

				i=i+1
				rsget.moveNext
			loop

		End IF
		rsget.Close
	End Function

	' /admin/shopmaster/gift/giftuserdetail.asp
	'== 사은품 내용 보기
	public Function fnGetGiftConts
		Dim strSql

		if FGCode="" or isnull(FGCode) then exit Function

		strSql ="   SELECT  [gift_code], [gift_name], [gift_scope], [evt_code], [evtgroup_code], [makerid], [gift_type], [gift_range1], [gift_range2], A.[giftkind_code]"&_
				"  		, [giftkind_type], [giftkind_cnt], [giftkind_limit],[giftkind_givecnt], [gift_startdate], [gift_enddate], [gift_status], A.[regdate], [gift_using], [adminid]"&_
				"		, B.giftkind_name, B.giftkind_img, B.giftkind_linkGbn, B.bcouponidx, gift_delivery, opendate, closedate,lastupdate, A.gift_itemname, A.site_scope, A.partner_id "&_
				" FROM [db_event].[dbo].[tbl_gift] AS A left outer join [db_event].[dbo].[tbl_giftkind] AS B ON A.giftkind_code = B.giftkind_code "&_
				" WHERE  gift_code = "&FGCode

		'response.write strSql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		IF not rsget.EOF THEN
			FGName  	= rsget("gift_name")
			FGScope 	= rsget("gift_scope")
			FECode  	= rsget("evt_code")
			FEGroupCode = rsget("evtgroup_code")
			FBrand      = rsget("makerid")
			FGType      = rsget("gift_type")
			FGRange1    = rsget("gift_range1")
			FGRange2    = rsget("gift_range2")
			FGKindCode  = rsget("giftkind_code")
			FGKindType  = rsget("giftkind_type")
			FGKindCnt   = rsget("giftkind_cnt")
			FGKindlimit = rsget("giftkind_limit")
			FSDate   	= rsget("gift_startdate")
			FEDate     	= rsget("gift_enddate")
			FGStatus    = rsget("gift_status")
			FGUsing     = rsget("gift_using")
			IF datediff("d",FEDate,now) > 0  THEN FGStatus = 9	'종료일이 지난 경우 종료로 표기
			'IF (datediff("d",FEDate,now) <= 0 and datediff("d",FSDate,now)>=0  and FGStatus=7) THEN FGStatus = 6
			FRegdate    = rsget("regdate")
			FAdminid    = rsget("adminid")
			FGKindName	= rsget("giftkind_name")
			FGKindImg	= rsget("giftkind_img")
			FGDelivery  = rsget("gift_delivery")
			FOpenDate	= rsget("opendate")
			FCloseDate	= rsget("closedate")
			FOldKindName= rsget("gift_itemname")

			FSiteScope 	= rsget("site_scope")
			FPartnerID	= rsget("partner_id")
			Fgiftkind_givecnt = rsget("giftkind_givecnt")

			Fbcouponidx = rsget("bcouponidx")
			Fgiftkind_linkGbn = rsget("giftkind_linkGbn")
		END IF
		rsget.close
	End Function

	' /admin/shopmaster/gift/giftuserdetail.asp
	'== 사은품 한정수량
	public Function fnLimitgiftCount
		Dim strSql

		strSql = "select (CASE WHEN m.ipkumdiv='2' THEN '주문접수' " &_
					"			 WHEN m.ipkumdiv='4' THEN '결제완료' " &_
					"			 WHEN m.ipkumdiv in ('5','6') THEN '상품준비' " &_
					"			 WHEN m.ipkumdiv in ('7') THEN '일부출고' " &_
					"			 WHEN m.ipkumdiv in ('8') THEN '출고완료' " &_
					"			ELSE '?' END),sum(giftKind_cnt)  " &_
					"	from db_order.dbo.tbl_order_gift g " &_
					"		Join db_order.dbo.tbl_order_master m " &_
					"		on g.orderserial=m.orderserial " &_
					"		and m.cancelyn='N' " &_
					"	where isNULL(chg_gift_code,gift_code)= " &FGCode&_
					"	group by (CASE WHEN m.ipkumdiv='2' THEN '주문접수' " &_
					"			 WHEN m.ipkumdiv='4' THEN '결제완료' " &_
					"			 WHEN m.ipkumdiv in ('5','6') THEN '상품준비' " &_
					"			 WHEN m.ipkumdiv in ('7') THEN '일부출고' " &_
					"			 WHEN m.ipkumdiv in ('8') THEN '출고완료' " &_
					"			ELSE '?' END)  "

		'response.write strSql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			fnLimitgiftCount = rsget.getRows()
		End IF
		rsget.Close
	End Function

	' /admin/shopmaster/gift/giftuserdetail.asp
	' 사은품 대상자 리스트	' 2019.09.25 한용민 생성
	public sub fngiftuserlist()
		dim sqlStr,i, sqlsearch

		if isnull(frectgift_code) or frectgift_code="" then exit sub

		sqlStr = "exec db_order.dbo.sp_ten_order_giftuser_list "& FPageSize &","& FCurrPage &","& frectgift_code &",'"& frectgiftexistsyn &"','"& frectipkumdiv &"','"& frectorderserial &"','"& frectuserid &"','"& frectreqname &"',''"

		if session("ssBctId")="tozzinet" then
		response.write sqlStr & "<Br>"
		else
		'response.write sqlStr & "<Br>"
		end if
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount
		' if (FCurrPage * FPageSize < FTotalCount) then
		' 	FResultCount = FPageSize
		' else
		' 	FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		' end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CGiftKindOption

				FItemList(i).forderserial = rsget("orderserial")
				FItemList(i).freqname = db2html(rsget("reqname"))
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fgiftkind_cnt = rsget("giftkind_cnt")
				FItemList(i).fipkumdivname = rsget("ipkumdivname")
				FItemList(i).fgift_name = db2html(rsget("gift_name"))
				FItemList(i).fEventConditionStr = db2html(rsget("EventConditionStr"))
				FItemList(i).fgiftexistsyn = rsget("giftexistsyn")
				FItemList(i).fgiftserviceyn = rsget("giftserviceyn")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	' /admin/shopmaster/gift/giftuserdetail.asp
	'## fnGetEventGiftBox : 이벤트 사은품 박스 정보 가져오기 ##
	public Function fnGetEventGiftBox
		Dim strSql

		strSql = " SELECT top 1 gift_text1, gift_img1, gift_text2, gift_img2, gift_text3, gift_img3, gift_isusing, contentsAlign FROM db_event.dbo.tbl_event_md_theme WHERE evt_code="+ Cstr(FECode)

		'response.write strSql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FGiftText1 = rsget("gift_text1")
			FGiftImage1 = rsget("gift_img1")
			FGiftText2 = rsget("gift_text2")
			FGiftImage2 = rsget("gift_img2")
			FGiftText3 = rsget("gift_text3")
			FGiftImage3 = rsget("gift_img3")
			FGiftIsusing = rsget("gift_isusing")
			FGiftInfoText = rsget("contentsAlign")
		End IF
		rsget.Close
	End Function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
        redim Fimage400List(0)
        FTotCnt = 0
	    FCPage  = 1
	    FPSize  = 20
	    FResultCount = 0
		FCurrPage =1
		FPageSize = 50
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

'################################################################
' 사은품 증정대상 상품
'################################################################
Class CGiftItem
	public FGCode
	public FTotCnt
	public FPSize
	public FCPage

	public Function fnGetItemConts
	 dim strSqlCnt,iDelCnt,strSql
	strSqlCnt = " SELECT COUNT(A.itemid) FROM [db_event].[dbo].[tbl_giftitem] AS A "&_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	WHERE A.gift_code = "&FGCode
	rsget.Open strSqlCnt,dbget
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close
	IF FTotCnt >0 THEN
		iDelCnt =  ((FCPage - 1) * FPSize )+1
		strSql = " SELECT  TOP "&FPSize&" A.itemid,  B.makerid, B.itemname, B.sellcash,B.buycash,B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.mileage "&_
				"		,B.smallimage, B.listimage,  B.sellyn, B.deliverytype,  B.limityn, B.danjongyn, B.sailyn, B.isusing, B.limitno , B.limitsold "&_
				"	    , B.itemcouponyn, B.itemcoupontype, B.itemcouponvalue"&_
				"		 , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end as couponbuyprice "&_
				"	FROM  [db_event].[dbo].[tbl_giftitem] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.gift_code = "&FGCode&"  and A.itemid <= (SELECT MIN(itemid) FROM (SELECT Top "&iDelCnt&" C.itemid FROM [db_event].[dbo].[tbl_giftitem] AS C "&_
				"	 	INNER JOIN [db_item].[dbo].tbl_item AS D ON C.itemid = D.itemid  WHERE gift_code = " &FGCode &" ORDER BY C.itemid DESC ) as T ) " &_
				"  ORDER BY A.itemid DESC "
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetItemConts = rsget.getRows()
		End IF
		rsget.Close

	END IF
	End Function

    public Function IsSoldOut(FSellYn,FLimitYn,FLimitNo,FLimitSold)
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa(FLimitNo,FLimitSold)<1))
	end function

    public function GetLimitEa(FLimitNo,FLimitSold)
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public Function IsUpcheBeasong(Fdeliverytype)
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function
End Class

Function fnSetDelivery(ByVal selValue)
	IF selValue = "Y" THEN
		fnSetDelivery = "업체"
	ELSE
		fnSetDelivery = "10x10"
	END IF
End Function

' 사은품구분명	' 2020.04.09 한용민 생성
Function getgift_deliveryname(ByVal selValue)
	dim gift_deliveryname

	if selValue="" or isnull(selValue) then exit Function
	IF selValue = "Y" THEN
		gift_deliveryname = "업체배송"
	elseIF selValue = "C" THEN
		gift_deliveryname = "쿠폰"
	ELSE
		gift_deliveryname = "텐바이텐배송"
	END IF
	getgift_deliveryname=gift_deliveryname
End Function
%>
