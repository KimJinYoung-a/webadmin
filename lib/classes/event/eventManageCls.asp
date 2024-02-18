<%
'####################################################
' Page : /lib/classes/event/eventManageCls.asp
' Description :  이벤트 관리
' History : 2007.02.07 정윤정 생성
'			2008.03.21 정윤정 수정
'           2008.10.20 상품이미지 크기 추가(허진원)
'           2009.05.14 중 카테고리 추가(허진원)
'           2010.01.25 담당MD 추가(허진원)
' /event/eventmanage/common/event_function.asp include 필수!
'####################################################

'------------------------------------------------------
'ClsEvent : 이벤트 내용
'------------------------------------------------------
Class ClsEvent
	public FECode	'해당 이벤트코드
	public FEKind
	public FEManager
	public FEScope
	public FEPartnerID
	public FEName
	public FESDay
	public FEEDay
	public FEPDay
	public FELevel
	public FEState
	public FERegdate
	public FECategory
	public FECateMid
	public FESale
	public FEGift
	public FECoupon
	public FECommnet
	public FEBbs
	public FEItemps
	public FEApply
	public FEBImg
	public FEBImg2010
	public FEGImg
	public FETemp
	public FEMImg
	public FEHtml
	public FEISort
	public FEIAddType
	public FEDId
	public FEMId
	public FEFwd
	public FChkDisp
	public FEBrand
	public FEIcon
	public FECommentTitle
	public FELinkCode

	public FELinkType
	public FELinkURL

	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt

	public FESGroup	'Set 그룹검색
	public FESSort	'Set 정렬

	public FSfDate
	public FSsDate
	public FSeDate
	public FSfEvt
	public FSeTxt
	public FScategory
	public FScateMid
	public FEDispCate
	public FSstate
	public FSkind
	public FSedid
	public FSemid

	public FSisSale
	public FSisGift
	public FSisCoupon
	public FSisOnlyTen
	public FSisGetBlogURL
	Public FSisDiary

	public FEUsing
	public FEOpenDate
	public FECloseDate

	public FRectMakerid
	public FRectItemid
	public FRectItemName

	public FRectSellYN
	public FRectIsUsing
	public FRectDanjongyn
	public FRectLimityn
	public FRectMWDiv
	public FRectDeliveryType
	public FRectSailYn
	public FRectCouponYn
	public FRectVatYn

	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small

	public FEKindDesc
	public FEStateDesc

	public FEFullYN
	public FEWideYN
	public FEIteminfoYN
	public FETag
	public FWorkTag

	public FPrizeYN

	public FEItempriceYN
	public FEBImgMobile
	public FEBImgMoToday
	public FEBImgMoListBanner

	Public FENameEng
	Public FsubcopyK
	Public FsubcopyE

	Public FEOneplusOne  '원+원
	Public FEFreedelivery   '무료배송
	Public FEBookingsell   '무료배송
	Public FENew
	public FESalePer
	public FECsalePer

	Public FEtcitemid '대표상품코드
	Public FEtcitemimg '대표상품이미지
	Public FEsortNo		'정렬번호(회차)
	Public FEdateview
	Public FEitemid

	Public FENamesub
	Public FEListType
	
	public FRectState
	public FRectSDate
	public FRectEDate
	public FRectDisp

	'## fnGetEventCont : 이벤트개요 내용 가져오기 ##
	public Function fnGetEventCont
	Dim strSql
	IF FECode = "" THEN Exit Function
		strSql = " SELECT  evt_kind, evt_manager, evt_scope, evt_name, evt_startdate, evt_enddate, evt_prizedate, evt_level, evt_state, evt_regdate, evt_using, opendate, closedate,partner_id "&_
				",(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'eventkind' and code_value = a.evt_kind) evt_kinddesc "&_
				",(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'eventstate' and code_value = a.evt_state) evt_statedesc, a.prizeyn "&_
				",evt_nameEng , evt_subcopyK , evt_subcopyE,evt_sortNo , evt_subname "&_
				" , (Case When A.evt_kind=13 Then (Select top 1 itemid from [db_event].[dbo].[tbl_eventitem] where evt_code=A.evt_code) else 0 end) as itemid  "& _
				" , B.issale, B.isgift, B.iscoupon, B.isOnlyTen, B.isoneplusone, B.isfreedelivery, B.isbookingsell, B.iscomment, B.ISNEW, B.SALEPER, B.SALECPER, b.brand "&_
				" FROM [db_event].[dbo].[tbl_event] a "&_
				" inner join db_event.dbo.tbl_event_display b on a.evt_code = b.evt_code "&_
				" WHERE a.evt_code = "&FECode
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FEKind 		= rsget("evt_kind")
		 	FEKindDesc	= rsget("evt_kinddesc")
			FEManager 	= rsget("evt_manager")
			FEScope 	= rsget("evt_scope")
			FEPartnerID	= rsget("partner_id")
			IF isNull(FEPartnerID) THEN FEPartnerID = ""
			FEName 		= rsget("evt_name")
			FESDay 		= rsget("evt_startdate")
			FEEDay 		= rsget("evt_enddate")
			FEPDay 		= rsget("evt_prizedate")
			IF FEPDay = "1900-01-01" THEN
				FEPDay = ""
			END IF
			FELevel 	= rsget("evt_level")
			FEState 	= rsget("evt_state")
			FEStateDesc = fnSetStatusDesc(FEState,FESDay,FEEDay, rsget("evt_statedesc"))

			FERegdate 	= rsget("evt_regdate")
			FEUsing 	= rsget("evt_using")
			FEOpenDate 	= rsget("opendate")
			FECloseDate	= rsget("closedate")
			FPrizeYN	= rsget("prizeyn")

			FENameEng =  rsget("evt_nameEng")
			FsubcopyK =  rsget("evt_subcopyK")
			FsubcopyE =  rsget("evt_subcopyE")
			FEsortNo	= rsget("evt_sortNo")
			FEitemid	= rsget("itemid")
			FENamesub	= rsget("evt_subname")

			'2019-05-13 값 추가 최종원
			FESale = rsget("issale")
			FEGift = rsget("isgift")
			FECoupon = rsget("iscoupon")
			FECommnet = rsget("iscomment")	
			FSisOnlyTen = rsget("isOnlyTen")
			FEOneplusOne = rsget("isoneplusone") 
			FEFreedelivery = rsget("isfreedelivery")
			FENew = rsget("isnew")
			FECsalePer = rsget("SALECPER")
			FESalePer = rsget("SALEPER")
			FEBrand = rsget("brand")
		End IF
		rsget.Close
	End Function


	'## fnGetEventDisplay :이벤트화면설정 내용가져오기 ##
	public Function fnGetEventDisplay
	Dim strSql
	IF FECode = "" THEN Exit Function
		strSql = " SELECT  evt_category, evt_cateMid, issale, isgift, iscoupon,iscomment,isbbs,isitemps, isapply, evt_bannerimg, evt_template,"&_
				"	evt_mainimg, evt_html, evt_itemsort, designerid, isNull(partMDid,'') as partMDid, evt_forward, brand, evt_icon, evt_comment,link_evtcode, evt_fullyn, evt_wideyn, evt_iteminfoyn,evt_giftimg "&_
				" 	,evt_bannerlink,evt_LinkType, evt_tag, evt_bannerimg2010, isOnlyTen, isGetBlogURL, workTag , evt_itempriceyn, evt_bannerimg_mo, isNull(evt_dispCate,'') evt_dispCate " &_
				" 	,isoneplusone , isfreedelivery , etc_itemid , etc_itemimg , isbookingsell, evt_dateview , evt_todaybanner , evt_mo_listbanner, evt_itemlisttype, isDiary" &_
				" FROM [db_event].[dbo].[tbl_event_display]  "&_
				" WHERE evt_code = "&FECode  
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FChkDisp = 1
			FECategory = rsget("evt_category")
			FECateMid = rsget("evt_cateMid")
			FEDispCate = rsget("evt_dispCate")
			FESale = rsget("issale")
			FEGift = rsget("isgift")
			FECoupon = rsget("iscoupon")
			FECommnet = rsget("iscomment")
			FEBbs = rsget("isbbs")
			FEItemps = rsget("isitemps")
			FEApply = rsget("isapply")
			FEBImg = rsget("evt_bannerimg")
			FEBImg2010 = rsget("evt_bannerimg2010")
			FEBImgMobile = rsget("evt_bannerimg_mo")
			FEGImg = rsget("evt_giftimg")
			FETemp = rsget("evt_template")
			FEMImg = rsget("evt_mainimg")
			FEHtml = rsget("evt_html")
			FEISort = rsget("evt_itemsort")
			FEDId = rsget("designerid")
			FEMId = rsget("partMDid")
			FEFwd = rsget("evt_forward")
			FEBrand = rsget("brand")
			FEIcon = rsget("evt_icon")

			FECommentTitle = rsget("evt_comment")
			FELinkCode = rsget("link_evtcode")

			FEFullYN = rsget("evt_fullyn")
			FEWideYN = rsget("evt_wideyn")
			FEIteminfoYN = rsget("evt_iteminfoyn")

			FELinkURL	= rsget("evt_bannerlink")
			FELinkType	= rsget("evt_LinkType")
			FETag		= rsget("evt_tag")
			
			FSisOnlyTen = rsget("isOnlyTen")
			FSisGetBlogURL = rsget("isGetBlogURL")
			FWorkTag	= rsget("workTag")
			FSisDiary = rsget("isDiary") '// 다이어리 상태값 추가

			FEItempriceYN = rsget("evt_itempriceyn") '특정 브랜드 할인상품가격 가리기를 원하여..-_-;; 

			FEOneplusOne = rsget("isoneplusone") '원+원 추가 2013-08-07
			FEFreedelivery = rsget("isfreedelivery") '무료배송 2013-08-07
			FEBookingsell = rsget("isbookingsell") '예약판매 2013-08-07

			FEtcitemid =  rsget("etc_itemid") '대표상품ID 추가 2013-08-07
			FEtcitemimg =  rsget("etc_itemimg") '대표상품이미지 추가 2013-08-07
			FEdateview = rsget("evt_dateview")

			FEBImgMoToday = rsget("evt_todaybanner")
			FEBImgMoListBanner = rsget("evt_mo_listbanner")
			FEListType = rsget("evt_itemlisttype")

		ELSE
			FChkDisp = 0
		End IF
		rsget.Close
	End Function

	'## fnGetEventItem :이벤트상품 가져오기 ##
	public Function fnGetEventItem

		Dim strSql, strSqlCnt,iDelCnt
		Dim strSort,strGroup, striSort,addSql

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and B.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            addSql = addSql & " and B.itemid in (" + FRectItemid + ")"
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and B.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectSellYN <> "") then
            addSql = addSql & " and B.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and B.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and B.danjongyn<>'Y'"
            addSql = addSql + " and B.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and B.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (B.mwdiv='M' or B.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and B.mwdiv='" + FRectMwDiv + "'"
        end if

		if FRectLimityn="Y0" then
            addSql = addSql + " and B.limityn='Y' and (B.limitno-B.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and B.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            addSql = addSql + " and B.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and B.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and B.cate_small='" + FRectCate_Small + "'"
        end if

        if FRectSailYn<>"" then
            addSql = addSql + " and B.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and B.itemcouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and B.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and B.deliverytype='" + FRectDeliveryType + "'"
        end if


	IF FESGroup <> "" THEN
		IF FESGroup = 0 THEN
			strGroup = " AND (evtgroup_code  is null OR evtgroup_code =0 )"
		ELSE
			strGroup = " AND evtgroup_code = "&FESGroup
		END IF
	END IF

	IF FESSort = "slsell" THEN
		strSort = " ORDER BY evtitem_imgsize desc, sellcash asc "
		striSort =	" ORDER BY evtitem_imgsize desc, sellcash asc "
	ELSEIF FESSort = "shsell" THEN
		strSort = " ORDER BY  evtitem_imgsize desc, sellcash desc "
		striSort = " ORDER BY  evtitem_imgsize desc, sellcash desc "
	ELSEIF FESSort = "sbest" THEN
		strSort = " ORDER BY evtitem_imgsize desc, recentsellcount desc, sellcash desc "
		striSort = " ORDER BY evtitem_imgsize desc, recentsellcount desc, sellcash desc "
	ELSEIF FESSort = "sevtitem" THEN
		strSort = " ORDER BY evtitem_imgsize desc, evtitem_sort ,A.itemid desc"
		striSort = " ORDER BY evtitem_imgsize desc, evtitem_sort ,C.itemid desc"
	ELSEIF FESSort = "sevtgroup" THEN
		strSort = " ORDER BY evtitem_imgsize desc, evtgroup_code "
		striSort = " ORDER BY evtitem_imgsize desc, evtgroup_code "
	ELSEIF FESSort = "sbrand" THEN
		strSort = " ORDER BY evtitem_imgsize desc, makerid "
		striSort = " ORDER BY evtitem_imgsize desc, makerid "
	ELSE
		strSort = " ORDER BY evtitem_imgsize desc, A.itemid DESC "
		striSort = " ORDER BY evtitem_imgsize desc, C.itemid DESC "
	END IF

	strSqlCnt = " SELECT COUNT(A.itemid) FROM [db_event].[dbo].[tbl_eventitem] AS A "&_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	WHERE A.evt_code = "&FECode& strGroup&addSql
				 
	rsget.Open strSqlCnt,dbget,1
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close
	IF FTotCnt >0 THEN
		iDelCnt =  (FCPage - 1) * FPSize
'		strSql = " SELECT  TOP "&FPSize&" A.itemid, A.evtgroup_code, A.evtitem_sort,  B.makerid, B.itemname, B.sellcash "&_
'				"		,B.buycash,B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.mileage, B.smallimage, B.listimage,   B.sellyn, B.deliverytype "&_
'				"	    ,  B.limityn, B.danjongyn, B.sailyn, B.isusing, B.limitno , B.limitsold, B.itemcouponyn, B.itemcoupontype, B.itemcouponvalue"&_
'				"		 , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end as couponbuyprice "&_
'				"		, B.mwdiv, A.evtitem_imgsize	"&_
'				"	FROM  [db_event].[dbo].[tbl_eventitem] AS A " &_
'				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
'				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
'				"	WHERE A.evt_code = "&FECode&"  and A.itemid not in (SELECT Top "&iDelCnt&" C.itemid FROM [db_event].[dbo].[tbl_eventitem] AS C "&_
'				"	 	INNER JOIN [db_item].[dbo].tbl_item AS D ON C.itemid = D.itemid "&_
'				"	 	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS F ON C.itemid = F.itemid "&_
'				"		WHERE evt_code = " &FECode &addSql& strGroup & striSort & " ) " & strGroup&addSql& strSort
		strSql = " SELECT  TOP "&FPSize*FCPage&" A.itemid, A.evtgroup_code, A.evtitem_sort,  B.makerid, B.itemname, B.sellcash "&_
				"		,B.buycash,B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.mileage, B.smallimage, B.listimage,   B.sellyn, B.deliverytype "&_
				"	    ,  B.limityn, B.danjongyn, B.sailyn, B.isusing, B.limitno , B.limitsold, B.itemcouponyn, B.itemcoupontype, B.itemcouponvalue"&_
				"		 , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end as couponbuyprice "&_
				"		, B.mwdiv, A.evtitem_imgsize	"&_
				"	FROM  [db_event].[dbo].[tbl_eventitem] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.evt_code = "&FECode & strGroup&addSql& strSort
		'response.write strSql
		
		rsget.pagesize = FPSize
		rsget.Open strSql,dbget,1
        
        rsget.absolutepage = FCPage
		IF not rsget.EOF THEN
			fnGetEventItem = rsget.getRows()
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
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public function getMwDivName(FmwDiv)
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "특정"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	'## fnGetEventList : 이벤트목록  ##
	public Function fnGetEventList
	Dim strSql, strSqlCnt,iDelCnt, strSearch
	strSearch = ""
	If FSsDate <> ""  or FSeDate <> "" THEN
		if CStr(FSfDate) = "S" THEN
			strSearch  = strSearch & " and  datediff(day, '"&FSsDate&"', evt_startdate) >= 0 and  datediff(day,'"&FSeDate&"', evt_startdate) <=0  "
		elseif CStr(FSfDate) = "E" THEN
			strSearch  = strSearch & " and  datediff(day,'"&FSsDate&"',evt_enddate) >= 0 and  datediff(day,'"&FSeDate&"',evt_enddate) <=0  "
		end if
	END IF
	If FSeTxt <> "" THEN
		IF Cstr(FSfEvt) = "evt_code" THEN
			If chkWord(FSeTxt,"[^-0-9 ]") = "False" Then
				Alert_return("이벤트코드는 숫자만 입력하세요")
				response.end
			End If
			strSearch  = strSearch &  " and A.evt_code = "&FSeTxt
		ELSE
			strSearch  = strSearch &  " and  (A.evt_name like '%"&FSeTxt&"%' or B.workTag like '%"&FSeTxt&"%') "
		END IF
	End If

	If FSstate <> "" THEN
		IF FSstate = 9 THEN	'종료
			strSearch  = strSearch & " and   (evt_state = "&FSstate & " or  datediff(day,getdate(),evt_enddate)< 0 )"
		ELSEIF FSstate = 6 THEN	'오픈예정
			strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0 and datediff(day,getdate(),evt_enddate) >= 0  "
		ELSEIF FSstate = 7 THEN	'오픈진행중
			strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)> 0 and  datediff(day,getdate(),evt_enddate)>=0 "
		ELSE
			strSearch  = strSearch & " and  evt_state = "&FSstate & " and  datediff(day,getdate(),evt_enddate)>=0"
		END IF
	End If
	If FScategory <> "" THEN strSearch  = strSearch &  " and  evt_category = "&FScategory
	If FScateMid <> "" THEN strSearch  = strSearch &  " and  evt_cateMid = "&FScateMid
	If FEDispCate<>"" then	strSearch  = strSearch &  " and  evt_dispcate = "& FEDispCate

	IF FSkind <> "" THEN
		strSearch  = strSearch &  " and evt_kind in ("& FSkind & ") "
	END IF

	IF FSedid <> "" THEN
		strSearch  = strSearch &  " and designerid = '"&FSedid&"'"
	END IF

	IF FSemid <> "" THEN
		strSearch  = strSearch &  " and partMDid = '"&FSemid&"'"
	END IF

	IF FEBrand <> "" THEN
		strSearch  = strSearch & " and brand = '"&FEBrand&"'"
	END IF

	IF FSisSale = "1" THEN strSearch  = strSearch & " and issale = 1 "
	IF FSisGift = "1" THEN strSearch  = strSearch & " and isgift = 1 "
	IF FSisCoupon = "1" THEN strSearch  = strSearch & " and iscoupon = 1 "
	IF FSisOnlyTen = "1" THEN strSearch  = strSearch & " and isOnlyTen = 1 "
	IF FSisDiary = "1" THEN strSearch  = strSearch & " and isDiary = 1 "

	strSqlCnt = " SELECT COUNT(A.evt_code) FROM [db_event].[dbo].[tbl_event] as A "&_
				"   LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code "&_
				" WHERE evt_using ='Y' "&strSearch
	'response.write strSqlCnt
	rsget.Open strSqlCnt,dbget
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close

	'response.write FTotCnt

	IF FTotCnt >0 THEN
		'이벤트 기간 종료시 상태 종료로 , 이벤트 오픈요청상태에서 기간이 진행중일때 상태 오픈으로 view 처리
		iDelCnt =  ((FCPage - 1) * FPSize )+1
		strSql = "SELECT  TOP "&FPSize&" A.evt_code, A.evt_kind, A.evt_manager, A.evt_scope, A.evt_name, A.evt_startdate, A.evt_enddate, A.evt_level "&_
		 		 " ,evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9 "&_
		 		 "				When A.evt_state = 7 and DateDiff(day,getdate(),evt_startdate) <= 0 Then 6 "&_
		 		 "				ELSE	A.evt_state end"&_
				" ,A.evt_regdate,B.evt_bannerimg "&_
				" , (Case When isNull(B.designerid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = B.designerid ) Else '' end) as designername "&_
				" ,(SELECT code_nm from  [db_item].[dbo].tbl_Cate_large WHERE code_large = B.evt_category) categoryname, A.evt_prizedate "&_
				" ,	B.brand, B.issale, B.isgift, B.iscoupon "&_
				" ,	(SELECT COUNT(sale_code) FROM [db_event].[dbo].[tbl_sale] WHERE evt_code = A.evt_code and sale_using =1) as sale_count "&_
				" ,	(SELECT COUNT(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = A.evt_code and gift_using ='y') as gift_count "&_
				" , A.prizeyn "& _
				" , (Case When A.evt_kind=13 Then (Select top 1 itemid from [db_event].[dbo].[tbl_eventitem] where evt_code=A.evt_code) else 0 end) as itemid  "& _
				" ,(select top 1 code_nm from db_item.dbo.tbl_Cate_mid where "& _
				" code_large=b.evt_category and code_mid=b.evt_cateMid) as code_nm "& _
				" , (Case When isNull(B.partMDid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = B.partMDid ) Else '' end) as mdname "&_
				" , isNull(B.evt_bannerimg2010,'') AS evt_bannerimg2010, B.workTag "& _
				" ,(select top 1 catename from db_item.dbo.tbl_display_cate where catecode=left(b.evt_dispcate,3)) as dispcate_nm "& _
				" ,B.evt_itemsort, B.isDiary "& _
				" FROM [db_event].[dbo].[tbl_event] as A LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code "&_
				"	WHERE A.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" A.evt_code FROM [db_event].[dbo].[tbl_event] as A " &_
				"		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B " &_
				"			ON A.evt_code = B.evt_code WHERE evt_using ='Y' " &strSearch&_
				" 		ORDER BY A.evt_code DESC ) as T )  and evt_using ='Y' "&strSearch&" ORDER BY A.evt_code DESC"
		rsget.Open strSql,dbget,0
		'response.write strSql

		IF not rsget.EOF THEN
			fnGetEventList = rsget.getRows()
		End IF
		rsget.Close
	End IF


	End Function

	'## fnGetEventList_LOG : 이벤트목록_물류용  ##
	public Function fnGetEventList_LOG
	Dim strSql, strSqlCnt,iDelCnt, strDate, strState, strCate, strEvt, strKind,strEvtType

	If FSsDate <> ""  or FSeDate <> "" THEN
		'if CStr(FSfDate) = "S" THEN
		'	strDate  = " and  datediff(day, '"&FSsDate&"', evt_startdate) >= 0 and  datediff(day,'"&FSeDate&"', evt_startdate) <=0  "
		'elseif CStr(FSfDate) = "E" THEN
		'	strDate  = " and  datediff(day,'"&FSsDate&"',evt_enddate) >= 0 and  datediff(day,'"&FSeDate&"',evt_enddate) <=0  "
		'end if
		strDate = " and evt_startdate <= convert(varchar(10),dateadd(day,1,'"&FSeDate&"'),121) and evt_enddate >= convert(varchar(10),'"&FSsDate&"',121)"
	END IF

	if FSisSale ="on" then
		strEvtType= strEvtType & " and isSale='1'"
	end if

	if FSisGift ="on" then
		strEvtType = strEvtType & " and isGift='1'"
	end if

	if FSisCoupon ="on" then
		strEvtType= strEvtType & " and isCoupon='1'"
	end if

	If FSeTxt <> "" THEN
		IF Cstr(FSfEvt) = "evt_code" THEN
			strEvt  = " and A.evt_code = "&FSeTxt
		ELSE
			strEvt  = " and  evt_name like '%"&FSeTxt&"%'"
		END IF
	End If

	If FSstate <> "" THEN
		IF FSstate = 9 THEN
			strState =" and   (evt_state = "&FSstate & " or  datediff(day,getdate(),evt_enddate)< 0 )"
		ELSEIF FSstate = 6 THEN	'오픈예정
			strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0 and datediff(day,getdate(),evt_enddate) >= 0  "
		ELSEIF FSstate = 7 THEN	'오픈진행중
			strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)> 0 and  datediff(day,getdate(),evt_enddate)>=0 "
		ELSE
			strState = " and  evt_state = "&FSstate & " and  datediff(day,getdate(),evt_enddate)>=0"
		END IF
	End If
	If FScategory <> "" THEN
		strCate = " and  evt_category = "&FScategory
	END IF
	If FScateMid <> "" THEN
		strCate = " and  evt_cateMid = "&FScateMid
	END IF

	IF FSkind <> "" THEN
		strKind = " and evt_kind = "& FSkind
	END IF

	strSqlCnt = " SELECT COUNT(A.evt_code) FROM [db_event].[dbo].[tbl_event] as A "&_
				"   LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code "&_
				" WHERE evt_using ='Y'  "&strDate&strEvt&strState&strCate&strKind&strEvtType
	rsget.Open strSqlCnt,dbget
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close

	IF FTotCnt >0 THEN
		iDelCnt =  ((FCPage - 1) * FPSize )+1
		strSql = "SELECT  TOP "&FPSize&" A.evt_code, A.evt_kind, A.evt_manager, A.evt_scope, A.evt_name, A.evt_startdate, A.evt_enddate, A.evt_level, "&_
		 		 " evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9 ELSE	A.evt_state end,"&_
				" A.evt_regdate,B.evt_bannerimg, (SELECT company_name from db_partner.[dbo].tbl_partner WHERE id = B.designerid ) designername,  "&_
				" (SELECT code_nm from  [db_item].[dbo].tbl_Cate_large WHERE code_large = B.evt_category) categoryname, A.evt_prizedate "&_
				" FROM [db_event].[dbo].[tbl_event] as A LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code "&_
				"	WHERE A.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" A.evt_code FROM [db_event].[dbo].[tbl_event] as A " &_
				"		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B " &_
				"			ON A.evt_code = B.evt_code WHERE evt_using ='Y' " &strDate&strEvt&strState&strCate&strKind&strEvtType&_
				" 		ORDER BY A.evt_code DESC ) as T ) and evt_using ='Y' "&strDate&strEvt&strState&strCate&strKind&strEvtType&" ORDER BY A.evt_code DESC"

		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetEventList_LOG = rsget.getRows()
		End IF
		rsget.Close
	End IF
	End Function

	public Function fnGetEventLastList
		Dim strSearch, strSqlCnt, iDelCnt, strSql

		If FRectDisp <> "" Then
			strSearch  = strSearch &  " and d.evt_dispCate = '" & FRectDisp & "' "
		END IF
		
		IF FSkind <> "" THEN
		strSearch  = strSearch &  " and e.evt_kind = "& FSkind
		END IF

		If FSeTxt <> "" THEN
			IF Cstr(FSfEvt) = "evt_code" THEN
				strSearch  = strSearch &  " and  e.evt_code = "&FSeTxt
			ELSE
				strSearch  = strSearch &  " and  e.evt_name like '%"&FSeTxt&"%'"
			END IF
		End If
		
		IF FRectState <> "" THEN
		strSearch  = strSearch &  " and e.evt_state = '" & FRectState & "' "
		END IF
		
		IF FRectSDate <> "" THEN
		strSearch  = strSearch &  " and e.evt_startdate >= '" & FRectSDate & "' "
		END IF
		
		IF FRectEDate <> "" THEN
		strSearch  = strSearch &  " and e.evt_enddate <= '" & DateAdd("d",1,FRectEDate) & "' "
		END IF

		strSqlCnt = " SELECT COUNT(e.evt_code) FROM [db_event].[dbo].[tbl_event] as e "&_
				" Left Join [db_event].[dbo].[tbl_event_display] as d on e.evt_code = d.evt_code "&_
				" WHERE e.evt_using ='Y'  "&strSearch
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = "SELECT  TOP "&FPSize&" e.evt_code, e.evt_kind, e.evt_manager, e.evt_scope, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_level "&_
			 		" 		,evt_state = Case When DateDiff(day,getdate(),e.evt_enddate) < 0 Then 9 "&_
		 		 	"					When  e.evt_state = 7 and DateDiff(day,getdate(),e.evt_startdate) <= 0 Then 6 "&_
		 		 	"					ELSE  e.evt_state end"&_
		 		 	"		, isNull(db_item.[dbo].[getDisplayCateName](d.evt_dispCate),'') " & _
					" FROM [db_event].[dbo].[tbl_event] as e  "&_
					" Left Join [db_event].[dbo].[tbl_event_display] as d on e.evt_code = d.evt_code "&_
					" WHERE e.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] " &_
					"			 WHERE evt_using ='Y' " &strSearch&" ORDER BY evt_code DESC ) as T ) "&_
					" and e.evt_using ='Y' "&strSearch&" ORDER BY e.evt_code DESC"
			rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetEventLastList = rsget.getRows()
			End IF
			rsget.Close
		End IF
	End Function

	public Function fnGetEventLastList_v2
		Dim strSearch, strSqlCnt, iDelCnt, strSql

		If FRectDisp <> "" Then
			strSearch  = strSearch &  " and d.evt_dispCate = '" & FRectDisp & "' "
		END IF
		
		IF FSkind <> "" THEN
		strSearch  = strSearch &  " and e.evt_kind = "& FSkind
		END IF

		If FSeTxt <> "" THEN
			IF Cstr(FSfEvt) = "evt_code" THEN
				strSearch  = strSearch &  " and  e.evt_code = "&FSeTxt
			ELSE
				strSearch  = strSearch &  " and  e.evt_name like '%"&FSeTxt&"%'"
			END IF
		End If
		
		IF FRectState <> "" THEN
		strSearch  = strSearch &  " and e.evt_state = '" & FRectState & "' "
		END IF
		
		IF FRectSDate <> "" THEN
		strSearch  = strSearch &  " and e.evt_startdate >= '" & FRectSDate & "' "
		END IF
		
		IF FRectEDate <> "" THEN
		strSearch  = strSearch &  " and e.evt_enddate <= '" & DateAdd("d",1,FRectEDate) & "' "
		END IF

		strSqlCnt = " SELECT COUNT(e.evt_code) FROM [db_event].[dbo].[tbl_event] as e "&_
				" Left Join [db_event].[dbo].[tbl_event_display] as d on e.evt_code = d.evt_code "&_
				" WHERE e.evt_using ='Y'  "&strSearch
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = "SELECT  TOP "&FPSize&" e.evt_code, e.evt_kind, e.evt_manager, e.evt_scope, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_level  "&_
			 		" 		,evt_state = Case When DateDiff(day,getdate(),e.evt_enddate) < 0 Then 9 "&_
		 		 	"					When  e.evt_state = 7 and DateDiff(day,getdate(),e.evt_startdate) <= 0 Then 6 "&_
		 		 	"					ELSE  e.evt_state end"&_
		 		 	"		, isNull(db_item.[dbo].[getDisplayCateName](d.evt_dispCate),'') " & _
					"		, e.evt_subcopyK, d.salePer, d.etc_itemimg " & _					  
					" FROM [db_event].[dbo].[tbl_event] as e  "&_
					" Left Join [db_event].[dbo].[tbl_event_display] as d on e.evt_code = d.evt_code "&_
					" WHERE e.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] " &_
					"			 WHERE evt_using ='Y' " &strSearch&" ORDER BY evt_code DESC ) as T ) "&_
					" and e.evt_using ='Y' "&strSearch&" ORDER BY e.evt_code DESC"
			rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetEventLastList_v2 = rsget.getRows()
			End IF
			rsget.Close
		End IF
	End Function

	'//아이템 복사 리스트(아이템이 포함된 이벤트 리스트)
	public Function fnGetEventLastItemList
		Dim strSearch, strSqlCnt, iDelCnt, strSql

		IF FSkind <> "" THEN
		strSearch  = strSearch &  " and e.evt_kind = "& FSkind
		END IF

		If FSeTxt <> "" THEN
			IF Cstr(FSfEvt) = "evt_code" THEN
				strSearch  = strSearch &  " and  e.evt_code = "&FSeTxt
			ELSE
				strSearch  = strSearch &  " and  e.evt_name like '%"&FSeTxt&"%'"
			END IF
		End If

		strSqlCnt = " SELECT COUNT(e.evt_code) FROM [db_event].[dbo].[tbl_event] as e "&_
				" WHERE e.evt_using ='Y'  "&strSearch
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = "SELECT  TOP "&FPSize&" e.evt_code, e.evt_kind, e.evt_manager, e.evt_scope, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_level  " + vbCrlf
			strSql = strSql + " ,evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9 "+ vbCrlf
		 	strSql = strSql + "	When  evt_state = 7 and DateDiff(day,getdate(),evt_startdate) <= 0 Then 6 "+ vbCrlf
			strSql = strSql + "	ELSE  evt_state end , count(i.evt_code) as itemcnt "+ vbCrlf
			strSql = strSql + " FROM [db_event].[dbo].[tbl_event] as e  "+ vbCrlf
			strSql = strSql + " inner join db_event.dbo.tbl_eventitem as i  "+ vbCrlf
			strSql = strSql + " on e.evt_code = i.evt_code "+ vbCrlf
			strSql = strSql + " WHERE e.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] " + vbCrlf
			strSql = strSql + "			 WHERE evt_using ='Y' " &strSearch&" ORDER BY evt_code DESC ) as T ) "+ vbCrlf
			strSql = strSql + " and evt_using ='Y' "&strSearch&"" + vbCrlf
			strSql = strSql + " group by e.evt_code, e.evt_kind, e.evt_manager, e.evt_scope, e.evt_name, e.evt_startdate , e.evt_enddate, e.evt_level ,evt_state " + vbCrlf
			strSql = strSql + " ORDER BY e.evt_code DESC"
			
			'Response.write strSql

			rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetEventLastItemList = rsget.getRows()
			End IF
			rsget.Close
		End IF
	End Function
End Class


'------------------------------------------------------
'ClsEventPrize : 당첨자
'------------------------------------------------------
Class  ClsEventPrize
	public FECode
	public FEGKindCode
	public FCPage
	public FPSize
	public FTotCnt

	public FEPrizeCode
	public FEPType
	public FEPRanking
	public FEPRankname
	public FEPwinner
	public FEGiftkindCode
	public FEGiftkindName
	public FGiveEPCode
	public FEPTypeDesc

	'## fnGetPrize :당첨자목록 가져오기 ##
	public Function fnGetPrize
	Dim strSql, strSqlAdd,strSqlCnt
	IF FEGKindCode = "" THEN FEGKindCode = 0
	If FEGKindCode > 0 THEN
		strSqlAdd = " and evtgroup_code = "&FEGKindCode
	END IF

	strSqlCnt = " SELECT count(evtprize_code) FROM  [db_event].[dbo].[tbl_event_prize] WHERE evt_code = "&FECode&strSqlAdd
	rsget.Open strSqlCnt,dbget
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close
	IF FTotCnt >0 THEN
		iDelCnt =  (FCPage - 1) * FPSize
		strSql = " SELECT  TOP "&FPSize&" evtprize_code, evt_ranking,evt_rankname, a.itemid, evt_giftname,evt_winner,evt_regdate"&_
				" 		,evtprize_startdate, evtprize_enddate, evtprize_status, a.giftkind_code, " & _
				"		case when a.evtprize_type = '5' then a.evtprize_name else b.giftkind_name end giftkind_name " & _
				"		, b.giftkind_img, b.itemid, evtprize_type,give_evtprizecode "&_
				" FROM  [db_event].[dbo].[tbl_event_prize] a left outer join  [db_event].[dbo].[tbl_giftkind] b  on a.giftkind_code = b.giftkind_code"&_
				"	WHERE evt_code = "&FECode&strSqlAdd&" AND evtprize_code not in ( SELECT TOP "&iDelCnt&" evtprize_code FROM  [db_event].[dbo].[tbl_event_prize] "&_
				"			WHERE evt_code = "&FECode&strSqlAdd&" ORDER BY evt_ranking, evtprize_code desc ) " &_
				" ORDER BY evt_ranking, evtprize_code desc "
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetPrize = rsget.getRows()
			END IF
		rsget.Close
	END IF
	End Function

	public Function fnGetPrizeConts
		Dim strSql
		strSql =" SELECT  evtprize_code, evt_code, evtgroup_code, evt_ranking,evt_rankname,evt_winner,evt_regdate "&_
			",evtprize_startdate, evtprize_enddate, evtprize_status, a.giftkind_code, b.giftkind_name, b.giftkind_img, b.itemid, evtprize_type, give_evtprizecode"&_
			",(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'evtprizetype' and code_value = a.evtprize_type) evtprize_typedesc"&_
			" FROM  [db_event].[dbo].[tbl_event_prize] a left outer join  [db_event].[dbo].[tbl_giftkind] b  on a.giftkind_code = b.giftkind_code"&_
			"	WHERE a.evtprize_code = "&FEPrizeCode
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FECode			= rsget("evt_code")
				FEGKindCode	 	= rsget("evtgroup_code")
				FEPType 		= rsget("evtprize_type")
				FEPTypeDesc		= rsget("evtprize_typedesc")
			 	FEPRanking 		= rsget("evt_ranking")
			 	FEPRankname 	= rsget("evt_rankname")
				FEPwinner 		= rsget("evt_winner")
			  	FEGiftkindCode 	= rsget("giftkind_code")
			  	FEGiftkindName 	= rsget("giftkind_name")
			  	FGiveEPCode		= rsget("give_evtprizecode")

			END IF
		rsget.Close
	End Function

	public Function fnGetPrizeListUserInfo
		Dim strSql, strSqlAdd,strSqlCnt
		IF FEGKindCode = "" THEN FEGKindCode = 0
		If FEGKindCode > 0 THEN
			strSqlAdd = " and evtgroup_code = "&FEGKindCode
		END IF


			strSql = " SELECT u.usercell, u.userid" + vbcrlf
			strSql = strSql + " FROM  [db_event].[dbo].[tbl_event_prize] a left join  [db_user].[dbo].[tbl_user_n] u  on a.evt_winner = u.userid" + vbcrlf
			strSql = strSql + " WHERE a.evtprize_code in (" & FECode & ")" & strSqlAdd + vbcrlf
			strSql = strSql + " ORDER BY a.evtprize_code asc"
			'Response.write strSql
			rsget.Open strSql,dbget
				IF not rsget.EOF THEN
					fnGetPrizeListUserInfo = rsget.getRows()
				END IF
			rsget.Close
	End Function

End Class

'-------------------------------------------------------------
'ClsEventSchedule : 이벤트 스케쥴
'-------------------------------------------------------------
Class ClsEventSchedule
	public FFDate
	public FLDate

	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt

	public FSCategory
	public FSState

	public Function fnGetList
		Dim strSql, strSqlCnt,iDelCnt, strC, strE

		IF FSCategory <> "" THEN
			IF FSCategory = "-1" THEN
				strC =  " and B.evt_category = ''"
			ELSE
				strC =  " and B.evt_category = "&FSCategory
			END IF
		END IF

		IF FSState = "-1" THEN
			strE = " AND DateDiff(day,getdate(),evt_enddate) >= 0 AND  A.evt_state < 9 "
		ELSEIF FSState ="7" THEN
		 	strE = " AND DateDiff(day,getdate(),evt_startdate) > 0 AND  A.evt_state = 7  "
		ELSEIF FSState ="6" THEN
			strE = " AND DateDiff(day,getdate(),evt_startdate) <= 0 AND DateDiff(day,getdate(),evt_enddate) >= 0  AND  A.evt_state = 7  "
		ELSEIF FSState ="9" THEN
			strE = " AND (DateDiff(day,getdate(),evt_enddate) < 0  OR  A.evt_state = 9) "
		ELSE
			strE = " AND A.evt_state = "&FSState&" AND  DateDiff(day,getdate(),evt_enddate) >= 0"
		END IF

		strSqlCnt = " SELECT COUNT(A.evt_code) FROM [db_event].[dbo].[tbl_event] as A "&_
					" 	LEFT OUTER  JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code "&_
					" where ((DateDiff(day,'"&FFDate&"' ,evt_startdate) >= 0 and DateDiff(day,'"&FLDate&"',evt_startdate) <=0  ) "&_
	 				" 		or (DateDiff(day,'"&FFDate&"',evt_enddate) >=0  and DateDiff(day,'"&FLDate&"' ,evt_enddate) <= 0))  "&strC&strE
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " select  TOP "&FPSize&" A.evt_code, A.evt_kind, A.evt_manager, A.evt_scope, A.evt_name, A.evt_level, "&_
					" evt_state = Case When DateDiff(day,getdate(), A.evt_enddate) < 0 Then 9 ELSE	A.evt_state end,"&_
					" A.evt_startdate, A.evt_enddate, (select top 1 isNull(db_item.dbo.getCateCodeFullDepthName(d.catecode),'') as catename from db_item.dbo.tbl_display_cate as d where d.catecode=left(B.evt_dispCate,3)) AS categoryname, M.evt_copy_code"&_
					" from  [db_event].[dbo].[tbl_event] as A "&_
					" 		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code " &_
					"		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_md_theme] as M ON A.evt_code = M.evt_code " &_
	 				" where( (DateDiff(day,'"&FFDate&"' ,evt_startdate) >= 0 and DateDiff(day,'"&FLDate&"',evt_startdate) <=0  ) "&_
	 				" 		or (DateDiff(day,'"&FFDate&"',evt_enddate) >=0  and DateDiff(day,'"&FLDate&"' ,evt_enddate) <= 0 ) )"&_
	 				"	   and  A.evt_code  <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] "&_
	 				" 		where ((DateDiff(day,'"&FFDate&"' ,evt_startdate) >= 0  and DateDiff(day,'"&FLDate&"',evt_startdate) <=0  ) "&_
	 				" 		or (DateDiff(day,'"&FFDate&"',evt_enddate) >=0  and DateDiff(day,'"&FLDate&"' ,evt_enddate) <= 0) )"&strC &strE&"  ORDER BY evt_code DESC ) as T )"&strC&strE&_
	 				" ORDER BY A.evt_code DESC "
	 		rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetList = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function
End Class

'-------------------------------------------------------------
'ClsEventGroup : 이벤트 그룹
'-------------------------------------------------------------
Class ClsEventGroup
	public FECode
	public FEGCode

	public FGDesc
	public FGSort
	public FGImg
	public FGPCode
	public FGDepth
	public FGPDesc
	public FGlink
	public FRegdate
    public FRectGroupDelInc
    
	public Function fnGetRootGroup
		Dim strSql
		strSql = " SELECT evtgroup_code, evtgroup_desc FROM [db_event].[dbo].tbl_eventitem_group "&_
				" WHERE evt_code = "&FECode&" and evtgroup_pcode = 0 and evtgroup_using ='Y' "
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetRootGroup = rsget.getRows()
			End IF
			rsget.Close
	End Function

	' /admin/shopmaster/gift/giftuserdetail.asp
	'## fnGetEventItemGroup :이벤트화면설정 그룹내용가져오기 ##
	public Function fnGetEventItemGroup
		Dim strSql

		IF FECode = "" THEN Exit Function
	
		if (FRectGroupDelInc="Y") then ''삭제된 그룹도 표시 2015/05/26
			strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
					"		(select evtgroup_desc from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode), evtgroup_using "&_
					" FROM  [db_event].[dbo].[tbl_eventitem_group] as a" &_
					"	WHERE evt_code = "&FECode&"  ORDER BY evtgroup_depth, evtgroup_sort "
		else
			strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
					"		(select evtgroup_desc from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode), evtgroup_using "&_
					" FROM  [db_event].[dbo].[tbl_eventitem_group] as a" &_
					"	WHERE evt_code = "&FECode&" and evtgroup_using ='Y' ORDER BY evtgroup_depth, evtgroup_sort "
		end if

		'response.write strSql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			fnGetEventItemGroup = rsget.getRows()
		End IF
		rsget.Close

		rsget.CursorLocation = adUseClient
		rsget.Open "SELECT Year(evt_regdate) FROM [db_event].[dbo].[tbl_event] WHERE evt_code = '" & FECode & "'",dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FRegdate = rsget(0)
		End IF
		rsget.Close
	End Function

	public Function fnGetEventItemGroupCont
	Dim strSql
	IF FEGCode = "" THEN Exit Function
	strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
			"		isnull((select evtgroup_desc from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode),'최상위') as evtgroup_pdesc"&_
			"	FROM  [db_event].[dbo].[tbl_eventitem_group] as a " &_
			"	WHERE evt_code = "&FECode&" and evtgroup_code="&FEGCode&" and evtgroup_using ='Y' "
	rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FGDesc = rsget("evtgroup_desc")
			FGSort = rsget("evtgroup_sort")
			FGImg  = rsget("evtgroup_img")
			FGPCode= rsget("evtgroup_pcode")
			FGDepth= rsget("evtgroup_depth")
			FGPDesc= rsget("evtgroup_pdesc")
			FGlink= rsget("evtgroup_link")
		End IF
		rsget.Close
	End Function
End Class

'-------------------------------------------------------------
'ClsEventSummary : 이벤트 요약 내용 - 사은품, 할인, 쿠폰에 연계 되는 간략한 내용
'-------------------------------------------------------------
Class ClsEventSummary
	public FECode
	public FEName
	public FESDay
	public FEEDay
	public FEState
	public FBrand
	public FEOpenDate
	public FEStateDesc
	public FECloseDate
	public FEScope
	public FPartnerID

	public Function fnGetEventConts
	 Dim strSql
	 strSql = " SELECT  evt_name, evt_startdate, evt_enddate, evt_state, brand, opendate, closedate, evt_scope, partner_id "&_
	 		",(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'eventstate' and code_value = A.evt_state) evt_statedesc"&_
	 		" FROM [db_event].[dbo].[tbl_event] as A inner join [db_event].[dbo].[tbl_event_display] as B on A.evt_code = B.evt_code "&_
	 		" WHERE A.evt_code = "&FECode
	 rsget.Open strSql,dbget
	 IF not rsget.EOF THEN
	 	 FEName 	= db2html(rsget("evt_name"))
	 	 FESDay 	= rsget("evt_startdate")
	 	 FEEDay 	= rsget("evt_enddate")
	 	 FEState 	= rsget("evt_state")
	 	 FEStateDesc= fnSetStatusDesc(FEState,FESDay,FEEDay, rsget("evt_statedesc"))
	 	 'IF datediff("d",FEEDay,now) > 0  THEN FEState = 9	'종료일이 지난 경우 종료로 표기
	 	 FBrand 	= db2html(rsget("brand"))
	 	 FEOpenDate = rsget("opendate")
	 	 FECloseDate= rsget("closedate")
	 	 FEScope	= rsget("evt_scope")
	 	 FPartnerID	= rsget("partner_id")
	 END IF
	 rsget.close
	End Function
End Class
%>