<%
'####################################################
' Description :  이벤트 관리
' History : 2022.05.06 V6 생성 : 이벤트 그룹 정렬 순서 변경 (김형태) - 전시여부에따른 정렬 기준 삭제
'			2022.00.00 홍길동 수정
' /event/eventmanage/common/event_function.asp include 필수!
'####################################################

'------------------------------------------------------
'ClsEvent : 이벤트 내용
'------------------------------------------------------
Class ClsEvent
	public FECode	'해당 이벤트코드
	public FECodeArr
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
	public FEDgId
	public FEMdId
	public FEMID
	public FEFwd
	public FEFwdMO
	public FChkDisp
	public FEBrand
	public FEIcon
	public FECommentTitle
	public FELinkCode

	public FELinkType
	public FELinkURL

	public FEType		'// 이벤트 유형
	public FisConfirm	'// 상급자 이벤트 내용 확인 여부

	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt
    public FDispYCnt
    public FDispNCnt
    public FDispYMCnt
    public FDispNMCnt
	public FESGroup	'Set 그룹검색
	public FESSort	'Set 정렬
	public FRectDispCate

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
	public FSedid2
	public FSemid

	public FchComm
	public FchBbs
	public FchItemps
	public Fisblogurl

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
	public FRectItemidArr
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
	public FRectEvtType
	public FRectEvtManager '이벤트 등록주체(1-10x10, 2-업체)
	public FRectIsConfirm

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
	public Fnocate
	public Ftitle_pc
	public Ftitle_mo
	public Feval_isusing
	public Feval_text
	public Feval_freebie_img
	public Feval_start
	public Feval_end
	public Fboard_isusing
	public Fboard_text
	public Fboard_freebie_img
	public Fboard_start
	public Fboard_end
	public FBrandName
	public FBrandContents
	public FGroupItemPriceView
	public FGroupItemCheck
	public FGroupItemType
	public FPrizeYN
	public FEvtCopyCode
	public FEvtImgCopyUserid
	public FEvtImgCopyCode

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

	Public FEtcitemid '대표상품코드
	Public FEtcitemimg '대표상품이미지
	Public FEsortNo		'정렬번호(회차)
	Public FEdateview
	Public FEitemid

	Public FENamesub
	Public FEListType

	Public FIsWeb
	public FIsMobile
	public FIsApp

	public FSort

	public FEPSId
	public FEDPId
	public FEDId
	public FECCId
	public FEDgName
	public FEMdName
	public FEPsName
	public FEDpName
	public FECCName

	public FEDgId2
	public FEDgName2
	public FEDgStat1
	public FEDgStat2

	public FEMImg_mo
	public FEHtml_mo

	public FSepsid
	public FSedpid

	public FSednm
	public FSemnm
	public FSepsnm
	public FSedpnm

	public FSisoneplusone
	public FSisfreedelivery
	public FSisbookingsell
	public FSisNew
    public FisReqPublish
	public FRectOnlyMobile

	public FEisExec
    public FEexecFile
    public FEisExec_mo
    public FEexecFile_mo
    public FETemp_mo

    public FEChannel
    public FEImgRegdate

	Public FEsgroup_W  '// 이벤트 그룹형 최상위 랜덤노출 pc 웹
	Public FEsgroup_M  '// 이벤트 그룹형 최상위 랜덤노출 모바일

	Public FESlide_W_Flag '// 웹슬라이드 사용유무
	Public FESlide_M_Flag '// 모바일슬라이드 사용유무
	Public FEvt_pc_addimg_cnt '// PC 상중하단 추가 이미지 카운트
	Public FEvt_m_addimg_cnt '// 모바일 상중하단 추가 이미지 카운트

	Public Fmdtheme
	Public Fmdthememo
	Public Fthemecolor
	Public Fthemecolormo
	Public Ftextbgcolor
	Public Ftextbgcolormo
	Public Fmdbntype
	Public Fmdbntypemo
	Public FsalePer
	Public FsaleCPer
	Public FendlessView
	Public Feventtype_pc
	Public Feventtype_mo

	Public Fcomm_isusing
	Public Fcomm_text
	Public Ffreebie_img
	Public Fcomm_start
	Public Fcomm_end
	Public Fgift_isusing
	Public Fgift_text1
	Public Fgift_img1
	Public Fgift_text2
	Public Fgift_img2
	Public Fgift_text3
	Public Fgift_img3
	Public Fusinginfo
	Public Fusing_text1
	Public Fusing_contents1
	Public Fusing_text2
	Public Fusing_contents2
	Public Fusing_text3
	Public Fusing_contents3
	Public FRectMDTheme
	Public FbannerTypeDiv
	Public FbannerCouponTxt
	Public FbannerGubun
	public FDispCateGroup
	Public FvideoLink
	Public FRectEventType_PC
	Public FRectEventType_MO
	Public FvideoType
	public FcontentsAlign
	public FestimateSalePrice
	public FRectStartESP
	public FRectEndESP
	public FRectEvtLevel
	public FRectendlessView
	public Fmarketing_event_kind
	public Fkakao_title

	'## fnGetEventCont : 이벤트개요 내용 가져오기 ##
	public Function fnGetEventCont
	Dim strSql
	IF FECode = "" THEN Exit Function
		strSql = " SELECT  evt_kind, evt_manager, evt_scope, evt_name, evt_startdate, evt_enddate, evt_prizedate, evt_level, evt_state, evt_regdate, evt_using, opendate, closedate,partner_id "&_
				",(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'eventkind' and code_value = a.evt_kind) evt_kinddesc "&_
				",(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'eventstate' and code_value = a.evt_state) evt_statedesc, a.prizeyn "&_
				",evt_nameEng , evt_subcopyK , evt_subcopyE,evt_sortNo , evt_subname "&_
				" , (Case When A.evt_kind=13 Then (Select top 1 itemid from [db_event].[dbo].[tbl_eventitem] where evt_code=A.evt_code) else 0 end) as itemid  "& _
				" , isWeb, isMobile, isApp, evt_imgregdate, evt_type, isConfirm "&_
				" FROM [db_event].[dbo].[tbl_event] a "&_
				" WHERE evt_code = "&FECode
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

			FIsWeb		= rsget("isWeb")
			FIsMobile	= rsget("isMobile")
			FIsApp		= rsget("isApp")

			FEImgRegdate = rsget("evt_imgregdate")
			FEType		= rsget("evt_type")
			FisConfirm	= rsget("isConfirm")
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
				" 	,isoneplusone , isfreedelivery , etc_itemid , etc_itemimg , isbookingsell, evt_dateview , evt_todaybanner , evt_mo_listbanner, evt_itemlisttype" &_
				"	,publisherid,developerid,tdg.username as dgName, tmd.username as mdName, tps.username as psName, tdp.username as dpName "&_
				"	,evt_mainimg_mo, evt_html_mo, isDiary, evt_forward_mo, isNew, isReqPublish, evt_isExec,evt_execFile, evt_isExec_mo, evt_execFile_mo	, evt_template_mo ,  evt_sgroup_w , evt_sgroup_m "&_
				"	, evt_slide_w_flag , evt_slide_m_flag , evt_pc_addimg_cnt, evt_m_addimg_cnt, codecheckerid, tcc.username as CCName "&_
				"	, isNull(dsn_state1,'') as dsn_state1 , isNull(dsn_state2,'') as dsn_state2 , isNull(designerid2,'') as designerid2, tdg2.username as dgName2 "&_
				"	, mdtheme, mdthememo, themecolor, themecolormo, textbgcolor, textbgcolormo, mdbntype, mdbntypemo, salePer"&_
				", saleCPer, endlessView, eventtype_pc, eventtype_mo, bannerTypeDiv, bannerCouponTxt, bannerGubun, videoLink, videoType, estimateSalePrice, marketing_event_kind, ess.kakao_title"&_
				" FROM [db_event].[dbo].[tbl_event_display] as ed  "&_
				" LEFT OUTER JOIN [db_event].[dbo].[tbl_event_share_sns] as ess ON ed.evt_code = ess.evt_code "&_
				"			Left OUter Join db_partner.dbo.tbl_user_tenbyten as tdg on ed.designerid = tdg.userid  and  ed.designerid is not null and  ed.designerid  <> '' "&_
				"			Left OUter Join db_partner.dbo.tbl_user_tenbyten as tdg2 on ed.designerid2 = tdg2.userid  and  ed.designerid2 is not null and  ed.designerid2  <> '' "&_
				"			Left OUter Join db_partner.dbo.tbl_user_tenbyten as tmd on ed.partMDid = tmd.userid and  ed.partMDid is not null and  ed.partMDid  <> '' "&_
				"			Left OUter Join db_partner.dbo.tbl_user_tenbyten as tps on ed.publisherid = tps.userid  and ed.publisherid is not null and  ed.publisherid  <> '' "&_
				"			Left OUter Join db_partner.dbo.tbl_user_tenbyten as tdp on ed.developerid = tdp.userid and ed.developerid is not null and  ed.developerid  <> '' "&_
				"			Left OUter Join db_partner.dbo.tbl_user_tenbyten as tcc on ed.codecheckerid = tcc.userid and ed.codecheckerid is not null and  ed.codecheckerid  <> '' "&_
				" WHERE ed.evt_code = "&FECode
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
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
			FEGImg = rsget("evt_giftimg")
			FETemp = rsget("evt_template")
			FEMImg = rsget("evt_mainimg")
			FEHtml = rsget("evt_html")
			FEISort = rsget("evt_itemsort")
			FEDgId = rsget("designerid")
			FEDid = FEDgId
			FEMdId = rsget("partMDid")
			FEMID	= FEMdId
			FEPSId = rsget("publisherid")
			FEDPId = rsget("developerid")
			FECCId = rsget("codecheckerid")
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
			FSisDiary = rsget("isDiary") '// 다이어리 상태값 추가
			FSisNew		= rsget("isNew")
			FWorkTag	= rsget("workTag")
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

			FEDgName = rsget("dgName")
			FEMdName = rsget("MdName")
			FEPsName = rsget("PsName")
			FEDpName = rsget("DpName")
			FECCName = rsget("CCName")

			FEDgId2 = rsget("designerid2")
			FEDgName2 = rsget("dgName2")
			FEDgStat1 = rsget("dsn_state1")
			FEDgStat2 = rsget("dsn_state2")

			FisReqPublish = rsget("isReqPublish")
			FEisExec    = rsget("evt_isExec")
			FEexecFile  = rsget("evt_execFile")

			FEBImgMobile    = rsget("evt_bannerimg_mo")
			FEFwdMO         = rsget("evt_forward_mo")
			FEMImg_mo       = rsget("evt_mainimg_mo")
			FEHtml_mo       = rsget("evt_html_mo")
			FEisExec_mo     = rsget("evt_isExec_mo")
			FEexecFile_mo   = rsget("evt_execFile_mo")
			FETemp_mo       = rsget("evt_template_mo")

			FEsgroup_W       = rsget("evt_sgroup_w")
			FEsgroup_M       = rsget("evt_sgroup_m")

			FESlide_W_Flag       = rsget("evt_slide_w_flag")
			FESlide_M_Flag       = rsget("evt_slide_m_flag")

			FEvt_pc_addimg_cnt       = rsget("evt_pc_addimg_cnt")
			FEvt_m_addimg_cnt       = rsget("evt_m_addimg_cnt")

			Fmdtheme       = rsget("mdtheme")
			Fmdthememo       = rsget("mdthememo")
			Fthemecolor       = rsget("themecolor")
			Fthemecolormo       = rsget("themecolormo")
			Ftextbgcolor       = rsget("textbgcolor")
			Ftextbgcolormo       = rsget("textbgcolormo")
			Fmdbntype       = rsget("mdbntype")
			Fmdbntypemo       = rsget("mdbntypemo")
			FsalePer      = rsget("salePer")
			FsaleCPer      = rsget("saleCPer")
			FendlessView      = rsget("endlessView")
			Feventtype_pc      = rsget("eventtype_pc")
			Feventtype_mo      = rsget("eventtype_mo")
			FbannerTypeDiv      = rsget("bannerTypeDiv")
			FbannerCouponTxt      = rsget("bannerCouponTxt")
			FbannerGubun      = rsget("bannerGubun")
			FvideoLink      = rsget("videoLink")
			FvideoType      = rsget("videoType")
			FestimateSalePrice      = rsget("estimateSalePrice")
			Fmarketing_event_kind = rsget("marketing_event_kind")
			Fmarketing_event_kind = rsget("marketing_event_kind")
			Fmarketing_event_kind = rsget("marketing_event_kind")
			Fkakao_title = rsget("kakao_title")
		End IF
		rsget.Close
	End Function

'## fnGetEventMDThemeInfo :이벤트 엠디 등록 이벤트 테마 정보 ##
	public Function fnGetEventMDThemeInfo
	Dim strSql
	IF FECode = "" THEN Exit Function
		strSql = " SELECT  comm_isusing, comm_text, freebie_img, comm_start, comm_end, gift_isusing, gift_text1, gift_img1"&_
				" , gift_text2, gift_img2, gift_text3, gift_img3, usinginfo, using_text1, using_contents1, using_text2"&_
				" , using_contents2, using_text3, using_contents3, nocate, title_pc, title_mo, eval_isusing ,eval_text" &_
				" ,eval_freebie_img ,eval_start ,eval_end, BrandName, BrandContents, GroupItemPriceView, GroupItemCheck, GroupItemType" &_
				" , board_isusing, board_text, board_freebie_img, board_start, board_end, contentsAlign, isnull(evt_copy_code,0) as evt_copy_code, isnull(evt_imgCopy_userid,'') as evt_imgCopy_userid, isnull(evt_imgCopy_code,0) as evt_imgCopy_code" &_
				" FROM [db_event].[dbo].[tbl_event_md_theme]"&_
				" WHERE evt_code = "&FECode
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			Fcomm_isusing = rsget("comm_isusing")
			Fcomm_text = rsget("comm_text")
			Ffreebie_img = rsget("freebie_img")
			Fcomm_start = rsget("comm_start")
			Fcomm_end = rsget("comm_end")
			Fgift_isusing = rsget("gift_isusing")
			Fgift_text1 = rsget("gift_text1")
			Fgift_img1 = rsget("gift_img1")
			Fgift_text2 = rsget("gift_text2")
			Fgift_img2 = rsget("gift_img2")
			Fgift_text3 = rsget("gift_text3")
			Fgift_img3 = rsget("gift_img3")
			Fusinginfo = rsget("usinginfo")
			Fusing_text1 = rsget("using_text1")
			Fusing_contents1 = rsget("using_contents1")
			Fusing_text2 = rsget("using_text2")
			Fusing_contents2 = rsget("using_contents2")
			Fusing_text3 = rsget("using_text3")
			Fusing_contents3 = rsget("using_contents3")
			Fnocate = rsget("nocate")
			Ftitle_pc = rsget("title_pc")
			Ftitle_mo = rsget("title_mo")
			Feval_isusing = rsget("eval_isusing")
			Feval_text = rsget("eval_text")
			Feval_freebie_img = rsget("eval_freebie_img")
			Feval_start = rsget("eval_start")
			Feval_end = rsget("eval_end")
			FBrandName = rsget("BrandName")
			FBrandContents = rsget("BrandContents")
			FGroupItemPriceView = rsget("GroupItemPriceView")
			FGroupItemCheck = rsget("GroupItemCheck")
			FGroupItemType = rsget("GroupItemType")
			Fboard_isusing = rsget("board_isusing")
			Fboard_text = rsget("board_text")
			Fboard_freebie_img = rsget("board_freebie_img")
			Fboard_start = rsget("board_start")
			Fboard_end = rsget("board_end")
			FEvtCopyCode = rsget("evt_copy_code")
			FEvtImgCopyUserid = rsget("evt_imgCopy_userid")
			FEvtImgCopyCode = rsget("evt_imgCopy_code")
			FcontentsAlign = rsget("contentsAlign")
		End IF
		rsget.Close
	End Function

	'## fnGetEventTextTitle :이벤트 텍스트 타이틀  가져오기 ##
	public Function fnGetEventTextTitle
		Dim strSql
		strSql = " SELECT ttCode, ttType, MainTitle, subTitle FROM db_event.dbo.tbl_event_TextTitle WHERE evt_code ="+ FECode +"  and isusing = 1 order by ttType"
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetEventTextTitle = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 기획전목록 정보
	public Function fnGetMailzineEventListData
		Dim strSql, i
		dim strEvtCodeList, arrEvtCodeList
		strEvtCodeList = Replace(FECodeArr, vbCrLf, ",")
		arrEvtCodeList = Split(FECodeArr, vbCrLf)
		strSql = " select e.evt_code, d.etc_itemimg, d.evt_mo_listbanner, replace(e.evt_name, '|~' + convert(varchar,d.salePer) + '%', '') as evt_name, e.evt_subcopyK, (case when d.issale = 1 then d.salePer else '' end), (case when d.iscoupon = 1 then d.saleCPer else '' end)  "
		strSql = strSql + " from "
		strSql = strSql + " 	[db_event].[dbo].[tbl_event] e "
		strSql = strSql + " 	left join [db_event].[dbo].[tbl_event_display] d "
		strSql = strSql + " 	on "
		strSql = strSql + " 		e.evt_code = d.evt_code "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + " 	and e.evt_code in (" & strEvtCodeList & ") "
		strSql = strSql + "  "
		strSql = strSql + " order by "
		strSql = strSql + " 	(case "
		for i = 0 to UBound(arrEvtCodeList)
			strSql = strSql + " 		when e.evt_code = " & arrEvtCodeList(i) & " then " & i
		next
		strSql = strSql + " 		else 10000 end) "
		
		'response.write strSql & "<br>"
		'response.end
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzineEventListData = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 다이어리 정보
	public Function fnGetMailzinediaryData
		Dim strSql, i

		dim strItemidList
		strItemidList = Replace(FRectItemidArr, vbCrLf, ",")

		strSql = " select top 24 "
		strSql = strSql + " 	i.itemid, '' as photoimg, i.icon1image "
 		strSql = strSql + " 	, ('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=' + convert(nvarchar,i.itemid)) as linkinfo, i.itemname "
 		strSql = strSql + " 	, i.orgprice "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sellcash "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then ss.saleprice "
 		strSql = strSql + " 			else i.orgprice end) as sellcash "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sailyn "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then 'Y' "
 		strSql = strSql + " 			else 'N' end) as sailyn "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then FLOOR(100.0-100.0*i.sailprice/i.orgprice) "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then FLOOR(100.0-100.0*ss.saleprice/i.orgprice) "
 		strSql = strSql + " 			else 0.0 end) as salePer "
 		strSql = strSql + " 	, (case when DateDiff(d, i.regdate, getdate()) < 14 then 'Y' else 'N' end) as isNew, i.itemdiv "
 		strSql = strSql + " 	, (case when cs.itemcoupontype = 1 then cs.itemcouponvalue else 0 end) as itemcoupon "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when cs.itemcoupontype = 1 and ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then FLOOR((100-cs.itemcouponvalue)*(i.sellcash)/100) "
 		strSql = strSql + " 			when cs.itemcoupontype = 1 and ss.saleprice is not NULL and i.orgprice > ss.saleprice then FLOOR((100-cs.itemcouponvalue)*(ss.saleprice)/100) "
 		strSql = strSql + " 			when cs.itemcoupontype = 1 then FLOOR((100-cs.itemcouponvalue)*(i.orgprice)/100) "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sellcash "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then ss.saleprice "
 		strSql = strSql + " 			else i.orgprice end) as itemcouponprice "
 		strSql = strSql + " 	, i.optioncnt "
 		strSql = strSql + " 	, ss.saleprice as startsaleprice "
 		strSql = strSql + " 	, se.itemid as saleenditemid "
 		strSql = strSql + " 	, cs.itemcoupontype, cs.itemcouponvalue "
 		strSql = strSql + " 	, ce.itemid as couponenditemid "
 		strSql = strSql + " from [db_item].[dbo].tbl_item i"
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 12 "
 		strSql = strSql + " 		d.itemid, min(d.saleprice) as saleprice "
 		strSql = strSql + " 	from db_event.dbo.tbl_sale as m with (noLock) "
 		strSql = strSql + " 		join db_event.dbo.tbl_saleitem as d with (noLock) "
 		strSql = strSql + " 			on m.sale_code=d.sale_code "
 		strSql = strSql + " 	where m.sale_status in (6,7) "
 		strSql = strSql + " 		and d.saleItem_status in (6,7) "
 		strSql = strSql + " 		and sale_using=1 "
 		strSql = strSql + " 		and '" & Replace(FESDay, ".", "-") & "' between sale_startdate and sale_enddate "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) ss on ss.itemid = i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 12 "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " 	from db_event.dbo.tbl_sale as m with (noLock) "
 		strSql = strSql + " 		join db_event.dbo.tbl_saleitem as d with (noLock) "
 		strSql = strSql + " 			on m.sale_code=d.sale_code "
 		strSql = strSql + " 	where m.sale_status in (6,9) "
 		strSql = strSql + " 		and d.saleItem_status in (6,9) "
 		strSql = strSql + " 		and sale_using=1 "
 		strSql = strSql + " 		and datediff(day,sale_enddate,'" & Replace(FESDay, ".", "-") & "')=1 "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) se on se.itemid = i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 12 d.itemid, max(m.itemcoupontype) as itemcoupontype, min(m.itemcouponvalue) as itemcouponvalue "
 		strSql = strSql + " 	from "
 		strSql = strSql + " 		[db_item].[dbo].[tbl_item_coupon_master] m "
 		strSql = strSql + " 		join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx "
 		strSql = strSql + " 	where	m.openstate in (6,7) "
 		strSql = strSql + " 		and '" & Replace(FESDay, ".", "-") & "' between m.itemcouponstartdate and itemcouponexpiredate "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
 		strSql = strSql + " 		and m.itemcoupontype = 1 "
		strSql = strSql + " 		and m.coupongubun = 'C' "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) cs on cs.itemid = i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 12 d.itemid "
 		strSql = strSql + " 	from "
 		strSql = strSql + " 		[db_item].[dbo].[tbl_item_coupon_master] m "
 		strSql = strSql + " 		join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx "
 		strSql = strSql + " 	where	m.openstate in (6,9) "
 		strSql = strSql + " 		and datediff(day,itemcouponexpiredate,'" & Replace(FESDay, ".", "-") & "')=1 "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
		strSql = strSql + " 		and m.coupongubun = 'C' "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) ce on ce.itemid = i.itemid "
 		strSql = strSql + " where 1=1"
 		strSql = strSql + " and i.itemid in (" & strItemidList & ") "
 		strSql = strSql + " order by i.itemid desc"

		'response.write strSql & "<br>"
		'response.end
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzinediaryData = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 엠디추천 정보
	public Function fnGetMailzineMDPickData
		Dim strSql, i

		dim strItemidList
		strItemidList = Replace(FRectItemidArr, vbCrLf, ",")

		''strSql = " select top 12 "
		''strSql = strSql + " 	f.linkitemid, 'http://imgstatic.10x10.co.kr/contents/maincontents/' + f.photoimg as photoimg, i.icon1image "
		''strSql = strSql + " 	, 'http://www.10x10.co.kr' + f.linkinfo as linkinfo, f.textinfo, i.orgprice, i.sellcash "
		''strSql = strSql + " 	, i.sailyn, Round(100.0-100.0*i.sailprice/i.orgprice,0) as salePer "
		''strSql = strSql + " 	, (case when DateDiff(d, i.regdate, getdate()) < 14 then 'Y' else 'N' end) as isNew, i.itemdiv, (case when i.itemcouponyn = 'Y' and i.itemcoupontype = 1 then i.itemcouponvalue else 0 end) as itemcoupon "
		''strSql = strSql + "		, (case when i.itemcouponyn = 'Y' and i.itemcoupontype = 1 then round((100-i.itemcouponvalue)*(i.sellcash)/100,0) else i.sellcash end) as itemcouponprice, i.optioncnt "
		''strSql = strSql + " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash f "
		''strSql = strSql + " left join [db_item].[dbo].tbl_item i on f.linkitemid=i.itemid "
		''strSql = strSql + " where 1=1 and f.isusing in ('Y','M') "
		''strSql = strSql + " and f.linkitemid in (" & strItemidList & ") "
		''strSql = strSql + " and f.startdate <= '" & FESDay & "' "
		''strSql = strSql + " and f.enddate >= '" & FESDay & "' "
		''strSql = strSql + " and f.linkitemid is not NULL "
		''strSql = strSql + " order by f.disporder, f.startdate desc, f.idx desc "

		strSql = " select top 15 " & vbcrlf
		strSql = strSql + " 	f.linkitemid, 'http://imgstatic.10x10.co.kr/contents/maincontents/' + f.photoimg as photoimg, i.icon1image " & vbcrlf
 		strSql = strSql + " 	, 'http://www.10x10.co.kr' + f.linkinfo as linkinfo, f.textinfo " & vbcrlf
 		strSql = strSql + " 	, i.orgprice " & vbcrlf
 		strSql = strSql + " 	, (case " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sellcash " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then ss.saleprice " & vbcrlf
 		strSql = strSql + " 			else i.orgprice end) as sellcash " & vbcrlf
 		strSql = strSql + " 	, (case " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sailyn " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then 'Y' " & vbcrlf
 		strSql = strSql + " 			else 'N' end) as sailyn " & vbcrlf
 		strSql = strSql + " 	, (case " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then FLOOR(100.0-100.0*i.sailprice/i.orgprice) " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then FLOOR(100.0-100.0*ss.saleprice/i.orgprice) " & vbcrlf
 		strSql = strSql + " 			else 0.0 end) as salePer " & vbcrlf
 		strSql = strSql + " 	, (case when DateDiff(d, i.regdate, getdate()) < 14 then 'Y' else 'N' end) as isNew, i.itemdiv " & vbcrlf
 		strSql = strSql + " 	, (case when cs.itemcoupontype = 1 then cs.itemcouponvalue else 0 end) as itemcoupon " & vbcrlf
 		strSql = strSql + " 	, (case " & vbcrlf
 		strSql = strSql + " 			when cs.itemcoupontype = 1 and ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then FLOOR((100-cs.itemcouponvalue)*(i.sellcash)/100) " & vbcrlf
 		strSql = strSql + " 			when cs.itemcoupontype = 1 and ss.saleprice is not NULL and i.orgprice > ss.saleprice then FLOOR((100-cs.itemcouponvalue)*(ss.saleprice)/100) " & vbcrlf
 		strSql = strSql + " 			when cs.itemcoupontype = 1 then FLOOR((100-cs.itemcouponvalue)*(i.orgprice)/100) " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sellcash " & vbcrlf
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then ss.saleprice " & vbcrlf
 		strSql = strSql + " 			else i.orgprice end) as itemcouponprice " & vbcrlf
 		strSql = strSql + " 	, i.optioncnt " & vbcrlf
 		strSql = strSql + " 	, ss.saleprice as startsaleprice " & vbcrlf
 		strSql = strSql + " 	, se.itemid as saleenditemid " & vbcrlf
 		strSql = strSql + " 	, cs.itemcoupontype, cs.itemcouponvalue " & vbcrlf
 		strSql = strSql + " 	, ce.itemid as couponenditemid, i.TENTENIMAGE600" & vbcrlf
 		strSql = strSql + " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash f " & vbcrlf
 		strSql = strSql + " left join [db_item].[dbo].tbl_item i on f.linkitemid=i.itemid " & vbcrlf
 		strSql = strSql + " left join ( " & vbcrlf
 		strSql = strSql + " 	select top 15 " & vbcrlf
 		strSql = strSql + " 		d.itemid, min(d.saleprice) as saleprice " & vbcrlf
 		strSql = strSql + " 	from db_event.dbo.tbl_sale as m with (noLock) " & vbcrlf
 		strSql = strSql + " 		join db_event.dbo.tbl_saleitem as d with (noLock) " & vbcrlf
 		strSql = strSql + " 			on m.sale_code=d.sale_code "
 		strSql = strSql + " 	where m.sale_status in (6,7) " & vbcrlf
 		strSql = strSql + " 		and d.saleItem_status in (6,7) " & vbcrlf
 		strSql = strSql + " 		and sale_using=1 " & vbcrlf
 		strSql = strSql + " 		and '" & Replace(FESDay, ".", "-") & "' between sale_startdate and sale_enddate " & vbcrlf
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") " & vbcrlf
 		strSql = strSql + " 	group by " & vbcrlf
 		strSql = strSql + " 		d.itemid " & vbcrlf
 		strSql = strSql + " ) ss on ss.itemid = i.itemid " & vbcrlf
 		strSql = strSql + " left join ( " & vbcrlf
 		strSql = strSql + " 	select top 15 " & vbcrlf
 		strSql = strSql + " 		d.itemid " & vbcrlf
 		strSql = strSql + " 	from db_event.dbo.tbl_sale as m with (noLock) " & vbcrlf
 		strSql = strSql + " 		join db_event.dbo.tbl_saleitem as d with (noLock) " & vbcrlf
 		strSql = strSql + " 			on m.sale_code=d.sale_code " & vbcrlf
 		strSql = strSql + " 	where m.sale_status in (6,9) " & vbcrlf
 		strSql = strSql + " 		and d.saleItem_status in (6,9) " & vbcrlf
 		strSql = strSql + " 		and sale_using=1 " & vbcrlf
 		strSql = strSql + " 		and datediff(day,sale_enddate,'" & Replace(FESDay, ".", "-") & "')=1 " & vbcrlf
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") " & vbcrlf
 		strSql = strSql + " 	group by " & vbcrlf
 		strSql = strSql + " 		d.itemid " & vbcrlf
 		strSql = strSql + " ) se on se.itemid = i.itemid " & vbcrlf
 		strSql = strSql + " left join ( " & vbcrlf
 		strSql = strSql + " 	select top 15 d.itemid, max(m.itemcoupontype) as itemcoupontype, min(m.itemcouponvalue) as itemcouponvalue " & vbcrlf
 		strSql = strSql + " 	from " & vbcrlf
 		strSql = strSql + " 		[db_item].[dbo].[tbl_item_coupon_master] m " & vbcrlf
 		strSql = strSql + " 		join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx " & vbcrlf
 		strSql = strSql + " 	where	m.openstate in (6,7) " & vbcrlf
 		strSql = strSql + " 		and '" & Replace(FESDay, ".", "-") & "' between m.itemcouponstartdate and itemcouponexpiredate " & vbcrlf
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") " & vbcrlf
 		strSql = strSql + " 		and m.itemcoupontype = 1 " & vbcrlf
		strSql = strSql + " 		and m.coupongubun = 'C' " & vbcrlf
 		strSql = strSql + " 	group by " & vbcrlf
 		strSql = strSql + " 		d.itemid " & vbcrlf
 		strSql = strSql + " ) cs on cs.itemid = i.itemid " & vbcrlf
 		strSql = strSql + " left join ( " & vbcrlf
 		strSql = strSql + " 	select top 15 d.itemid " & vbcrlf
 		strSql = strSql + " 	from " & vbcrlf
 		strSql = strSql + " 		[db_item].[dbo].[tbl_item_coupon_master] m " & vbcrlf
 		strSql = strSql + " 		join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx " & vbcrlf
 		strSql = strSql + " 	where	m.openstate in (6,9) " & vbcrlf
 		strSql = strSql + " 		and datediff(day,itemcouponexpiredate,'" & Replace(FESDay, ".", "-") & "')=1 " & vbcrlf
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") " & vbcrlf
		strSql = strSql + " 		and m.coupongubun = 'C' " & vbcrlf
 		strSql = strSql + " 	group by " & vbcrlf
 		strSql = strSql + " 		d.itemid " & vbcrlf
 		strSql = strSql + " ) ce on ce.itemid = i.itemid " & vbcrlf
 		strSql = strSql + " where 1=1 and f.isusing in ('Y','M') " & vbcrlf
 		strSql = strSql + " and f.linkitemid in (" & strItemidList & ") " & vbcrlf
 		strSql = strSql + " and f.startdate <= '" & Replace(FESDay, ".", "-") & "' " & vbcrlf
 		strSql = strSql + " and f.enddate >= '" & Replace(FESDay, ".", "-") & "' " & vbcrlf
 		strSql = strSql + " and f.linkitemid is not NULL " & vbcrlf
 		strSql = strSql + " order by f.disporder, f.startdate desc, f.idx desc "
		'response.write "<pre>" & strSql & "</pre>"
		'response.end
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzineMDPickData = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 저스트원데이 정보
	public Function fnGetMailzineJustOneDayData
		Dim strSql, i

		strSql = " select top 1 j.itemid, j.JustDate, j.orgPrice, j.justSalePrice, j.justDesc, i.icon1image, i.itemdiv, i.sailyn, FLOOR(100.0-100.0*i.sailprice/i.orgprice) as salePer, i.optioncnt, FLOOR(100.0-100.0*j.justSalePrice/j.orgPrice) as justSalePer "
		strSql = strSql + " from "
		strSql = strSql + " 	[db_sitemaster].[dbo].tbl_just1day j "
		strSql = strSql + " 	join [db_item].[dbo].tbl_item i "
		strSql = strSql + " 	on "
		strSql = strSql + " 		j.itemid = i.itemid "
		strSql = strSql + " where j.JustDate = '" & Replace(FESDay, ".", "-")& "' and j.itemid = " & FRectItemid
		''response.write strSql
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzineJustOneDayData = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 저스트원데이 정보
	public Function fnGetMailzineJustOneDayData2018
		Dim strSql, i

		strSql = " select top 3 "
		strSql = strSql + " 	d.itemid, convert(varchar(10), j.endDate, 121) as JustDate "
		strSql = strSql + " 	, i.orgPrice "
		strSql = strSql + " 	, IsNull(( "
		strSql = strSql + " 			select top 1 convert(varchar,convert(int,d.salePrice)) + '원' "
		strSql = strSql + " 			from "
		strSql = strSql + " 				db_event.dbo.tbl_sale as m "
		strSql = strSql + " 				join db_event.dbo.tbl_saleitem as d "
		strSql = strSql + " 				on "
		strSql = strSql + " 					m.sale_code=d.sale_code "
		strSql = strSql + " 				join [db_item].[dbo].tbl_item ii "
		strSql = strSql + " 				on "
		strSql = strSql + " 					d.itemid = ii.itemid "
		strSql = strSql + " 			where "
		strSql = strSql + " 				1 = 1 "
		strSql = strSql + " 				and d.itemid = i.itemid "
		strSql = strSql + " 				and i.itemdiv <> '21' "
		strSql = strSql + " 				and m.sale_status in (6,7) "
		strSql = strSql + " 				and d.saleItem_status in (6,7) "
		strSql = strSql + " 				and '" & Replace(FESDay, ".", "-" )& "' between sale_startdate and sale_enddate "
		strSql = strSql + " 			order by "
		strSql = strSql + " 				m.sale_code desc "
		strSql = strSql + " 		), d.price) as justSalePrice "
		strSql = strSql + " 	, d.title as justDesc, d.frontimage as icon1image, i.itemdiv, i.sailyn "
		strSql = strSql + " 	, i.optioncnt "
		strSql = strSql + " 	, IsNull(( "
		strSql = strSql + " 		select top 1 convert(varchar,convert(int,Round(100 - d.salePrice/ii.orgPrice*100,0))) + '%' "
		strSql = strSql + " 		 "
		strSql = strSql + " 		from "
		strSql = strSql + " 			db_event.dbo.tbl_sale as m "
		strSql = strSql + " 			join db_event.dbo.tbl_saleitem as d "
		strSql = strSql + " 			on "
		strSql = strSql + " 				m.sale_code=d.sale_code "
		strSql = strSql + " 			join [db_item].[dbo].tbl_item ii "
		strSql = strSql + " 			on "
		strSql = strSql + " 				d.itemid = ii.itemid "
		strSql = strSql + " 		where "
		strSql = strSql + " 			1 = 1 "
		strSql = strSql + " 			and d.itemid = i.itemid "
		strSql = strSql + " 			and i.itemdiv <> '21' "
		strSql = strSql + " 			and m.sale_status in (6,7) "
		strSql = strSql + " 			and d.saleItem_status in (6,7) "
		strSql = strSql + " 			and '" & Replace(FESDay, ".", "-" )& "' between sale_startdate and sale_enddate "
		strSql = strSql + " 		order by "
		strSql = strSql + " 			m.sale_code desc "
		strSql = strSql + " 	), d.salePer) as justSalePer "
		strSql = strSql + " 	, IsNull(( "
		strSql = strSql + " 		select top 1 (case when m.itemcoupontype = '1' then convert(varchar,convert(int,m.itemcouponvalue)) + '%' else '쿠폰' end) "
		strSql = strSql + " 		from "
		strSql = strSql + " 			[db_item].[dbo].[tbl_item_coupon_master] m "
		strSql = strSql + " 			join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx "
		strSql = strSql + " 		where "
		strSql = strSql + " 			1 = 1 "
		strSql = strSql + " 			and m.itemcoupontype in ('1', '3') "
		strSql = strSql + " 			and m.openstate in (6,7) "
		strSql = strSql + " 			and '" & Replace(FESDay, ".", "-" )& "' between m.itemcouponstartdate and itemcouponexpiredate "
		strSql = strSql + " 			and d.itemid = i.itemid "
		strSql = strSql + " 		order by m.itemcouponidx desc "
		strSql = strSql + " 	), '') as couponPer "
		strSql = strSql + " 	, i.icon1image "
		strSql = strSql + " from "
		strSql = strSql + " 	db_sitemaster.[dbo].[tbl_just1day2018_list] j "
		strSql = strSql + " 	join  "
		strSql = strSql + " 	db_sitemaster.dbo.tbl_just1day2018_item d "
		strSql = strSql + " 	on "
		strSql = strSql + " 		j.idx = d.listidx "
		strSql = strSql + " 	join [db_item].[dbo].tbl_item i "
		strSql = strSql + " 	on "
		strSql = strSql + " 		d.itemid = i.itemid "
		strSql = strSql + " where j.idx = " & FRectItemid
		strSql = strSql + " and d.isusing = 'Y' "
		strSql = strSql + " order by d.sortnum "
		''response.write strSql
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzineJustOneDayData2018 = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 텐바이텐 클래스 정보
	public Function fnGetMailzineTenTenClassData
		Dim strSql, i

		strSql = " select top 1 classDate, itemid1, salePer1, classDesc1, classSubDesc1, i1.itemdiv, i1.icon1image, itemid2, salePer2, classDesc2, classSubDesc2, i2.itemdiv, i2.icon1image, itemid3, salePer3, classDesc3, classSubDesc3, i3.itemdiv, i3.icon1image "
		strSql = strSql + " from "
		strSql = strSql + " 	[db_sitemaster].[dbo].[tbl_mailzine_class] c "
		strSql = strSql + " 	left join [db_item].[dbo].tbl_item i1 "
		strSql = strSql + " 	on "
		strSql = strSql + " 		c.itemid1 = i1.itemid "
		strSql = strSql + " 	left join [db_item].[dbo].tbl_item i2 "
		strSql = strSql + " 	on "
		strSql = strSql + " 		c.itemid2 = i2.itemid "
		strSql = strSql + " 	left join [db_item].[dbo].tbl_item i3 "
		strSql = strSql + " 	on "
		strSql = strSql + " 		c.itemid3 = i3.itemid "
		strSql = strSql + " where classDate = '" & Replace(FESDay, ".", "-")& "' "
		''response.write strSql
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzineTenTenClassData = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'// 메일진 표시용 신규 상품 추천 정보
	public Function fnGetMailzineMDPickNewBestData
		Dim strSql, i, strItemidList

		if FRectItemidArr="" or isnull(FRectItemidArr) then exit Function
		if FRectItemidArr<>"" then
			strItemidList = Replace(FRectItemidArr, vbCrLf, ",")
		end if

		strSql = " select top 15 "
		strSql = strSql + " 	f.linkitemid, 'http://imgstatic.10x10.co.kr/contents/maincontents/' + f.photoimg as photoimg, i.icon1image "
 		strSql = strSql + " 	, 'http://www.10x10.co.kr' + f.linkinfo as linkinfo, i.itemname as textinfo "
 		strSql = strSql + " 	, i.orgprice "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sellcash "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then ss.saleprice "
 		strSql = strSql + " 			else i.orgprice end) as sellcash "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sailyn "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then 'Y' "
 		strSql = strSql + " 			else 'N' end) as sailyn "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then FLOOR(100.0-100.0*i.sailprice/i.orgprice) "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then FLOOR(100.0-100.0*ss.saleprice/i.orgprice) "
 		strSql = strSql + " 			else 0.0 end) as salePer "
 		strSql = strSql + " 	, (case when DateDiff(d, i.regdate, getdate()) < 14 then 'Y' else 'N' end) as isNew, i.itemdiv "
 		strSql = strSql + " 	, (case when cs.itemcoupontype = 1 then cs.itemcouponvalue else 0 end) as itemcoupon "
 		strSql = strSql + " 	, (case "
 		strSql = strSql + " 			when cs.itemcoupontype = 1 and ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then FLOOR((100-cs.itemcouponvalue)*(i.sellcash)/100) "
 		strSql = strSql + " 			when cs.itemcoupontype = 1 and ss.saleprice is not NULL and i.orgprice > ss.saleprice then FLOOR((100-cs.itemcouponvalue)*(ss.saleprice)/100) "
 		strSql = strSql + " 			when cs.itemcoupontype = 1 then FLOOR((100-cs.itemcouponvalue)*(i.orgprice)/100) "
 		strSql = strSql + " 			when ss.saleprice is NULL and se.itemid is NULL and i.sailyn = 'Y' then i.sellcash "
 		strSql = strSql + " 			when ss.saleprice is not NULL and i.orgprice > ss.saleprice then ss.saleprice "
 		strSql = strSql + " 			else i.orgprice end) as itemcouponprice "
 		strSql = strSql + " 	, i.optioncnt "
 		strSql = strSql + " 	, ss.saleprice as startsaleprice "
 		strSql = strSql + " 	, se.itemid as saleenditemid "
 		strSql = strSql + " 	, cs.itemcoupontype, cs.itemcouponvalue "
 		strSql = strSql + " 	, ce.itemid as couponenditemid, i.TENTENIMAGE600"
 		strSql = strSql + " from [db_sitemaster].[dbo].tbl_main_mdchoice_Best_New f "
 		strSql = strSql + " left join [db_item].[dbo].tbl_item i on f.linkitemid=i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 15 "
 		strSql = strSql + " 		d.itemid, min(d.saleprice) as saleprice "
 		strSql = strSql + " 	from db_event.dbo.tbl_sale as m with (noLock) "
 		strSql = strSql + " 		join db_event.dbo.tbl_saleitem as d with (noLock) "
 		strSql = strSql + " 			on m.sale_code=d.sale_code "
 		strSql = strSql + " 	where m.sale_status in (6,7) "
 		strSql = strSql + " 		and d.saleItem_status in (6,7) "
 		strSql = strSql + " 		and sale_using=1 "
 		strSql = strSql + " 		and '" & Replace(FESDay, ".", "-") & "' between sale_startdate and sale_enddate "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) ss on ss.itemid = i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 15 "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " 	from db_event.dbo.tbl_sale as m with (noLock) "
 		strSql = strSql + " 		join db_event.dbo.tbl_saleitem as d with (noLock) "
 		strSql = strSql + " 			on m.sale_code=d.sale_code "
 		strSql = strSql + " 	where m.sale_status in (6,9) "
 		strSql = strSql + " 		and d.saleItem_status in (6,9) "
 		strSql = strSql + " 		and sale_using=1 "
 		strSql = strSql + " 		and datediff(day,sale_enddate,'" & Replace(FESDay, ".", "-") & "')=1 "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) se on se.itemid = i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 15 d.itemid, max(m.itemcoupontype) as itemcoupontype, min(m.itemcouponvalue) as itemcouponvalue "
 		strSql = strSql + " 	from "
 		strSql = strSql + " 		[db_item].[dbo].[tbl_item_coupon_master] m "
 		strSql = strSql + " 		join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx "
 		strSql = strSql + " 	where	m.openstate in (6,7) "
 		strSql = strSql + " 		and '" & Replace(FESDay, ".", "-") & "' between m.itemcouponstartdate and itemcouponexpiredate "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
 		strSql = strSql + " 		and m.itemcoupontype = 1 "
		strSql = strSql + " 		and m.coupongubun = 'C' "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) cs on cs.itemid = i.itemid "
 		strSql = strSql + " left join ( "
 		strSql = strSql + " 	select top 15 d.itemid "
 		strSql = strSql + " 	from "
 		strSql = strSql + " 		[db_item].[dbo].[tbl_item_coupon_master] m "
 		strSql = strSql + " 		join [db_item].[dbo].[tbl_item_coupon_detail] d on m.itemcouponidx = d.itemcouponidx "
 		strSql = strSql + " 	where	m.openstate in (6,9) "
 		strSql = strSql + " 		and datediff(day,itemcouponexpiredate,'" & Replace(FESDay, ".", "-") & "')=1 "
 		strSql = strSql + " 		and d.itemid in (" & strItemidList & ") "
		strSql = strSql + " 		and m.coupongubun = 'C' "
 		strSql = strSql + " 	group by "
 		strSql = strSql + " 		d.itemid "
 		strSql = strSql + " ) ce on ce.itemid = i.itemid "
 		strSql = strSql + " where 1=1 and f.isusing in ('Y','M') "
 		strSql = strSql + " and f.linkitemid in (" & strItemidList & ") "
 		strSql = strSql + " and f.startdate <= '" & Replace(FESDay, ".", "-") & "' "
 		strSql = strSql + " and f.enddate >= '" & Replace(FESDay, ".", "-") & "' "
 		strSql = strSql + " and f.linkitemid is not NULL "
		strSql = strSql + " and f.gubun='" & FEType & "'"
 		strSql = strSql + " order by f.disporder, f.startdate desc, f.idx desc "
		'response.write strSql
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetMailzineMDPickNewBestData = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'## fnGetEventMobileItemEvent : 모바일 상품 이벤트 내용 가저오기 ##
	public Function fnGetEventMobileItemEvent
		Dim strSql
		strSql = " SELECT evt_tagkind, evt_tagopt1, etc_opt1, etc_opt2 FROM db_event.dbo.tbl_event_mobile_addetc WHERE evt_code ="+ FECode +""
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetEventMobileItemEvent = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'## fnGetEventItem :이벤트상품 가져오기 ##
	public Function fnGetEventItem

		Dim strSql, strSqlCnt,iDelCnt
		Dim strSort,strGroup, striSort,addSql
		dim addSort
		addSort = ""
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

		'전시 카테고리 검색 필터 추가 정태훈(2021-01-12)
		if FRectDispCate<>"" then
			if LEN(FRectDispCate)>3 then
					addSql = addSql + " and B.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
			end if
			addSql = addSql + " and B.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

	IF FESGroup <> "" THEN
		IF FESGroup = 0 THEN
			strGroup = " AND (evtgroup_code  is null OR evtgroup_code =0 )"
		ELSE
		    if FEChannel ="P" THEN
			    strGroup = " AND evtgroup_code =  "&FESGroup
		    else  '모바일/App 은 합친 코드로 처리
		       strGroup = " AND evtgroup_code in (select evtgroup_code from db_event.dbo.tbl_Eventitem_Group where evtgroup_code_mo =  "&FESGroup&")"
		    end if
		END IF
	END IF

'신상품순 -1, 저가격순-2,지정번호순-3, 베스트셀러순-4,고가격순-5, 할인율순 - 6, 후기순후기순(별점3개이상) -7,위시순-8, (그룹순-9, 브랜드순-10)
	IF FESSort = "1" THEN
		strSort = "  , A.itemid DESC "
	ELSEIF FESSort = "2" THEN
		strSort = " , B.sellyn desc, lsold, B.sellcash ASC"
	ELSEIF FESSort = "3" THEN
	    if FEChannel ="P" THEN
		strSort = "  ,evtitem_imgsize desc, B.sellyn desc, lsold, evtitem_sort ,A.itemid desc"
        else
        strSort = "  , B.sellyn desc, lsold, evtitem_sort_mo ,A.itemid desc"
        end if
	ELSEIF FESSort = "4" THEN
		strSort = "  , B.sellyn desc, lsold, E.recentsellcount desc, E.sellcount desc, B.itemid desc"
	ELSEIF FESSort = "5" THEN
		strSort = " , B.sellyn desc, lsold, B.sellcash DESC"
	ELSEIF FESSort = "6" THEN
		strSort = " , sailpercent desc, B.sellyn desc, lsold, B.sellcash DESC"
	ELSEIF FESSort = "7" THEN
		strSort = " , B.sellyn desc, lsold, E.favcount desc"
	ELSEIF FESSort = "8" THEN
		strSort = "  , favcount desc "
	ELSEIF FESSort = "9" THEN
	      if FEChannel ="P" THEN
		strSort = "  , evtgroup_code desc,  evtitem_sort ,A.itemid desc "
        else
        strSort = "  , evtgroup_code_mo desc,  evtitem_sort ,A.itemid desc "
        end if

	ELSEIF FESSort = "10" THEN
		strSort = "  , makerid , evtitem_sort ,A.itemid desc  "
	ELSE
	     if FEChannel ="P" THEN
		strSort = "  , evtitem_sort ,A.itemid desc"
		else
        strSort = "  , evtitem_sort_mo ,A.itemid desc"
        end if
 	END IF

	strSqlCnt = " select isNull(sum(Totcnt),0), isNull(sum(isY),0), isNull(sum(isN),0), isNull(sum(isY_M),0), isNull(sum(isN_M),0) "&vbCrlf
	strSqlCnt = strSqlCnt & " from ( "&vbCrlf
	strSqlCnt = strSqlCnt &" SELECT COUNT(A.itemid) as Totcnt "&vbCrlf
	strSqlCnt = strSqlCnt &"   , case when a.evtitem_isDisp=1 then count(a.evtitem_isDisp) else 0 end as isY "&vbCrlf
	strSqlCnt = strSqlCnt &"   ,case when a.evtitem_isDisp=0 then count(a.evtitem_isDisp) else 0 end as isN "&vbCrlf
	strSqlCnt = strSqlCnt &"   , case when a.evtitem_isDisp_mo=1 then count(a.evtitem_isDisp_mo) else 0 end as isY_M "&vbCrlf
	strSqlCnt = strSqlCnt &"   ,case when a.evtitem_isDisp_mo=0 then count(a.evtitem_isDisp_mo) else 0 end as isN_M "&vbCrlf
	strSqlCnt = strSqlCnt &" FROM [db_event].[dbo].[tbl_eventitem] AS A "&vbCrlf
	strSqlCnt =	strSqlCnt &"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&vbCrlf
	strSqlCnt =	strSqlCnt &"	WHERE A.evt_code = "&FECode& strGroup&addSql & " and  A.evtitem_isUsing = 1 "&vbCrlf
	strSqlCnt =	strSqlCnt &" group by A.evtitem_isDisp, A.evtitem_isDisp_mo "
	strSqlCnt =	strSqlCnt &" ) as  T "

	rsget.Open strSqlCnt,dbget,1
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
		FDispYCnt =  rsget(1)
		FDispNCnt =  rsget(2)
		FDispYMCnt = rsget(3)
		FDispNMCnt = rsget(4)
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
        if FEChannel ="P" then
		strSql = " SELECT  TOP "&FPSize*FCPage&" A.itemid, A.evtgroup_code, A.evtitem_sort,  B.makerid, B.itemname, B.sellcash "&_
				"		,B.buycash,B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.mileage, B.smallimage, B.listimage,   B.sellyn, B.deliverytype "&_
				"	    ,  B.limityn, B.danjongyn, B.sailyn, B.isusing, B.limitno , B.limitsold, B.itemcouponyn, B.itemcoupontype, B.itemcouponvalue"&_
				"		 , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end as couponbuyprice "&_
				"		, B.mwdiv, A.evtitem_imgsize   	"&_
				"		, case sailyn when 'Y' then ((orgprice-sailprice)/ orgprice)*100 else 0 end  as sailpercent "&_
				"       , A.evtitem_isDisp, B.itemdiv"&_
				" ,case B.limityn when 'Y' then case when ((B.limitno-B.limitsold)<=0) then '2' else '1' end Else '1' end as lsold"&_
				"	FROM  [db_event].[dbo].[tbl_eventitem] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.evt_code = "&FECode & strGroup&addSql& " and  A.evtitem_isUsing = 1 "&_
				"   ORDER BY   evtitem_isDisp desc  "& strSort
		else
		strSql = " SELECT  TOP "&FPSize*FCPage&" A.itemid, A.evtgroup_code, A.evtitem_sort_mo,  B.makerid, B.itemname, B.sellcash "&_
				"		,B.buycash,B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.mileage, B.smallimage, B.listimage,   B.sellyn, B.deliverytype "&_
				"	    ,  B.limityn, B.danjongyn, B.sailyn, B.isusing, B.limitno , B.limitsold, B.itemcouponyn, B.itemcoupontype, B.itemcouponvalue"&_
				"		 , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end as couponbuyprice "&_
				"		, B.mwdiv, A.evtitem_imgsize   	"&_
				"		, case sailyn when 'Y' then ((orgprice-sailprice)/ orgprice)*100 else 0 end  as sailpercent "&_
				"       , A.evtitem_isDisp_mo, B.itemdiv"&_
				" ,case B.limityn when 'Y' then case when ((B.limitno-B.limitsold)<=0) then '2' else '1' end Else '1' end as lsold"&_
				"	FROM  [db_event].[dbo].[tbl_eventitem] AS A " &_
				"	 INNER JOIN [db_item].[dbo].tbl_item AS B ON A.itemid = B.itemid "&_
				"	LEFT OUTER JOIN [db_item].[dbo].[tbl_item_contents] AS E ON A.itemid = E.itemid "&_
				"	WHERE A.evt_code = "&FECode & strGroup&addSql& " and  A.evtitem_isUsing = 1 "&_
				"   ORDER BY    evtitem_isDisp_mo desc  "& strSort
	    end if
		'  response.write strSql&"<BR>"

		rsget.pagesize = FPSize
		rsget.Open strSql,dbget,1

        rsget.absolutepage = FCPage
		IF not rsget.EOF THEN
			fnGetEventItem = rsget.getRows()
		End IF
		rsget.Close

	END IF
	End Function

	public function fnGetEventItemCouponMax
		dim strSql
		'ID 검색
		strSql = "select max(B.itemcouponvalue) as itemcouponvalue"
		strSql = strSql + " from db_event.dbo.tbl_eventitem as A"
		strSql = strSql + " inner join db_item.dbo.tbl_item as B on A.itemid=B.itemid"
		strSql = strSql + " where A.evt_code=" & FECode
		strSql = strSql + " and A.evtitem_isusing=1"
		strSql = strSql + " and B.itemcoupontype=1"
		rsget.Open strSql,dbget,0
		IF not rsget.EOF THEN
			fnGetEventItemCouponMax = rsget("itemcouponvalue")
		End IF
		rsget.Close
	end function

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

	public function fnGetWorkerNameToID(WorkerName)
		dim strSql
		'ID 검색
		strSql = "select userid from db_partner.dbo.tbl_user_tenbyten where username='"&WorkerName&"' and statediv='Y' and isusing=1"
		rsget.Open strSql,dbget,0
		IF not rsget.EOF THEN
			fnGetWorkerNameToID = rsget("userid")
		End IF
		rsget.Close
	end function

	'## fnGetEventList : 이벤트목록  ##
	public Function fnGetEventList
	Dim strSql, strSqlCnt,iDelCnt, strSearch,strSort,strSubSort, DesignID

	if FSednm<>"" then
		DesignID = fnGetWorkerNameToID(FSednm)
	end if

	strSearch = ""

	'//정렬
	IF FSort = "SD" then
	    strSubSort = " A.evt_state Desc , A.evt_code desc "
	    strSort = " evt_state Desc , evt_code desc "
	ELSEIF FSort = "SA" then
	    strSubSort = " A.evt_state Asc ,  A.evt_code desc  "
	    strSort = " evt_state Asc , evt_code desc  "
	ELSEIF FSort = "DD" then
	    strSubSort = " A.evt_startdate Desc ,  A.evt_code desc  "
	     strSort = " evt_startdate Desc , evt_code desc  "
	ELSEIF FSort = "DA" then
	    strSubSort = "  A.evt_startdate Asc ,  A.evt_code desc  "
	    strSort = " evt_startdate Asc , evt_code desc  "
	ELSEIF FSort = "ID" then
	    strSubSort = " A.evt_imgregdate Desc ,  A.evt_code desc  "
	     strSort = " evt_imgregdate Desc , evt_code desc  "
	ELSEIF FSort = "IA" then
	    strSubSort = "  A.evt_imgregdate Asc ,  A.evt_code desc  "
	    strSort = " evt_imgregdate Asc , evt_code desc  "
	ELSEIF FSort = "CA" then
	    strSubSort = "  A.evt_code Asc "
	    strSort = " evt_code Asc "
	ELSE
	    strSubSort = " A.evt_code DESC "
	    strSort = " evt_code DESC "
    END IF

	'//검색조건
	If FSsDate <> ""  or FSeDate <> "" THEN
		if CStr(FSfDate) = "S" THEN
			strSearch  = strSearch & " and  datediff(day, '"&FSsDate&"', evt_startdate) >= 0 and  datediff(day,'"&FSeDate&"', evt_startdate) <=0  "
		elseif CStr(FSfDate) = "E" THEN
			strSearch  = strSearch & " and  datediff(day,'"&FSsDate&"',evt_enddate) >= 0 and  datediff(day,'"&FSeDate&"',evt_enddate) <=0  "
		elseif CStr(FSfDate) = "O" THEN
			strSearch  = strSearch & " and  datediff(day,'"&FSsDate&"',opendate) >= 0 and  datediff(day,'"&FSeDate&"',opendate) <=0  "
		end if
	END IF

	If FSeTxt <> "" THEN
		IF Cstr(FSfEvt) = "evt_code" THEN
			FSeTxt=getNumeric(FSeTxt)
			'이벤트 코드 검색
			'If chkWord(FSeTxt,"[^-0-9 ]") = "False" Then
			'	Alert_return("이벤트코드는 숫자만 입력하세요")
			'	response.end
			'End If
			if InStr(FSeTxt , ",,") > 0 then
				Alert_return("이벤트 코드를 확인해주세요.")
				response.end
			end if
			if right(FSeTxt,1) = "," then FSeTxt = left(FSeTxt,len(FSeTxt)-1)
			If FSeTxt <> "" THEN
				strSearch  = strSearch &  " and A.evt_code in (" & FSeTxt & ")"
			end if
		ElseIF Cstr(FSfEvt) = "evt_tag" THEN
			'이벤트 태그 검색
			strSearch  = strSearch &  " and B.evt_tag like '%"&FSeTxt&"%'"
		ElseIF Cstr(FSfEvt) = "evt_sub" THEN
			'이벤트 서브카피 검색
			strSearch  = strSearch &  " and  (A.evt_subcopyK like '%"&FSeTxt&"%' or A.evt_subname like '%"&FSeTxt&"%') "
		ELSE
			'이벤트명 + 작업태그 검색
			strSearch  = strSearch &  " and  (A.evt_name like '%"&FSeTxt&"%' or B.workTag like '%"&FSeTxt&"%') "
		END IF
	End If

	If FSstate <> "" THEN
		IF FSstate = "9" THEN	'종료
			strSearch  = strSearch & " and   (evt_state = 9 or  datediff(day,getdate(),evt_enddate)< 0 )"
		ELSEIF FSstate = "7" THEN	'오픈예정
		    strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)> 0 and  datediff(day,getdate(),evt_enddate)>=0 "
		ELSEIF FSstate = "6" THEN	'오픈진행중
			strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0 and datediff(day,getdate(),evt_enddate) >= 0  "
		ELSEIF FSstate = "1^3" THEN
		    strSearch  = strSearch & " and  ( evt_state = 1 or  evt_state = 3 ) and     datediff(day,getdate(),evt_enddate)>=0"
		ELSEIF FSstate = "6^9" THEN
		    strSearch  = strSearch & " and  (( evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0   ) or  evt_state = 9  or   datediff(day,getdate(),evt_enddate)< 0)        "
		ELSE
			strSearch  = strSearch & " and  evt_state = "&FSstate & " and  datediff(day,getdate(),evt_enddate)>=0"
		END IF
	End If

	If FRectEventType_PC <> "" THEN strSearch  = strSearch &  " and  eventtype_pc='" & FRectEventType_PC & "'"
	If FRectEventType_MO <> "" THEN strSearch  = strSearch &  " and  eventtype_mo='" & FRectEventType_MO & "'"

	If FScategory <> "" THEN strSearch  = strSearch &  " and  evt_category = "&FScategory
	If FScateMid <> "" THEN strSearch  = strSearch &  " and  evt_cateMid = "&FScateMid
	If FEDispCate<>"" then	strSearch  = strSearch &  " and  evt_dispcate like '"& FEDispCate & "%'"

	If FDispCateGroup<>"" Then strSearch  = strSearch &  " and  left(evt_dispcate,3) in ("& FDispCateGroup & ")"

	If FchComm <> "" THEN strSearch  = strSearch &  " and  iscomment=1"
	If FchBbs <> "" THEN strSearch  = strSearch &  " and  isbbs=1"
	If FchItemps <> "" THEN strSearch  = strSearch &  " and  isitemps=1"
	If Fisblogurl <> "" THEN strSearch  = strSearch &  " and  isGetBlogURL=1"

	IF FSkind <> "" THEN
		strSearch  = strSearch &  " and evt_kind in ("& FSkind & ") "
	END IF

	IF FSedid <> "" THEN
		strSearch  = strSearch &  " and (B.designerid = '"&FSedid&"' or B.designerid2 = '"&FSedid&"')"
	END IF
	IF FRectendlessView <> "" THEN
		strSearch  = strSearch &  " and B.endlessView = 'Y'"
	END IF

	IF FSemid <> "" THEN
		strSearch  = strSearch &  " and B.partMDid = '"&FSemid&"'"
	END IF

	IF FSepsid <> "" THEN
		strSearch  = strSearch &  " and B.publisherid = '"&FSepsid&"'"
	END IF

	IF FSedpid <> "" THEN
		strSearch  = strSearch &  " and B.developerid = '"&FSedpid&"'"
	END IF


	IF DesignID <> "" THEN
		''strSearch  = strSearch &  " and (C.username = '"&FSednm&"' or C2.username = '"&FSednm&"')"  ''느려서 방식 수정 2017/11/16 eastone
		strSearch  = strSearch &  " and ("
		strSearch  = strSearch &  " B.designerid='"&DesignID&"' or "
		strSearch  = strSearch &  " B.designerid2='"&DesignID&"')"
	END IF

	IF FSemnm <> "" THEN
		strSearch  = strSearch &  " and D.username = '"&FSemnm&"'"
	END IF

	IF FSepsnm <> "" THEN
		strSearch  = strSearch &  " and E.username = '"&FSepsnm&"'"
	END IF

	IF FSedpnm <> "" THEN
		strSearch  = strSearch &  " and F.username = '"&FSedpnm&"'"
	END IF

	IF FEBrand <> "" THEN
		strSearch  = strSearch & " and brand = '"&FEBrand&"'"
	END If

	IF FRectMDTheme <> "" THEN
		strSearch  = strSearch & " and (B.mdtheme='"&FRectMDTheme&"' or B.mdthememo='"&FRectMDTheme&"')"
	END IF

	'if FRectEvtType<>"" then strSearch  = strSearch & " and evt_type=" & FRectEvtType
	if FEDgStat1<>"" then strSearch  = strSearch & " and dsn_state1=" & FEDgStat1
	if FEDgStat2<>"" then strSearch  = strSearch & " and dsn_state2=" & FEDgStat2

	if FRectIsConfirm="1" then strSearch  = strSearch & " and evt_type=50 and isConfirm=1 "
	if FRectEvtManager <> "" then strSearch = strSearch & " and evt_manager ="&FRectEvtManager
	IF FSisSale = "1" THEN strSearch  = strSearch & " and issale = 1 "
	IF FSisGift = "1" THEN strSearch  = strSearch & " and isgift = 1 "
	IF FSisCoupon = "1" THEN strSearch  = strSearch & " and iscoupon = 1 "
	IF FSisOnlyTen = "1" THEN strSearch  = strSearch & " and isOnlyTen = 1 "
	IF FSisDiary = "1" THEN strSearch  = strSearch & " and isDiary = 1 "
	IF FSisoneplusone   = "1" THEN strSearch  = strSearch & " and isoneplusone = 1 "
	IF FSisfreedelivery = "1" THEN strSearch  = strSearch & " and isfreedelivery = 1 "
	IF FSisbookingsell = "1" THEN strSearch  = strSearch & " and isbookingsell = 1 "
	IF FSisNew = "1" THEN strSearch  = strSearch & " and isNew = 1 "

	if Not(FIsWeb="" and FIsMobile="" and FIsApp="") then
		IF FIsWeb = "1" then
			strSearch = strSearch & " and isWeb = 1 "
		else
			strSearch = strSearch & " and isWeb = 0 "
		end IF
		IF FIsMobile = "1" then
			strSearch = strSearch & " and isMobile=1 "
		else
			strSearch = strSearch & " and isMobile=0 "
		end if
		IF FIsApp = "1" then
			strSearch = strSearch & " and isApp=1 "
		else
			strSearch = strSearch & " and isApp=0 "
		end if
	end if

	if FRectStartESP <> "" And FRectEndESP <> "" then
		strSearch = strSearch & " and estimateSalePrice>="&FRectStartESP
		strSearch = strSearch & " and estimateSalePrice<="&FRectEndESP
	end if

	IF FisReqPublish ="1" then strSearch = strSearch & " and isReqPublish = 1 "

	if FETemp<>"" then
		strSearch = strSearch & " and evt_template="&FETemp
	end if
	
	if FETemp_mo<>"" then
		strSearch = strSearch & " and evt_template_mo="&FETemp_mo
	end if

	IF FchComm = "1" then
		strSearch = strSearch & " and iscomment = 1 "
	end IF
	IF FchBbs = "1" then
		strSearch = strSearch & " and isbbs=1 "
	end if
	IF FchItemps = "1" then
		strSearch = strSearch & " and isitemps=1 "
	end if
	IF Fisblogurl = "1" then
		strSearch = strSearch & " and isGetBlogURL=1 "
	end if

	IF FRectEvtLevel <> "" THEN
		strSearch  = strSearch & " and A.evt_level = '"&FRectEvtLevel&"'"
	END If

	strSqlCnt = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" &vbcrlf
	strSqlCnt =	strSqlCnt & "SELECT COUNT(A.evt_code) FROM [db_event].[dbo].[tbl_event] as A " &vbcrlf
	strSqlCnt =	strSqlCnt &	"   LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code " &vbcrlf
	strSqlCnt =	strSqlCnt &	"	LEFT OUTER JOIN [db_event].[dbo].[tbl_event_md_theme] as M ON A.evt_code = M.evt_code " &vbcrlf
	'    IF FSednm <> "" THEN
	'strSqlCnt =	strSqlCnt &	"	LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as C ON C.userid = B.designerid   and b.designerid is not null and b.designerid <> '' "
	'strSqlCnt =	strSqlCnt &	"	LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as C2 ON C2.userid = B.designerid   and b.designerid2 is not null and b.designerid2 <> '' "
    '    END IF
        IF FSemnm <> "" THEN
	strSqlCnt =	strSqlCnt &	"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as D ON D.userid = B.partMDid  and b.partMDid is not null and b.partMDid <> '' " &vbcrlf
	    END IF
	    IF FSepsnm <> "" THEN
	strSqlCnt =	strSqlCnt &	"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as E ON E.userid = B.publisherid and b.publisherid is not null and b.publisherid <> '' " &vbcrlf
	    END IF
	    IF FSedpnm <> "" THEN
	strSqlCnt =	strSqlCnt &	"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as F ON F.userid = B.developerid and b.developerid is not null and b.developerid <> '' " &vbcrlf
	    END IF
	strSqlCnt =	strSqlCnt &	" WHERE evt_using ='Y' "&strSearch

	'response.write strSqlCnt
	'response.end
	rsget.CursorLocation = adUseClient
	rsget.Open strSqlCnt, dbget, adOpenForwardOnly, adLockReadOnly
	IF not rsget.EOF THEN
		FTotCnt = rsget(0)
	End IF
	rsget.Close


	IF FTotCnt >0 THEN
		'이벤트 기간 종료시 상태 종료로 , 이벤트 오픈요청상태에서 기간이 진행중일때 상태 오픈으로 view 처리
		dim iSPageNo, iEPageNo
		iSPageNo = (FPSize*(FCPage-1)) + 1
		iEPageNo = FPSize*FCPage

		strSql = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" &vbcrlf
		strSql = strSql & "SELECT evt_code,evt_kind,evt_manager,evt_scope,evt_name,evt_startdate,evt_enddate,evt_level,evt_state,evt_regdate,evt_bannerimg " &vbcrlf
		strSql = strSql & "		,designername,categoryname,evt_prizedate,brand,issale,isgift,iscoupon,sale_count,gift_count,prizeyn" &vbcrlf
		strSql = strSql & "		,itemid,code_nm,mdname,evt_bannerimg2010,workTag,dispcate_nm,evt_itemsort,psname,dpname, isWeb, isMobile, isApp,isDiary " &vbcrlf
		strSql = strSql & "       ,etc_itemimg,evt_mo_listbanner, evt_imgregdate, evt_mo_listbannerTXT, ccname, evt_type  " &vbcrlf
		strSql = strSql & "       ,designerid2, designername2, dsn_state1, dsn_state2, eventtype_pc, eventtype_mo, iscomment, isbbs, isitemps, isGetBlogURL, evt_template, evt_template_mo, evt_copy_code, evt_mainimg, endlessView " &vbcrlf
		strSql = strSql & " FROM" &vbcrlf
		strSql = strSql & " ( " &vbcrlf
		strSql = strSql & "	SELECT ROW_NUMBER() OVER (ORDER BY  "&strSubSort&" ) as RowNum " &vbcrlf
		strSql = strSql & "		,A.evt_code, A.evt_kind, A.evt_manager, A.evt_scope, A.evt_name, A.evt_startdate, A.evt_enddate, A.evt_level  " &vbcrlf
		strSql = strSql & "		,evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9  " &vbcrlf
		strSql = strSql & "		When A.evt_state = 7 and DateDiff(day,getdate(),evt_startdate)  <= 0 Then 6 " &vbcrlf
		strSql = strSql & "		ELSE A.evt_state  " &vbcrlf
		strSql = strSql & "		end  " &vbcrlf
		strSql = strSql & "		,A.evt_regdate,B.evt_bannerimg, isNull(C.username,'') as designername " &vbcrlf
		strSql = strSql & "		,(SELECT code_nm from [db_item].[dbo].tbl_Cate_large WHERE code_large = B.evt_category) categoryname " &vbcrlf
		strSql = strSql & "		, A.evt_prizedate , B.brand, B.issale, B.isgift, B.iscoupon " &vbcrlf
		strSql = strSql & "		, (SELECT COUNT(sale_code) FROM [db_event].[dbo].[tbl_sale] WHERE evt_code = A.evt_code and sale_using =1) as sale_count  " &vbcrlf
		strSql = strSql & "		, (SELECT COUNT(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = A.evt_code and gift_using ='y') as gift_count " &vbcrlf
		strSql = strSql & "		, A.prizeyn  " &vbcrlf
		strSql = strSql & "		, (Case When A.evt_kind=13 Then (Select top 1 itemid from [db_event].[dbo].[tbl_eventitem] where evt_code=A.evt_code) else 0 end) as itemid  " &vbcrlf
		strSql = strSql & "		,(select top 1 code_nm from db_item.dbo.tbl_Cate_mid where code_large=b.evt_category and code_mid=b.evt_cateMid) as code_nm  " &vbcrlf
		strSql = strSql & "		, D.username as mdname " &vbcrlf
		strSql = strSql & "		, isNull(B.evt_bannerimg2010,'') AS evt_bannerimg2010, B.workTag  " &vbcrlf
		strSql = strSql & "		,(select top 1 catename from db_item.dbo.tbl_display_cate where catecode=left(b.evt_dispcate,3)) as dispcate_nm ,B.evt_itemsort " &vbcrlf
		strSql = strSql & "		, E.username as psname, F.username as dpname , A.isWeb, A.isMobile, A.isApp, B.isDiary ,etc_itemimg ,evt_mo_listbanner, evt_imgregdate, B.evt_mo_listbannerTXT " &vbcrlf
		strSql = strSql & "		, G.username as ccname, A.evt_type " &vbcrlf
		strSql = strSql & "		, isNull(B.designerid2,'') as designerid2, isNull(C2.username,'') as designername2, isNull(B.dsn_state1,'') as dsn_state1, isNull(B.dsn_state2,'') as dsn_state2, B.eventtype_pc, B.eventtype_mo " &vbcrlf
		strSql = strSql & "		, B.iscomment, B.isbbs, B.isitemps, B.isGetBlogURL, B.evt_template, B.evt_template_mo, M.evt_copy_code, B.evt_mainimg, B.endlessView " &vbcrlf
		strSql = strSql & "	FROM [db_event].[dbo].[tbl_event] as A  " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_md_theme] as M ON A.evt_code = M.evt_code " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as C ON C.userid = B.designerid   and b.designerid is not null and b.designerid <> '' " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as C2 ON C2.userid = B.designerid2   and b.designerid2 is not null and b.designerid2 <> '' " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as D ON D.userid = B.partMDid  and b.partMDid is not null and b.partMDid <> '' " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as E ON E.userid = B.publisherid and b.publisherid is not null and b.publisherid <> '' " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as F ON F.userid = B.developerid and b.developerid is not null and b.developerid <> '' " &vbcrlf
		strSql = strSql & "		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as G ON G.userid = B.codecheckerid and b.codecheckerid is not null and b.codecheckerid <> '' " &vbcrlf
		strSql = strSql & "	WHERE evt_using ='Y'  " &strSearch &vbcrlf
		strSql = strSql & ") AS TB " &vbcrlf
		strSql = strSql & " WHERE TB.RowNum Between "&iSPageNo&" AND "  &iEPageNo & " " &vbcrlf
		strSql = strSql & " order by "&strSort

''		response.Write strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
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
        strSearch =""
        IF FIsWeb = "1" then strSearch = strSearch & " and isWeb = 1 "
	    IF FIsMobile = "1" then strSearch = strSearch & " and isMobile = 1 "
	    IF FIsApp = "1" then strSearch = strSearch & " and isApp = 1 "

		IF FSkind <> "" THEN
		strSearch  = strSearch &  " and evt_kind = "& FSkind
		END IF

		If FSeTxt <> "" THEN
			IF Cstr(FSfEvt) = "evt_code" THEN
				strSearch  = strSearch &  " and  evt_code = "&FSeTxt
			ELSE
				strSearch  = strSearch &  " and  evt_name like '%"&FSeTxt&"%'"
			END IF
		End If

		strSqlCnt = " SELECT COUNT(evt_code) FROM [db_event].[dbo].[tbl_event] "&_
				" WHERE evt_using ='Y'  "&strSearch
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = "SELECT  TOP "&FPSize&" evt_code, evt_kind, evt_manager, evt_scope, evt_name, evt_startdate, evt_enddate, evt_level "&_
			 		" 		,evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9 "&_
		 		 	"					When  evt_state = 7 and DateDiff(day,getdate(),evt_startdate) <= 0 Then 6 "&_
		 		 	"					ELSE  evt_state end"&_
		 		 	"       , isWeb, isMobile, isApp "&_
					" FROM [db_event].[dbo].[tbl_event]  "&_
					" WHERE evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] " &_
					"			 WHERE evt_using ='Y' " &strSearch&" ORDER BY evt_code DESC ) as T ) "&_
					" and evt_using ='Y' "&strSearch&" ORDER BY evt_code DESC"
			rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetEventLastList = rsget.getRows()
			End IF
			rsget.Close
		End IF
	End Function

	public Function fnGetEventLastList2
		Dim strSearch, strSqlCnt, iDelCnt, strSql
        strSearch =""
        IF FIsWeb = "1" then strSearch = strSearch & " and e.isWeb = 1 "
	    IF FIsMobile = "1" then strSearch = strSearch & " and e.isMobile = 1 "
	    IF FIsApp = "1" then strSearch = strSearch & " and e.isApp = 1 "

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

		strSqlCnt = " SELECT COUNT(evt_code) FROM [db_event].[dbo].[tbl_event] e"&_
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
		 		 	"       , e.isWeb, e.isMobile, e.isApp, d.salePer, d.saleCPer, e.evt_subcopyK "&_
					" FROM [db_event].[dbo].[tbl_event] e"&_
					" left join [db_event].[dbo].[tbl_event_display] d on d.evt_code=e.evt_code"&_
					" WHERE e.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] " &_
					"			 WHERE evt_using ='Y' " &strSearch&" ORDER BY evt_code DESC ) as T ) "&_
					" and e.evt_using ='Y' "&strSearch&" ORDER BY e.evt_code DESC"

			rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetEventLastList2 = rsget.getRows()
			End IF
			rsget.Close
		End IF
	End Function

	'//아이템 복사 리스트(아이템이 포함된 이벤트 리스트)
	public Function fnGetEventLastItemList
		Dim strSearch, strSqlCnt, iDelCnt, strSql
        strSearch = ""
        IF FIsWeb = "1" then strSearch = strSearch & " and isWeb = 1 "
	    IF FIsMobile = "1" then strSearch = strSearch & " and isMobile = 1 "
	    IF FIsApp = "1" then strSearch = strSearch & " and isApp = 1 "

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
			strSql = strSql + "  , isWeb, isMobile, isApp "+ vbCrlf
			strSql = strSql + " FROM [db_event].[dbo].[tbl_event] as e  "+ vbCrlf
			strSql = strSql + " inner join db_event.dbo.tbl_eventitem as i  "+ vbCrlf
			strSql = strSql + " on e.evt_code = i.evt_code "+ vbCrlf
			strSql = strSql + " WHERE e.evt_code <=  ( SELECT MIN(evt_code) FROM  (SELECT Top "&iDelCnt&" evt_code FROM [db_event].[dbo].[tbl_event] " + vbCrlf
			strSql = strSql + "			 WHERE evt_using ='Y' " &strSearch&" ORDER BY evt_code DESC ) as T ) "+ vbCrlf
			strSql = strSql + " and evt_using ='Y' "&strSearch&"" + vbCrlf
			strSql = strSql + " group by e.evt_code, e.evt_kind, e.evt_manager, e.evt_scope, e.evt_name, e.evt_startdate , e.evt_enddate, e.evt_level ,evt_state , isWeb, isMobile, isApp " + vbCrlf
			strSql = strSql + " ORDER BY e.evt_code DESC"

			' Response.write strSql

			rsget.Open strSql,dbget,0
			IF not rsget.EOF THEN
				fnGetEventLastItemList = rsget.getRows()
			End IF
			rsget.Close
		End IF
	End Function

	'// 전시 카테고리 정보 접수 //
	public function getDispCategory(iid)
		dim SQL, i, strPrt

		SQL = "select top 1 isNull(db_item.dbo.getCateCodeFullDepthName(d.catecode),'') as catename" &_
			" from db_item.dbo.tbl_display_cate as d" &_
			" where d.catecode='" & iid & "'"
		rsget.Open SQL,dbget,1

		strPrt = ""
		if Not(rsget.EOf or rsget.BOf) then
			i = 0
			Do Until rsget.EOF
				strPrt = strPrt & Replace(rsget(0),"^^"," >> ")
				i = i + 1
			rsget.MoveNext
			Loop
		end if
		'결과값 반환
		getDispCategory = strPrt

		rsget.Close
	end Function

	'## fnGetRelationEvent : 관련 이벤트 리스트 가져오기 ##
	public Function fnGetRelationEvent
		Dim strSql
		strSql = " SELECT r.idx, r.ecode, r.viewidx, e.evt_kind, e.evt_name," + vbcrlf
		strSql = strSql + " evt_state = CASE WHEN DATEDIFF(day,getdate(),e.evt_enddate) < 0 THEN 9" + vbcrlf
		strSql = strSql + "				WHEN e.evt_state = 7 AND DATEDIFF(day,getdate(),e.evt_startdate) <= 0 THEN 6" + vbcrlf
		strSql = strSql + "				ELSE e.evt_state END" + vbcrlf
		strSql = strSql + " FROM db_event.dbo.tbl_relation_event r" + vbcrlf
		strSql = strSql + " LEFT JOIN db_event.dbo.tbl_event e on r.ecode=e.evt_code" + vbcrlf
		strSql = strSql + " WHERE r.mastercode=" & FECode & " ORDER BY viewidx ASC"
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetRelationEvent = rsget.getRows()
		End IF
		rsget.Close
	End Function

	'## fnGetLoginMileageEvent : 관련 이벤트 정보 가져오기 ##
	public Function fnGetLoginMileageEvent
		Dim strSql
		strSql = " SELECT top 1 mileage, jukyo" + vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_login_mileage]" + vbcrlf
		strSql = strSql + " WHERE evt_code=" & FECode
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetLoginMileageEvent = rsget.getRows()
		End IF
		rsget.Close
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
			strSql = " select  TOP "&FPSize&" A.evt_code, evt_kind,evt_manager,evt_scope, evt_name, evt_level, "&_
					" evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9 ELSE	A.evt_state end,"&_
					" evt_startdate, evt_enddate, (SELECT code_nm from  [db_item].[dbo].tbl_Cate_large WHERE code_large = B.evt_category) categoryname"&_
					" from  [db_event].[dbo].[tbl_event] as A "&_
					" 		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code " &_
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
    public FEChannel

	public FGImg_mo
	public FGlink_mo

	public FGDisp
	public FGBrand
	public FGLinkKind

	public Function fnGetRootGroup
		Dim strSql

		if FEChannel ="P" then
		 strSql = " SELECT evtgroup_code, evtgroup_desc FROM [db_event].[dbo].tbl_eventitem_group "&_
				" WHERE evt_code = "&FECode&" and evtgroup_pcode = 0 and evtgroup_using ='Y' and evtgroup_isDisp = '"&FGDisp&"'"
		else
		 strSql = " SELECT distinct evtgroup_code_mo, evtgroup_desc_mo FROM [db_event].[dbo].tbl_eventitem_group "&_
				" WHERE evt_code = "&FECode&" and evtgroup_pcode_mo = 0 and evtgroup_using ='Y' and evtgroup_isDisp_mo= '"&FGDisp&"' "
	    end if
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetRootGroup = rsget.getRows()
			End IF
			rsget.Close
	End Function

	'## fnGetEventItemGroup :이벤트화면설정 그룹내용가져오기 ##
	' event_modify , pop_eventitem_group
	public Function fnGetEventItemGroup
	IF FECode = "" THEN Exit Function
	Dim strSql
	if FEChannel ="P" then
	    strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
		    	"		(select evtgroup_desc from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode) "&_
		    	"           , evtgroup_isDisp, evtgroup_brand, evtgroup_linkkind "&_
		    	" FROM  [db_event].[dbo].[tbl_eventitem_group] as a " &_
			    "	WHERE evt_code = "&FECode&" and evtgroup_using ='Y'  ORDER BY evtgroup_depth, evtgroup_sort, evtgroup_code "
	else
	   strSql = " SELECT evtgroup_code, evtgroup_desc_mo, evtgroup_sort_mo, evtgroup_img_mo,evtgroup_link_mo,evtgroup_pcode_mo,evtgroup_depth_mo, "&_
		    	"		(select evtgroup_desc_mo from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode_mo) "&_
		    	"       , evtgroup_isDisp_mo,isNull(evtgroup_code_mo,0), evtgroup_brand_mo, evtgroup_linkkind_mo "&_
		    	" FROM  [db_event].[dbo].[tbl_eventitem_group] as a" &_
			    "	WHERE evt_code = "&FECode&" and evtgroup_using ='Y'  ORDER BY evtgroup_depth_mo, evtgroup_sort_mo, evtgroup_code_mo "
    end If

	rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetEventItemGroup = rsget.getRows()
		End IF
		rsget.Close

		rsget.Open "SELECT Year(evt_regdate) FROM [db_event].[dbo].[tbl_event] WHERE evt_code = '" & FECode & "'",dbget,1
		IF not rsget.EOF THEN
			FRegdate = rsget(0)
		End IF
		rsget.Close
	End Function



	public Function fnGetEventItemGroupCont
	Dim strSql
	IF FEGCode = "" THEN Exit Function
	IF FEChannel = "P" THEN
	strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
			"		isnull((select evtgroup_desc from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode),'최상위') as evtgroup_pdesc"&_
			"       , evtgroup_isDisp, evtgroup_brand, evtgroup_linkkind "&_
			"	FROM  [db_event].[dbo].[tbl_eventitem_group] as a " &_
			"	WHERE evt_code = "&FECode&" and evtgroup_code="&FEGCode&" and evtgroup_using ='Y' "
	ELSE
	 strSql = " SELECT evtgroup_code,evtgroup_desc_mo as evtgroup_desc, evtgroup_sort_mo as evtgroup_sort, evtgroup_img_mo as evtgroup_img,evtgroup_link_mo as evtgroup_link,evtgroup_pcode_mo as evtgroup_pcode ,evtgroup_depth_mo as evtgroup_depth, "&_
			"		isnull((select evtgroup_desc_mo from [db_event].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode_mo),'최상위') as evtgroup_pdesc"&_
			"       , evtgroup_isDisp_mo as evtgroup_isDisp, evtgroup_brand_mo as evtgroup_brand, evtgroup_linkkind_mo as evtgroup_linkkind "&_
			"	FROM  [db_event].[dbo].[tbl_eventitem_group] as a " &_
			"	WHERE evt_code = "&FECode&" and evtgroup_code="&FEGCode&" and evtgroup_using ='Y' "
    END If
	rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FGDesc = rsget("evtgroup_desc")
			FGSort = rsget("evtgroup_sort")
			FGImg  = rsget("evtgroup_img")
			FGPCode= rsget("evtgroup_pcode")
			FGDepth= rsget("evtgroup_depth")
			FGPDesc= rsget("evtgroup_pdesc")
			FGlink= rsget("evtgroup_link")
			FGDisp = rsget("evtgroup_isDisp")
			FGBrand = rsget("evtgroup_brand")
			FGLinkKind = rsget("evtgroup_linkkind")
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
	public FIsWeb
	public FIsMobile
	public FIsApp

	public Function fnGetEventConts
	 Dim strSql
	 strSql = " SELECT  evt_name, evt_startdate, evt_enddate, evt_state, brand, opendate, closedate, evt_scope, partner_id, isWeb, isMobile, isApp "&_
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
	 	 FIsWeb		= rsget("isWeb")
	 	 FIsMobile	= rsget("isMobile")
	 	 FIsApp		= rsget("isApp")
	 END IF
	 rsget.close
	End Function
End Class

Class CEventBannerItem
    public Fidx
	public Fitemid
    public Fitemname
    public Fviewidx
	public Fimgurl
	public Ftitle
	public Fgroupcode
	public Ficonnew
	public Ficonbest
	public Fbasicimage
	public FBrandID
    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CEventBanner
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectEvtCode
	public FRectSiteDiv
	public FRectGroupType
	public FRectIDX

	public function GetBannerItemList()
        dim sqlStr, addSql, i

        '// 본문 내용 접수
        sqlStr = "select top 5"
        sqlStr = sqlStr & " idx, evt_code, itemid, itemname, viewidx"
        sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_itembanner]"
        sqlStr = sqlStr & " where evt_code=" & CStr(FRectEvtCode)
		sqlStr = sqlStr & " and sdiv='" & CStr(FRectSiteDiv) & "'"
		sqlStr = sqlStr & " Order by idx asc"
        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)
        i=0
        if Not(rsget.EOF or rsget.BOF) then
            do until rsget.EOF
                set FItemList(i) = new CEventBannerItem
				FItemList(i).Fidx	= rsget("idx")
                FItemList(i).Fitemid	= rsget("itemid")
                FItemList(i).Fitemname	= rsget("itemname")
                FItemList(i).Fviewidx	= rsget("viewidx")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public function GetTrainThemeItemList()
        dim sqlStr, addSql, i

        '// 본문 내용 접수
        sqlStr = "select top 100"
        sqlStr = sqlStr & " g.idx, g.evt_code, g.title, g.itemid, g.itemname, g.imgurl, g.viewidx, g.groupcode, g.iconnew, g.iconbest, i.basicimage, g.makerid" & vbcrlf
        sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_manual_group] g" & vbcrlf
		sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] i on i.itemid=g.itemid" & vbcrlf
        sqlStr = sqlStr & " where g.evt_code='" & CStr(FRectEvtCode) & "'" & vbcrlf
		sqlStr = sqlStr & " and g.grouptype='" & CStr(FRectGroupType) & "'" & vbcrlf
		sqlStr = sqlStr & " Order by g.viewidx asc, g.idx asc"

        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)
        i=0
        if Not(rsget.EOF or rsget.BOF) then
            do until rsget.EOF
                set FItemList(i) = new CEventBannerItem
				FItemList(i).Fidx		= rsget("idx")
                FItemList(i).Ftitle		= rsget("title")
				FItemList(i).Fitemid	= rsget("itemid")
                FItemList(i).Fitemname	= rsget("itemname")
				FItemList(i).Fimgurl	= rsget("imgurl")
                FItemList(i).Fviewidx	= rsget("viewidx")
				FItemList(i).Fgroupcode	= rsget("groupcode")
				FItemList(i).Ficonnew	= rsget("iconnew")
				FItemList(i).Ficonbest	= rsget("iconbest")
				FItemList(i).Fbasicimage= webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/"  + rsget("basicimage")
				FItemList(i).FBrandID	= rsget("makerid")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public function GetMultiTrainThemeItemList()
        dim sqlStr, addSql, i

        '// 본문 내용 접수
        sqlStr = "select top 100"
        sqlStr = sqlStr & " c.idx, c.title, c.itemid, c.itemname, c.imgurl, c.viewidx"
		sqlStr = sqlStr & ", c.groupcode, c.iconnew, c.iconbest, i.basicimage, c.makerid" & vbcrlf
        sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_multi_contents] c" & vbcrlf
		sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] i on i.itemid=c.itemid" & vbcrlf
        sqlStr = sqlStr & " where c.menuidx='" & CStr(FRectIDX) & "'" & vbcrlf
		sqlStr = sqlStr & " and c.grouptype='" & CStr(FRectGroupType) & "'" & vbcrlf
		sqlStr = sqlStr & " Order by c.viewidx asc, c.idx asc"

        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)
        i=0
        if Not(rsget.EOF or rsget.BOF) then
            do until rsget.EOF
                set FItemList(i) = new CEventBannerItem
				FItemList(i).Fidx		= rsget("idx")
                FItemList(i).Ftitle		= rsget("title")
				FItemList(i).Fitemid	= rsget("itemid")
                FItemList(i).Fitemname	= rsget("itemname")
				FItemList(i).Fimgurl	= rsget("imgurl")
                FItemList(i).Fviewidx	= rsget("viewidx")
				FItemList(i).Fgroupcode	= rsget("groupcode")
				FItemList(i).Ficonnew	= rsget("iconnew")
				FItemList(i).Ficonbest	= rsget("iconbest")
				FItemList(i).Fbasicimage= webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/"  + rsget("basicimage")
				FItemList(i).FBrandID	= rsget("makerid")
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class

'------------------------------------------------------
'ClsMultiContentsMenu : 멀티 컨텐츠 메뉴
'------------------------------------------------------
Class  ClsMultiContentsMenu
	public Fmenudiv
	public Fviewsort
	public Fisusing
	public FGroupItemPriceView
	public FGroupItemCheck
	public FGroupItemType

	public FRectEvtCode
	public FRectIDX
	public Fidx
	public FBrandName
	public FBrandContents
	public FTitle
	public FCustomContents
	public FvideoFullLink
	public Fvideotype
	public FvideoLink
	public FRectDevice
	public FImgURL
	public FBGImage
	public FBGImagePC
	public FBGColorLeft
	public FBGColorRight
	public FcontentsAlign
	public FMargin '// 상단 여백(상품가격연동에서는 mobile만)
	public FtextColor

	public FGroupItemTitleName
	public FGroupItemViewType
	public FGroupItemBrandName
	public FsaleColor
	public FpriceColor
	public ForgpriceColor

	'/* 상품가격연동 */
	public FMarginBottom '// 하단 여백 - Mobile
	public FMarginColor '// 상단 여백 배경색 코드 - Mobile
	public FMarginBottomColor '// 하단 여백 배경색 코드 - Mobile
	public FMarginPC '// 상단 여백 - PC
	public FMarginBottomPC '// 하단 여백 - PC
	public FMarginColorPC '// 상단 여백 배경색 코드 - PC
	public FMarginBottomColorPC '// 하단 여백 배경색 코드 - PC

	public Function fnGetMultiContentsMenu
		Dim strSql
		strSql ="SELECT top 1 menudiv, viewsort, isusing, GroupItemPriceView, GroupItemCheck, GroupItemType"
		strSql = strSql + ", BGImage, BGImagePC, BGColorLeft, BGColorRight, contentsAlign, Margin, textColor" & vbcrlf
		strSql = strSql + ", GroupItemTitleName , GroupItemViewType , GroupItemBrandName , saleColor , priceColor , orgpriceColor" & vbcrlf
		strSql = strSql + ", MarginBottom , MarginColor , MarginBottomColor, MarginPC, MarginBottomPC , MarginColorPC , MarginBottomColorPC" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master]" & vbcrlf
		strSql = strSql + "	WHERE evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	and idx=" & FRectIDX & vbcrlf
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				Fmenudiv = rsget("menudiv")
				Fviewsort = rsget("viewsort")
				Fisusing = rsget("isusing")
				FGroupItemPriceView	= rsget("GroupItemPriceView")
			 	FGroupItemCheck	= rsget("GroupItemCheck")
			 	FGroupItemType = rsget("GroupItemType")
				FBGImage = rsget("BGImage")
				FBGImagePC = rsget("BGImagePC")
				FBGColorLeft = rsget("BGColorLeft")
				FBGColorRight = rsget("BGColorRight")
				FcontentsAlign = rsget("contentsAlign")
				FMargin = rsget("Margin")
				FtextColor = rsget("textColor")
				FGroupItemTitleName = rsget("GroupItemTitleName")
				FGroupItemViewType = rsget("GroupItemViewType")
				FGroupItemBrandName = rsget("GroupItemBrandName")
				FsaleColor = rsget("saleColor")
				FpriceColor = rsget("priceColor")
				ForgpriceColor = rsget("orgpriceColor")
				FMarginBottom = rsget("MarginBottom")
				FMarginColor = rsget("MarginColor")
				FMarginBottomColor = rsget("MarginBottomColor")
				FMarginPC = rsget("MarginPC")
				FMarginBottomPC = rsget("MarginBottomPC")
				FMarginColorPC = rsget("MarginColorPC")
				FMarginBottomColorPC = rsget("MarginBottomColorPC")
			END IF
		rsget.Close
	End Function

	public Function fnGetMultiContentsMenuList
		Dim strSql

		strSql ="SELECT M.idx, M.menudiv, M.viewsort, M.isusing" & vbcrlf
		strSql = strSql + ", (SELECT COUNT(idx) FROM [db_event].[dbo].[tbl_event_multi_contents] WHERE menuidx=M.idx AND device='M')" & vbcrlf
		strSql = strSql + ", (SELECT COUNT(idx) FROM [db_event].[dbo].[tbl_event_multi_contents] WHERE menuidx=M.idx AND device='W')" & vbcrlf
		strSql = strSql + ", CASE WHEN ISNULL(C.mobileimageurl,'')<>'' THEN C.mobileimageurl" & vbcrlf
		strSql = strSql + "   ELSE 'http://webimage.10x10.co.kr/image/basic/' + CONCAT(REPLICATE('0', 2 - LEN(C.itemid/10000)),(C.itemid/10000)) + '/' + I.basicimage END as mobile_image" & vbcrlf
		strSql = strSql + ", CASE WHEN ISNULL(C.pcimageurl,'')<>'' THEN C.pcimageurl" & vbcrlf
		strSql = strSql + "   ELSE 'http://webimage.10x10.co.kr/image/basic/' + CONCAT(REPLICATE('0', 2 - LEN(C.itemid/10000)),(C.itemid/10000)) + '/' + I.basicimage END as pc_image" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] AS M" & vbcrlf
		strSql = strSql + " LEFT JOIN (SELECT * FROM [db_event].[dbo].[tbl_event_multi_contents]) c ON m.idx=c.menuidx AND c.viewidx = 1" & vbcrlf
		strSql = strSql + " LEFT JOIN db_item.dbo.tbl_item I ON C.itemid = I.itemid" & vbcrlf
		strSql = strSql + "	WHERE M.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	order by M.viewsort asc, M.idx asc"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetMultiContentsMenuList = rsget.getRows()
			End IF
			rsget.Close
	End Function

	public Function fnGetMultiContentsSwifeList
		Dim strSql
		strSql ="SELECT c.idx, c.imgurl, c.viewidx, c.videoFullLink" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] m" & vbcrlf
		strSql = strSql + " LEFT JOIN [db_event].[dbo].[tbl_event_multi_contents] c ON m.idx=c.menuidx" & vbcrlf
		strSql = strSql + "	WHERE m.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	AND m.idx=" & FRectIDX & vbcrlf
		strSql = strSql + "	AND isnull(c.idx,'')<>''" & vbcrlf
		strSql = strSql + "	AND c.device='" + FRectDevice + "'" & vbcrlf
		strSql = strSql + "	order by c.viewidx asc, c.idx asc"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetMultiContentsSwifeList = rsget.getRows()
			End IF
		rsget.Close
	End Function

	public Function fnGetMultiContentsVideo
		Dim strSql
		strSql ="SELECT c.idx, c.videolink, c.videoFullLink, c.videotype, m.BGImage, m.BGColorLeft, m.BGColorRight, m.contentsAlign, m.Margin" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] m" & vbcrlf
		strSql = strSql + " LEFT JOIN [db_event].[dbo].[tbl_event_multi_contents] c ON m.idx=c.menuidx" & vbcrlf
		strSql = strSql + "	WHERE m.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	AND m.idx=" & FRectIDX & vbcrlf
		strSql = strSql + "	AND isnull(c.idx,'')<>''" & vbcrlf
		strSql = strSql + "	order by c.viewidx asc, c.idx asc"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				Fidx = rsget("idx")
				FvideoLink = rsget("videoLink")
				FvideoFullLink = rsget("videoFullLink")
				Fvideotype = rsget("videotype")
				FBGImage = rsget("BGImage")
				FBGColorLeft = rsget("BGColorLeft")
				FBGColorRight = rsget("BGColorRight")
				FcontentsAlign = rsget("contentsAlign")
				FMargin = rsget("Margin")
			End IF
		rsget.Close
	End Function

	public Function fnGetMultiContentsBrandStory
		Dim strSql
		strSql ="SELECT top 1 c.idx, c.BrandName, c.BrandContents, m.BGImage, m.BGColorLeft, m.BGColorRight, m.contentsAlign, m.Margin" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] m" & vbcrlf
		strSql = strSql + " LEFT JOIN [db_event].[dbo].[tbl_event_multi_contents] c ON m.idx=c.menuidx" & vbcrlf
		strSql = strSql + "	WHERE m.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	AND m.idx=" & FRectIDX & vbcrlf
		strSql = strSql + "	AND isnull(c.idx,'')<>''" & vbcrlf
		strSql = strSql + "	order by c.viewidx asc, c.idx asc"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				Fidx = rsget("idx")
				FBrandName = rsget("BrandName")
				FBrandContents = rsget("BrandContents")
				FBGImage = rsget("BGImage")
				FBGColorLeft = rsget("BGColorLeft")
				FBGColorRight = rsget("BGColorRight")
				FcontentsAlign = rsget("contentsAlign")
				FMargin = rsget("Margin")
			End IF
		rsget.Close
	End Function

	public Function fnGetMultiContentsGroupTemplateList
		Dim strSql
		strSql ="SELECT c.idx, c.title, c.itemid, c.itemname, c.imgurl, c.groupcode" & vbcrlf
		strSql = strSql + ", c.iconnew, c.iconbest, c.viewidx, c.grouptype, c.makerid, i.basicimage" & vbcrlf
		strSql = strSql + ", c.itemname2 , c.mobileimageurl , c.pcimageurl , c.xposition , c.yposition, i.brandname" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] m" & vbcrlf
		strSql = strSql + " LEFT JOIN [db_event].[dbo].[tbl_event_multi_contents] c ON m.idx=c.menuidx" & vbcrlf
		strSql = strSql & " LEFT JOIN [db_item].[dbo].[tbl_item] i on i.itemid=c.itemid" & vbcrlf
		strSql = strSql + "	WHERE m.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	AND m.idx=" & FRectIDX & vbcrlf
		strSql = strSql + "	AND isnull(c.idx,'')<>''" & vbcrlf
		strSql = strSql + "	order by c.viewidx asc, c.idx asc"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetMultiContentsGroupTemplateList = rsget.getRows()
			End IF
		rsget.Close
	End Function

	public Function fnGetMultiContentsCustomBox
		Dim strSql
		strSql ="SELECT top 1 c.idx, c.title, c.BrandContents as CustomContents, m.BGImage, m.BGColorLeft, m.BGColorRight, m.contentsAlign, m.Margin" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] m" & vbcrlf
		strSql = strSql + " LEFT JOIN [db_event].[dbo].[tbl_event_multi_contents] c ON m.idx=c.menuidx" & vbcrlf
		strSql = strSql + "	WHERE m.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	AND m.idx=" & FRectIDX & vbcrlf
		strSql = strSql + "	AND isnull(c.idx,'')<>''" & vbcrlf
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				Fidx = rsget("idx")
				FTitle = rsget("title")
				FCustomContents = rsget("CustomContents")
				FBGImage = rsget("BGImage")
				FBGColorLeft = rsget("BGColorLeft")
				FBGColorRight = rsget("BGColorRight")
				FcontentsAlign = rsget("contentsAlign")
				FMargin = rsget("Margin")
			End IF
		rsget.Close
	End Function

	public Function fnGetMultiContentsImgText
		Dim strSql
		strSql ="SELECT c.idx, c.imgurl, c.BrandContents, m.BGImage, m.BGColorLeft, m.BGColorRight, m.contentsAlign, m.Margin" & vbcrlf
		strSql = strSql + " FROM [db_event].[dbo].[tbl_event_multi_contents_master] m" & vbcrlf
		strSql = strSql + " LEFT JOIN [db_event].[dbo].[tbl_event_multi_contents] c ON m.idx=c.menuidx" & vbcrlf
		strSql = strSql + "	WHERE m.evt_code=" & FRectEvtCode & vbcrlf
		strSql = strSql + "	AND m.idx=" & FRectIDX & vbcrlf
		strSql = strSql + "	AND c.device='" + FRectDevice + "'" & vbcrlf
		strSql = strSql + "	AND isnull(c.idx,'')<>''" & vbcrlf
		strSql = strSql + "	order by c.viewidx asc, c.idx asc"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				Fidx = rsget("idx")
				FImgURL = rsget("imgurl")
				FBrandContents = rsget("BrandContents")
				FBGImage = rsget("BGImage")
				FBGColorLeft = rsget("BGColorLeft")
				FBGColorRight = rsget("BGColorRight")
				FcontentsAlign = rsget("contentsAlign")
				FMargin = rsget("Margin")
			End IF
		rsget.Close
	End Function

End Class

Class ClsEventbbs_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public frownum
	public Fsub_idx
	public Fuserid
	public FuserAge
	public Fuserlevel
	public Fregdate
	public Fjoindate
	public Fsub_opt1
	public Fsub_opt2
	public Fsub_opt3
	public Fsitegubun
	public fusercell
	public fdevice
	public fevtcom_idx
	public fevtcom_regdate
	public fevtcom_txt
	public fevtcom_point
	public fblogurl
	public fwincnt
	public fwindate
	public fusername
	public fusermail
	public fevtbbs_idx
	public fevtbbs_regdate
	public fevtbbs_subject
	public fevtbbs_content
	public fevtbbs_img1
	public fevtbbs_img2
	public fevtbbs_icon
end class

Class ClsEventbbs
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public frecteCode
	public frectSdate
	public frectEdate
	public frectlimitLevel

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	public sub fevent_comment_notpaging()
		dim sqlStr,i

		if frecteCode="" or isnull(frecteCode) then exit sub
		if frectSdate="" or isnull(frectSdate) then exit sub
		if frectEdate="" or isnull(frectEdate) then exit sub

		sqlStr = "select " &_
				"	t1.evtcom_idx, t1.userid " &_
				"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
				"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
				"		end as userAge " &_
				"	, t3.userlevel " &_
				"	, t1.evtcom_regdate, t2.regdate as joindate " &_
				"	, replace(replace(convert(varchar(max), t1.evtcom_txt), char(10), ''), char(13),'') as evtcom_txt" &_
				"	, t1.evtcom_point, t1.blogurl " &_
				"	,(select count(*) FROM [db_event].[dbo].[tbl_event_prize] as p WHERE p.evt_winner = t2.userid) as wincnt  " &_
				"	,(select top 1 evt_regdate FROM [db_event].[dbo].[tbl_event_prize] as p WHERE p.evt_winner = t2.userid order by evt_regdate desc) as windate " &_
				"	,t2.username, t2.usermail, t2.usercell " &_
				" from db_event.dbo.tbl_event_comment as t1 " &_
				"	Join db_user.[dbo].tbl_user_n as t2 " &_
				"		on t1.userid=t2.userid " &_
				"	Join db_user.[dbo].tbl_logindata as t3 " &_
				"		on t2.userid=t3.userid " &_
				" left join db_user.dbo.tbl_invalid_user iu" &_
				" 	on t1.userid=iu.invaliduserid" &_
				" 	and iu.isusing='Y'" &_
				" 	and iu.gubun='ONEVT'" &_
				" where iu.idx is null and t1.evt_code=" & frecteCode &_
				"	and t1.evtcom_using='Y' " &_
				"	and t1.evtcom_regdate between '" & frectSdate & "' and dateadd(d,1,'" & frectEdate & "') "

			Select Case frectlimitLevel
				Case "orange"
					sqlStr = sqlStr & "	and t3.userlevel not in ('5') "
				Case "yellow"
					sqlStr = sqlStr & "	and t3.userlevel not in ('0','5') "
				Case "white"
					sqlStr = sqlStr & "	and t3.userlevel not in ('0') "
			end Select

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ClsEventbbs_item

				FItemList(i).fevtcom_idx = rsget("evtcom_idx")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fuserAge = rsget("userAge")
				FItemList(i).fuserlevel = rsget("userlevel")
				FItemList(i).fevtcom_regdate = rsget("evtcom_regdate")
				FItemList(i).fjoindate = rsget("joindate")
				FItemList(i).fevtcom_txt = rsget("evtcom_txt")
				FItemList(i).fevtcom_point = rsget("evtcom_point")
				FItemList(i).fblogurl = rsget("blogurl")
				FItemList(i).fwincnt = rsget("wincnt")
				FItemList(i).fwindate = rsget("windate")
				FItemList(i).fusername = rsget("username")
				FItemList(i).fusermail = rsget("usermail")
				FItemList(i).fusercell = rsget("usercell")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub fevent_subscript_notpaging()
		dim sqlStr,i

		if frecteCode="" or isnull(frecteCode) then exit sub
		if frectSdate="" or isnull(frectSdate) then exit sub
		if frectEdate="" or isnull(frectEdate) then exit sub

		sqlStr = "select " &_
			"	t1.sub_idx, t1.userid " &_
			"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
			"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
			"		end as userAge " &_
			"	, t3.userlevel " &_
			"	, t1.regdate, t2.regdate as joindate " &_
			"	, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3  " &_
			"	, case t1.device  " &_
			"	When 'W' then 'pc웹'  " &_
			"	When 'M' then '모바일웹'  " &_
			"	When 'A' then '텐바이텐앱'  " &_
			"	 End as sitegubun " &_
			" from [db_event].[dbo].[tbl_event_subscript] as t1 " &_
			"	Join db_user.[dbo].tbl_user_n as t2 " &_
			"		on t1.userid=t2.userid " &_
			"	Join db_user.[dbo].tbl_logindata as t3 " &_
			"		on t2.userid=t3.userid " &_
			" left join db_user.dbo.tbl_invalid_user iu" &_
			" 	on t1.userid=iu.invaliduserid" &_
			" 	and iu.isusing='Y'" &_
			" 	and iu.gubun='ONEVT'" &_			
			" where iu.idx is null and t1.evt_code=" & frecteCode &_
			"	and t1.regdate between '" & frectSdate & "' and dateadd(d,1,'" & frectEdate & "') "

			Select Case frectlimitLevel
				Case "orange"
					sqlStr = sqlStr & "	and t3.userlevel not in ('5') "
				Case "yellow"
					sqlStr = sqlStr & "	and t3.userlevel not in ('0','5') "
				Case "white"
					sqlStr = sqlStr & "	and t3.userlevel not in ('0') "
			end Select

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ClsEventbbs_item

				FItemList(i).fsub_idx = rsget("sub_idx")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fuserAge = rsget("userAge")
				FItemList(i).fuserlevel = rsget("userlevel")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fjoindate = rsget("joindate")
				FItemList(i).fsub_opt1 = rsget("sub_opt1")
				FItemList(i).fsub_opt2 = rsget("sub_opt2")
				FItemList(i).fsub_opt3 = rsget("sub_opt3")
				FItemList(i).fsitegubun = rsget("sitegubun")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub fevent_subscriptlite_notpaging()
		dim sqlStr,i

		if frecteCode="" or isnull(frecteCode) then exit sub
		if frectSdate="" or isnull(frectSdate) then exit sub
		if frectEdate="" or isnull(frectEdate) then exit sub

		sqlStr = "select row_number() over(order by t1.userid asc) as rownum " &_
				"	, t1.userid " &_
				"	, t2.usercell " &_
				"	, t1.sub_idx " &_
				"	, t1.regdate " &_
				"	, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3 , t1.device " &_
				" from [db_event].[dbo].[tbl_event_subscript] as t1 " &_
				" left join db_user.dbo.tbl_invalid_user iu" &_
				" 	on t1.userid=iu.invaliduserid" &_
				" 	and iu.isusing='Y'" &_
				" 	and iu.gubun='ONEVT'" &_	
				" Join db_user.[dbo].tbl_user_n as t2 " &_
				"	on t1.userid=t2.userid " &_
				" where iu.idx is null and t1.evt_code=" & frecteCode &_
				"	and t1.regdate between '" & frectSdate & "' and dateadd(d,1,'" & frectEdate & "') "

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ClsEventbbs_item

				FItemList(i).frownum = rsget("rownum")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fusercell = rsget("usercell")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fsub_opt1 = rsget("sub_opt1")
				FItemList(i).fsub_opt2 = rsget("sub_opt2")
				FItemList(i).fsub_opt3 = rsget("sub_opt3")
				FItemList(i).fdevice = rsget("device")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub fevent_subscriptguest_notpaging()
		dim sqlStr,i

		if frecteCode="" or isnull(frecteCode) then exit sub
		if frectSdate="" or isnull(frectSdate) then exit sub
		if frectEdate="" or isnull(frectEdate) then exit sub

		sqlStr = "select " &_
				"	t1.sub_idx " &_
				"	, t1.regdate " &_
				"	, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3 " &_
				" from [db_event].[dbo].[tbl_event_subscript] as t1 " &_
				" left join db_user.dbo.tbl_invalid_user iu" &_
				" 	on t1.userid=iu.invaliduserid" &_
				" 	and iu.isusing='Y'" &_
				" 	and iu.gubun='ONEVT'" &_			
				" where iu.idx is null and t1.evt_code=" & frecteCode &_
				"	and t1.regdate between '" & frectSdate & "' and dateadd(d,1,'" & frectEdate & "') "

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ClsEventbbs_item

				FItemList(i).fsub_idx = rsget("sub_idx")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fsub_opt1 = rsget("sub_opt1")
				FItemList(i).fsub_opt2 = rsget("sub_opt2")
				FItemList(i).fsub_opt3 = rsget("sub_opt3")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub fevent_bbs_notpaging()
		dim sqlStr,i

		if frecteCode="" or isnull(frecteCode) then exit sub
		if frectSdate="" or isnull(frectSdate) then exit sub
		if frectEdate="" or isnull(frectEdate) then exit sub

		sqlStr = "select " &_
				"	t1.evtbbs_idx, t1.userid " &_
				"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
				"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
				"		end as userAge " &_
				"	, t3.userlevel " &_
				"	, t1.evtbbs_regdate, t2.regdate as joindate " &_
				"	, t1.evtbbs_subject, t1.evtbbs_content, t1.evtbbs_img1, t1.evtbbs_img2, t1.evtbbs_icon " &_
				" from db_event.dbo.tbl_event_bbs as t1 " &_
				"	Join db_user.[dbo].tbl_user_n as t2 " &_
				"		on t1.userid=t2.userid " &_
				"	Join db_user.[dbo].tbl_logindata as t3 " &_
				"		on t2.userid=t3.userid " &_
				" left join db_user.dbo.tbl_invalid_user iu" &_
				" 	on t1.userid=iu.invaliduserid" &_
				" 	and iu.isusing='Y'" &_
				" 	and iu.gubun='ONEVT'" &_			
				" where iu.idx is null and t1.evt_code=" & frecteCode &_
				"	and t1.evtbbs_using='Y' " &_
				"	and t1.evtbbs_regdate between '" & frectSdate & "' and dateadd(d,1,'" & frectEdate & "') "

			Select Case limitLevel
				Case "orange"
					sqlStr = sqlStr & "	and t3.userlevel not in ('5') "
				Case "yellow"
					sqlStr = sqlStr & "	and t3.userlevel not in ('0','5') "
				Case "white"
					sqlStr = sqlStr & "	and t3.userlevel not in ('0') "
			end Select

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ClsEventbbs_item

				FItemList(i).fevtbbs_idx = rsget("evtbbs_idx")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fuserAge = rsget("userAge")
				FItemList(i).fuserlevel = rsget("userlevel")
				FItemList(i).fevtbbs_regdate = rsget("evtbbs_regdate")
				FItemList(i).fjoindate = rsget("joindate")
				FItemList(i).fevtbbs_subject = rsget("evtbbs_subject")
				FItemList(i).fevtbbs_content = rsget("evtbbs_content")
				FItemList(i).fevtbbs_img1 = rsget("evtbbs_img1")
				FItemList(i).fevtbbs_img2 = rsget("evtbbs_img2")
				FItemList(i).fevtbbs_icon = rsget("evtbbs_icon")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end class
%>