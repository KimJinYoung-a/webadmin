<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event_modify.asp
' Description :  �̺�Ʈ ���� ���
' History : 2007.02.13 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/classes/event/eventColorCodeCls.asp"-->
<%
Dim eCode, sqlStr, tmp_cdl, tmp_cdm
Dim cEvtCont, sumi,i
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, echkdisp, eusing, etag, eonlyten, eisblogurl
Dim ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,elktype,elkurl,ebimg,etemp,emimg,ehtml,ehtml5, eisort,eiaddtype, edgid,edgid2,edgstat1,edgstat2, emdid ,efwd,ebrand,eicon,ebimg2010
Dim selPartner,dopendate,dclosedate, sWorkTag, ebimgMo, eDispCate, eDateView , ebimgToday , ebimgMo2014 , enamesub,dImgregdate, eCCId, eCCNm
Dim intI
Dim arrGift, intg,blngift
Dim eFolder, backUrl 
dim gimg : gimg = ""
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim ecommenttitle, elinkcode
Dim strparm , sCateMid
Dim cEGroup, arrGroup,arrGroup_mo, intgroup,strG, blngroup,vYear, blngroup_mo
Dim blnFull, blnIteminfo ,blnitemprice, evt_sortNo, blnWide
Dim enameEng , subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary, eNew
Dim eEtcitemid , eEtcitemimg, eItemListType
Dim eitemid, etype, isConfirm
Dim isWeb, isMobile, isApp, eDpid, ePsid, eDpnm, ePsnm, eDgnm, eDgnm2, eMdnm
dim tHtml5_mo, tHtml_mo, main_mo,emimg_mo,ehtml_mo,ehtml5_mo , efwd_mo
Dim maxDepth,dispCate
Dim eCmtCd,eIpsCd,eGfCd,eBSCd, rdCmt, eCmtMT, eCmtST, eIpsMT, eIpsST, eGfMT, eGfST, eBSMT, eBSST
dim arrText,intT
dim blnReqPublish,blnExec,eExecFile ,blnExec_mo ,eExecFile_mo  , etemp_mo
dim eSalePer
dim blnWeb,blnMobile,blnApp
dim rdIps,rdGf ,rdBS 
Dim sgroup_w , sgroup_m
Dim arrItemAdd , intA, videoLink, catesort
Dim tmpietag , tmpietagval , tmpmcopy , tmpscopy, eval_isusing, eval_text, eval_freebie_img, eval_start, eval_end
Dim slide_w_flag , slide_m_flag , evt_pc_addimg_cnt, evt_m_addimg_cnt, cbrand, etcitemtype, eventtype_pc, eventtype_mo, bannerTypeDiv, bannerCouponTxt, bannerGubun
Dim mdtheme, mdthememo, themecolor, themecolormo, textbgcolor, textbgcolormo, mdbntype, mdbntypemo, salePer, saleCPer, endlessView
Dim comm_isusing, comm_text, freebie_img, comm_start, comm_end, gift_isusing, gift_text1, gift_img1, nocate, title_pc, title_mo
Dim gift_text2, gift_img2, gift_text3, gift_img3, usinginfo, using_text1, using_contents1, using_text2, using_contents2, using_text3, using_contents3
Dim retUrl
retUrl = Request.ServerVariables("HTTP_REFERER")

If cbrand ="" Then cbrand="O"
If etcitemtype ="" Then etcitemtype="C"
eCode = requestCheckVar(Request("eC"),10)
ekind = requestCheckVar(Request("eK"),10)
 
maxDepth = 2 '����ī�װ� 2depth���� �����ش�
eItemListType = "1"
blnIteminfo = True
isConfirm = False
  
	'## �˻� #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'�Ⱓ
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'�̺�Ʈ �ڵ�/�� �˻�
	strTxt 		= requestCheckVar(Request("sEtxt"),120)

	sCategory	= requestCheckVar(Request("selC"),10) 		'ī�װ�
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'ī�װ�(�ߺз�)
	sState		= requestCheckVar(Request("eventstate"),4)	'�̺�Ʈ ����
	sKind 		= requestCheckVar(Request("eventkind"),4)	'�̺�Ʈ����
	etype		= requestCheckVar(Request("eventtype"),4)	'�̺�Ʈ����
	edgid  		= requestCheckVar(Request("sDgId"),32)		'��� �����̳�
	edgid2  	= requestCheckVar(Request("sDgId2"),32)		'���� �����̳�
	emdid  		= requestCheckVar(Request("sMdId"),32)		'��� MD
	epsid			= requestCheckVar(Request("sPsId"),32)		'��� �ۺ���
	edpid			= requestCheckVar(Request("sDpId"),32)		'��� ����
	eCCId			= requestCheckVar(Request("sCCId"),32)		'��� ���߰˼���
    
    edgnm  		= requestCheckVar(Request("sdgnm"),32)		'��� �����̳�
    edgnm2 		= requestCheckVar(Request("sdgnm2"),32)		'���� �����̳�
	emdnm  		= requestCheckVar(Request("smdnm"),32)		'��� MD
	epsnm  		= requestCheckVar(Request("spsnm"),32)		'��� �ۺ���
	edpnm  		= requestCheckVar(Request("sdpnm"),32)		'��� ������
    
	ebrand		= requestCheckVar(Request("ebrand"),32)		'�귣��
	esale		= requestCheckVar(Request("chSale"),2) 		'��������
	egift		= requestCheckVar(Request("chGift"),2)		'����ǰ����
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'��������
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen����

	eOneplusone	 	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery	= requestCheckVar(Request("chFreedelivery"),2)	'������
	eBookingsell	= requestCheckVar(Request("chBookingsell"),2)	'�����Ǹ�
	eDiary		= requestCheckVar(Request("chDiary"),2)	'���̾
	eNew		= requestCheckVar(Request("chNew"),2)
	dispCate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�

	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)
	eventtype_pc		= requestCheckVar(Request("eventtype_pc"),4)	'PC �̺�Ʈ����
	eventtype_mo		= requestCheckVar(Request("eventtype_mo"),4)	'MO �̺�Ʈ����
	catesort		= requestCheckVar(Request("catesort"),1)	'MD��Ʈ �˻�(1,2��Ʈ)
	
	strparm  = "isWeb="&blnWeb&"&isMobile="&blnMobile&"&isApp="&blnApp&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&sDgId="&edgid&"&sMdId="&emdid&"&sCCId="&eCCId&_
				"&sdgnm="&edgnm&"&smdnm="&emdnm&"&spsnm="&epsnm&"&sdpnm="&edpnm&"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&dispCate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary&"&eventtype="&etype&"&catesort="&catesort
				 
	'#######################################

	IF eCode = "" THEN	'�̺�Ʈ �ڵ尪�� ���� ��� back
		call sbAlertMsg("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�", "back","")
	END IF

	eFolder = eCode
	dim arreventmanager
	arreventmanager = fnSetCommonCodeArr("eventmanager",False)
'--------------------------------------------------------
' �̺�Ʈ ������ ��������
'--------------------------------------------------------
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ekind 		=	cEvtCont.FEKind
	eman 		=	cEvtCont.FEManager
	escope 		=	cEvtCont.FEScope
	selPartner	=	cEvtCont.FEPartnerID
	ename 		=	db2html(cEvtCont.FEName)
	enamesub	=	db2html(cEvtCont.FENamesub) '�̺�Ʈ Ÿ��Ʋ ����ī��
	enameEng =	db2html(cEvtCont.FENameEng) '�̺�Ʈ ���� �߰�
	subcopyK =	db2html(cEvtCont.FsubcopyK) '����ī�� �ѱ�
	subcopyE =	db2html(cEvtCont.FsubcopyE) '����ī�� ����
	esday 		=	cEvtCont.FESDay
	eeday 		=	cEvtCont.FEEDay
	epday 		=	cEvtCont.FEPDay
	elevel 		=	cEvtCont.FELevel
	estate 		=	cEvtCont.FEState
	IF datediff("d",now,eeday) <0 THEN estate = 9 '�Ⱓ �ʰ��� ����ǥ��
	eregdate	=	cEvtCont.FERegdate
	eusing		= 	cEvtCont.FEUsing
	evt_sortNo	= 	cEvtCont.FESortNo
	eitemid		=	cEvtCont.FEitemid
	isWeb		= cEvtCont.FIsWeb
	isMobile	= cEvtCont.FIsMobile
	isApp		= cEvtCont.FIsApp
	If etype = "" Then
		etype = cEvtCont.FEType
	End If
	isConfirm	= cEvtCont.FIsConfirm
	
 
	
	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay 
	tmp_cdl 		= cEvtCont.FECategory
	tmp_cdm			= cEvtCont.FECateMid
	DispCate		= cEvtCont.FEDispCate

	esale 			= cEvtCont.FESale
	egift 			= cEvtCont.FEGift
	ecoupon 		= cEvtCont.FECoupon
	ecomment 		= cEvtCont.FECommnet
	ebbs 			= cEvtCont.FEBbs
	eitemps 		= cEvtCont.FEItemps
	eapply 			= cEvtCont.FEApply
	elktype			= cEvtCont.FELinkType
	IF elktype="" Then elktype="E" '//��ũŸ�� �⺻�� ����
	elkurl			= cEvtCont.FELinkURL
	ebimg 			= cEvtCont.FEBImg
	ebimg2010		= cEvtCont.FEBImg2010
	ebimgMo			= cEvtCont.FEBImgMobile
	ebimgToday		= cEvtCont.FEBImgMoToday
	ebimgMo2014		= cEvtCont.FEBImgMoListBanner '//2014 ����� ����Ʈ ��� �߰�
	gimg			= cEvtCont.FEGImg
	etemp			= cEvtCont.FETemp
	etemp_mo        = cEvtCont.FETemp_mo
	if isNull(etemp) then etemp = 1
	if isNull(etemp_mo) then etemp_mo = 1

	If etemp="9" And etype<>"80" Then etemp = 3
	If etemp_mo="9" And etype<>"80" Then etemp_mo = 3
	if etemp = 5 or etemp = 6  THEN	'���۾� �̺�Ʈ �� ��� ó��
		ehtml5 		= db2html(cEvtCont.FEHtml) 
	else
		emimg 		= cEvtCont.FEMImg
		ehtml 		= db2html(cEvtCont.FEHtml) 
	end if
	
	if etemp_mo = 5 or etemp_mo = 6  THEN	'���۾� �̺�Ʈ �� ��� ó�� 
		ehtml5_mo 	= db2html(cEvtCont.FEHtml_mo)
	else 
		emimg_mo 	= cEvtCont.FEMImg_mo
		ehtml_mo 	= db2html(cEvtCont.FEHtml_mo)
	end if
	
	eisort 			= cEvtCont.FEISort
	edgid 			= cEvtCont.FEDgId
	emdid 			= cEvtCont.FEMdId
	efwd 			= db2html(cEvtCont.FEFwd)
	efwd_mo 		= db2html(cEvtCont.FEFwdMo)
	ebrand			= cEvtCont.FEBrand
	eicon   		= cEvtCont.FEIcon
	ecommenttitle   = db2html(cEvtCont.FECommentTitle)
	elinkcode       = cEvtCont.FELinkCode
	dopendate		= cEvtCont.FEOpenDate
	dclosedate		= cEvtCont.FECloseDate
	dImgregdate     = cEvtCont.FEImgRegdate
 	blnFull			= cEvtCont.FEFullYN
 	blnWide			= cEvtCont.FEWideYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN
 	etag			= db2html(cEvtCont.FETag)
 	eonlyten		= cEvtCont.FSisOnlyTen
 	eisblogurl		= cEvtCont.FSisGetBlogURL
 	sWorkTag		= cEvtCont.FWorkTag

	blnitemprice	= cEvtCont.FEItempriceYN

	eOneplusone	=	cEvtCont.FEOneplusOne
	eFreedelivery	=	cEvtCont.FEFreedelivery
	eBookingsell	=	cEvtCont.FEBookingsell
	eDiary 			= cEvtCont.FSisDiary
	eNew			= cEvtCont.FSisNew
	eEtcitemid		=	cEvtCont.FEtcitemid
	eEtcitemimg		=	cEvtCont.FEtcitemimg
	eDateView		= cEvtCont.FEdateview
	eItemListType = cEvtCont.FEListType

	edgid 			= cEvtCont.FEDgId
	edgid2 			= cEvtCont.FEDgId2
  	emdid 			= cEvtCont.FEMdId 
	epsid			= cEvtCont.FEPsId
	edpid			= cEvtCont.FEDpId
	eCCid			= cEvtCont.FECCId
	
	edgnm 			= cEvtCont.FEDgName
	edgnm2 			= cEvtCont.FEDgName2
  	emdnm 			= cEvtCont.FEMdName 
	epsnm			= cEvtCont.FEPsName
	edpnm			= cEvtCont.FEDpName
	eCCNm			= cEvtCont.FECCName

	edgstat1		= cEvtCont.FEDgStat1
	edgstat2		= cEvtCont.FEDgStat2

	blnReqPublish   = cEvtCont.FisReqPublish
	blnExec         = cEvtCont.FEisExec      
    eExecFile       = cEvtCont.FEexecFile    
    blnExec_mo      = cEvtCont.FEisExec_mo   
    eExecFile_mo    = cEvtCont.FEexecFile_mo 

	arrText			= cEvtCont.fnGetEventTextTitle

	arrItemAdd		= cEvtCont.fnGetEventMobileItemEvent '//������ �̺�Ʈ

	sgroup_w		= cEvtCont.FEsgroup_W '// �ֻ��� �������� ��
	sgroup_m		= cEvtCont.FEsgroup_M '// �ֻ��� �������� �����

	slide_w_flag	= cEvtCont.FESlide_W_Flag '// �����̵� ��
	slide_m_flag	= cEvtCont.FESlide_M_Flag '// �����̵� �����
	evt_pc_addimg_cnt	= cEvtCont.FEvt_pc_addimg_cnt '// PC �߰� �̹��� ī��Ʈ
	evt_m_addimg_cnt	= cEvtCont.FEvt_m_addimg_cnt '// ����� �߰� �̹��� ī��Ʈ
	mdtheme = cEvtCont.Fmdtheme
	mdthememo = cEvtCont.Fmdthememo
	themecolor = cEvtCont.Fthemecolor
	themecolormo = cEvtCont.Fthemecolormo
	textbgcolor = cEvtCont.Ftextbgcolor
	textbgcolormo = cEvtCont.Ftextbgcolormo
	mdbntype = cEvtCont.Fmdbntype
	mdbntypemo = cEvtCont.Fmdbntypemo
	salePer = cEvtCont.FsalePer
	saleCPer = cEvtCont.FsaleCPer
	endlessView = cEvtCont.FendlessView
	If eventtype_pc="" Then
	eventtype_pc = cEvtCont.Feventtype_pc
	End If
	If eventtype_mo="" Then
	eventtype_mo = cEvtCont.Feventtype_mo
	End If

	If themecolor = "" Then themecolor=11
	If themecolormo = "" Then themecolormo=11
	If textbgcolor = "" Then textbgcolor=1
	If textbgcolormo = "" Then textbgcolormo=1

	bannerTypeDiv = cEvtCont.FbannerTypeDiv
	bannerCouponTxt = cEvtCont.FbannerCouponTxt
	bannerGubun = cEvtCont.FbannerGubun
	videoLink=cEvtCont.FvideoLink

	'���� ��� �̺�Ʈ �׸� ����
	cEvtCont.fnGetEventMDThemeInfo
	comm_isusing = cEvtCont.Fcomm_isusing
	comm_text = cEvtCont.Fcomm_text
	freebie_img = cEvtCont.Ffreebie_img
	comm_start = cEvtCont.Fcomm_start
	comm_end = cEvtCont.Fcomm_end
	gift_isusing = cEvtCont.Fgift_isusing
	gift_text1 = cEvtCont.Fgift_text1
	gift_img1 = cEvtCont.Fgift_img1
	gift_text2 = cEvtCont.Fgift_text2
	gift_img2 = cEvtCont.Fgift_img2
	gift_text3 = cEvtCont.Fgift_text3
	gift_img3 = cEvtCont.Fgift_img3
	usinginfo = cEvtCont.Fusinginfo
	using_text1 = cEvtCont.Fusing_text1
	using_contents1 = cEvtCont.Fusing_contents1
	using_text2 = cEvtCont.Fusing_text2
	using_contents2 = cEvtCont.Fusing_contents2
	using_text3 = cEvtCont.Fusing_text3
	using_contents3 = cEvtCont.Fusing_contents3
	nocate = cEvtCont.Fnocate
	title_pc = cEvtCont.Ftitle_pc
	title_mo = cEvtCont.Ftitle_mo

	eval_isusing=cEvtCont.Feval_isusing
	eval_text=cEvtCont.Feval_text
	eval_freebie_img=cEvtCont.Feval_freebie_img
	eval_start=cEvtCont.Feval_start
	eval_end=cEvtCont.Feval_end

	set cEvtCont = Nothing
	IF elinkcode = 0 THEN elinkcode = ""

	 set cEGroup = new ClsEventGroup
	 	cEGroup.FECode = eCode   
	 	cEGroup.FEChannel = "P"    
	  	arrGroup    = cEGroup.fnGetEventItemGroup
	  	 
	    cEGroup.FEChannel = "M"        
	    arrGroup_mo    = cEGroup.fnGetEventItemGroup     
	    
	  	vYear = cEGroup.FRegdate
	 set cEGroup = nothing
 
	 blngroup = False
	 blngroup_mo = False
	 IF isArray(arrGroup) THEN blngroup = True
	 IF isArray(arrGroup_mo) THEN blngroup_mo = True

	 If eItemListType = "" OR isNull(eItemListType) Then eItemListType = "1" End If
	
		IF isArray(arrText) THEN
		For intT = 0 To UBound(arrText,2)
			IF arrText(1,intT) = 1 or arrText(1,intT) = 2 THEN
				eCmtCd = arrText(0,intT)
				rdCmt  = arrText(1,intT)		
				eCmtMT = arrText(2,intT)
				eCmtST = arrText(3,intT) 
			ELSEIF  arrText(1,intT) = 3 THEN
				eIpsCd = arrText(0,intT) 
				eIpsMT = arrText(2,intT)
				eIpsST = arrText(3,intT) 
			ELSEIF  arrText(1,intT) = 4 THEN
				eGfCd = arrText(0,intT) 
				eGfMT = arrText(2,intT)
				eGfST = arrText(3,intT) 
			ELSEIF  arrText(1,intT) = 5 THEN
				eBSCd = arrText(0,intT) 
				eBSMT = arrText(2,intT)
				eBSST = arrText(3,intT) 
			END IF	
		Next
	END If
	
	'//��ǰ�̺�Ʈ ����� & ��
	If ekind = "13" And (isMobile Or isApp) Then
		IF isArray(arrItemAdd) Then
			For intA = 0 To UBound(arrItemAdd,2)
				tmpietag	= arrItemAdd(0,intA)
				tmpietagval = arrItemAdd(1,intA)
				tmpmcopy	= arrItemAdd(2,intA)
				tmpscopy	= arrItemAdd(3,intA)
			Next 
		End If 
	End If 

	if eCmtST = "" then
	   eCmtST = "������ �ڸ�Ʈ�� �����ֽ�     ���� ��÷�Ͽ�           �� ������ �帳�ϴ�." 
    end if
	 
	if  eCmtMT ="" or eCmtST="" then
	    chkeCmt = 0
    end if
    if  eIpsMT ="" or eIpsST="" then
	    chkeIps = 0
    end if
    if  eGfMT =""   then
	    chkeGf = 0
    end if
    if eBSMT =""  then
	    chkeBS = 0
    end if
	if (ekind = 1 or ekind=23) and (eSale or ecoupon ) then
	    dim tmpename
	    tmpename = Split(ename,"|") 
	  			 
	  	if Ubound(tmpename)>0 then
		    ename = tmpename(0)
		    eSalePer = tmpename(1)
		 end if

    end if
	 
	if eisort = "" then eisort = 3 
   
  dim idepartmentid, sdepartmentname,clsMem
    '�μ��� ��������
set clsMem = new CTenByTenMember
	clsMem.Fuserid = emdid
	clsMem.fnGetDepartmentInfo
	idepartmentid		= clsMem.Fdepartment_id
 	sdepartmentname = clsMem.FdepartmentNameFull 
 set clsMem = Nothing

Dim oitembanner, moitembanner
set oitembanner = new CEventBanner
oitembanner.FRectEvtCode = eCode
oitembanner.FRectSiteDiv="w"
oitembanner.GetBannerItemList

set moitembanner = new CEventBanner
moitembanner.FRectEvtCode = eCode
moitembanner.FRectSiteDiv="m"
moitembanner.GetBannerItemList
%>
<style>
	select {font-size:12px; vertical-align:top;}
	input[type=button], input[type=text] {vertical-align:top;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" > 
//-- jsEvtSubmit : �̺�Ʈ ��� --//
	function jsEvtSubmit(frm){ 
		if(frm.eventkind.value == "29"){
			if(frm.sPsId.value == ""){
				alert("�ۺ������� �� ���Ǹ� �ؼ� ����ڸ� �������ּ���.!!");
				return false;
			}
			if(frm.sDpId.value == ""){
				alert("�ý��۰������� �� ���Ǹ� �ؼ� ����ڸ� �������ּ���.!!");
				return false;
			}
		}
		
		//ä�μ��� ���� Ȯ��
		if (!frm.blnWeb.checked&&!frm.blnMobile.checked&&!frm.blnApp.checked){
			alert("ä���� �������ּ���");
			frm.blnWeb.focus();
			return false;
		}

		//�������� ���� Ȯ��
		if(!frm.eventtype.value){
			alert("�̺�Ʈ ������ ������ �ּ���");
			frm.eventtype.focus();
			return false;
		}

	  //�귣�������̸� �̺�Ʈ�� ���ջ���
	  if(frm.eventkind.value=='16') {
		if(!frm.ebrand.value){
			alert("�귣�带 ������ �ּ���");
			frm.ebrand.focus();
			return false;
		}
		if(!frm.sEDN.value){
			alert("�̺�Ʈ���� �Է����ּ���");
			frm.sEDN.focus();
			return false;
		}
		if(frm.sMDc.value<=0){
			alert("�ִ� �������� �Է����ּ���");
			frm.sMDc.focus();
			return false;
		} else {
			frm.sEN.value = frm.sEDN.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value;
			frm.sENEng.value = frm.sEDNEng.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value; // �����̺�Ʈ
		}
	  }

	  //��ǰ�̺�Ʈ�ε� ��ǥ��ǰ �ڵ尡 0�̰ų� ������-2017-04-24 ���¿� �߰�
	  if(frm.eventkind.value=='13') {
		if(frm.etcitemid.value == 0 || frm.etcitemid.value == "" || isNaN(frm.etcitemid.value)){
			alert("��ǰ�̺�Ʈ�ϰ�� ��ǥ��ǰ�ڵ带 �����ž� �մϴ�.");
			frm.etcitemid.focus();
			return false;
		}
	  }
//	  if(frm.chSale.checked && frm.sSP.value==""){
//		  alert("��ǰ ���� ������ �Է����ּ���.");
//		  return false;
//	  }
//	  if(frm.chCoupon.checked && frm.sCSP.value==""){
//		  alert("���� ���� ������ �Է����ּ���.");
//		  return false;
//	  }
	  //alert(frm.disp2.value);
//	  if((frm.disp.value==0 || frm.disp2.value.length<6) && frm.nocate.checked==false){
//		  alert("ī�װ��� 2Depth���� �������ּ���.");
//		  return false;
//	  }
//	if(!frm.eventscope.value) {
//		alert("�̺�Ʈ ������ �������ּ���");
//		frm.chkEscope[0].focus();
//		return false;
//	}

<% If etype="80" Then %>
  if(frm.blnWeb.checked){
		if(!frm.title_pc.value){
			alert("PC �̺�Ʈ ���� ������ �̺�Ʈ���� �Է����ּ���");
			frm.title_pc.focus();
			return false;
		}
	}
<% end if%>

  if(frm.blnWeb.checked){
		if(!frm.subcopyK.value){
			alert("PC �̺�Ʈ ���� ������ ����ī�Ǹ� �Է����ּ���");
			frm.subcopyK.focus();
			return false;
		}
	}

<% If etype="80" Then %>
  if(frm.blnMobile.checked || frm.blnApp.checked){
		if(!frm.title_mo.value){
			alert("����� �̺�Ʈ ���� ������ �̺�Ʈ���� �Է����ּ���");
			frm.title_mo.focus();
			return false;
		}
	}
<% end if%>
  if(frm.blnMobile.checked || frm.blnApp.checked){
		if(!frm.subsEN.value){
			alert("Mobile/App �� ����ī�Ǹ� �Է����ּ���");
			frm.subsEN.focus();
			return false;
		}
	}

<% If etype="80" Then %>
  if(frm.blnWeb.checked){
		if(frm.mdtheme.value=="2"){
			if (frm.pcbannercheck.value=="N")
			{
				alert("PC �̺�Ʈ ���� ������ ��� �̹����� �������ּ���");
				return false;
			}
		}
	}
<% end if%>

<% If etype="80" Then %>
  if(frm.blnWeb.checked){
		if(frm.mdtheme.value=="3"){
			if (frm.pcitemcheck.value=="N")
			{
				alert("PC �̺�Ʈ ���� ������ ��ǰ�׸� ��ǰ�� �������ּ���");
				return false;
			}
		}
	}
<% end if%>

<% If etype="80" Then %>
  if(frm.blnMobile.checked || frm.blnApp.checked){
		if(frm.mdthememo.value=="2"){
			if (frm.mobannercheck.value=="N")
			{
				alert("����� �̺�Ʈ ���� ������ ��� �̹����� �������ּ���");
				return false;
			}
		}
	}
<% end if%>

<% If etype="80" Then %>
  if(frm.blnMobile.checked || frm.blnApp.checked){
		if(frm.mdthememo.value=="3"){
			if (frm.moitemcheck.value=="N")
			{
				alert("����� �̺�Ʈ ���� ������ ��ǰ�׸� ��ǰ�� �������ּ���");
				return false;
			}
		}
	}
<% end if%>

	  if(!frm.sEN.value){
		alert("�̺�Ʈ���� �Է����ּ���");
		if(frm.eventkind.options[frm.eventkind.selectedIndex].value == 4){
		 frm.selStatic.focus();
		}else{
		 frm.sEN.focus();
		}
		return false;
	  }


<% If etype<>"80" Then %>
	if (frm.selC.value == '110'){
		if (frm.selCM.value==''){
			alert('����ä���� ��ī�װ��� �����ؾ߸� �մϴ�');
			frm.selCM.focus();
			return false;
		}

	}
<% end if%>

		if (frm.comm_isusing.value=="Y" &&  frm.comm_text.value ==""){
			alert("�ڸ�Ʈ ������ �Է��� �ֽðų� �������� üũ���ּ���");
			frm.comm_text.focus();
			return false;
		}

	  if(!frm.sSD.value || !frm.sED.value ){
		alert("�̺�Ʈ �Ⱓ�� �Է����ּ���");
		frm.sSD.focus();
		return false;
	  }

	  if(frm.sSD.value > frm.sED.value){
		alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
		frm.sED.focus();
		return false;
	  }

	var nowDate = jsNowDate();

	 if(frm.eventstate.value==7){
		if(frm.eOD.value !=""){
			nowDate = '<%IF dopendate <> ""THEN%><%=FormatDate(dopendate,"0000-00-00")%><%END IF%>';
		}
	 }


	 if(frm.eventstate.value < 7){
		if(frm.sSD.value < nowDate){
			alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
			frm.sSD.focus();
			return false;
		 }

		if(frm.sED.value < jsNowDate()){
			alert("�������� ���糯¥���� ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ��� ");
			frm.sED.focus();
			return false;
		}
	}

	if((frm.chComm.checked||frm.chBbs.checked||frm.chItemps.checked||frm.isblogurl.checked)&&frm.sPD.value=="") {
		alert("��÷�� ��ǥ���� �������ּ��� ");
		frm.sPD.focus();
		return false;
	}

	if(frm.sDgId.value!="" && frm.designerstatus[0].value==""){
		alert("�����̳�(PC) ���¸� �������ּ���.");
		frm.designerstatus[0].focus();
		return false;
	}
	if(frm.sDgId2.value!="" && frm.designerstatus[1].value==""){
		alert("�����̳�(Mobile) ���¸� �������ּ���.");
		frm.designerstatus[1].focus();
		return false;
	}

//	    if(!frm.eCT.value){
//	  		if(GetByteLength(frm.eCT.value) > 200){
//	  			alert("comment title�� 200�� �̳��� �ۼ����ּ���");
//	  			frm.eCT.focus();
//	  			return false;
//	  		}
//	  	}


		if(GetByteLength(frm.eTag.value) > 250){
			alert("Tag�� 250�� �̳��� �ۼ����ּ���");
			frm.eTag.focus();
			return false;
		}

		if(document.all.dvCmt.style.display ==""){
			if (!frm.chkeCmt.checked &&  (!frm.eCmtMT.value ||  !frm.eCmtST.value)){
				alert("�ڸ�Ʈ ������ �Է��� �ֽðų� �������� üũ���ּ���");
				frm.eCmtMT.focus();
				return false;
			}
		}



		  if(document.all.dvIps.style.display ==""){ 
		   if (!frm.chkeIps.checked &&  (!frm.eIpsMT.value ||  !frm.eIpsST.value)){
				alert("��ǰ�ı� ������ �Է��� �ֽðų� �������� üũ���ּ���");
				frm.eIpsMT.focus();
				return false;
			}
		}
		
		
		  if(document.all.dvGf.style.display ==""){
			if (!frm.chkeGf.checked && !frm.eGfMT.value ){
				alert("����ǰ ������ �Է��� �ֽðų� �������� üũ���ּ���");
				frm.eGfMT.focus();
				return false;
			}
		}
		
		  if(document.all.dvBS.style.display ==""){
			if (!frm.chkeBS.checked && !frm.eBSMT.value ){
				alert("�����Ǹ� ������ �Է��� �ֽðų� �������� üũ���ּ���");
				frm.eBSMT.focus();
				return false;
			}
		}
	}

	function jsNowDate(){
	var mydate=new Date()
		var year=mydate.getYear()
		    if (year < 1000)
		        year+=1900

		var day=mydate.getDay()
		var month=mydate.getMonth()+1
		    if (month<10)
		        month="0"+month

		var daym=mydate.getDate()
		    if (daym<10)
		        daym="0"+daym

		return year+"-"+month+"-"+ daym
	}

//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
 


	function jsChkSubj(chk){
		if(chk=='16') {
			//�귣�������ϰ�쿡�� ���� ��� ������ ������ ǥ��
			eNameTr_A.style.display = "none";
			eNameTr_C.style.display = "none";
			eNameTr_B.style.display = "";
			eNameTr_BL.style.display= "";
		}else if(chk=='13') {
			//��ǰ�̺�Ʈ
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
			eNameTr_BL.style.display= "none";
			itemevt.style.display = ""; // ��ǰ�̺�Ʈ
		} else {
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
			eNameTr_BL.style.display= "none";
		}
		
		if(chk=='22'){
			document.all.divDE.style.display = "";
		}else{
			document.all.divDE.style.display = "none";
		}
		
		if((chk=='1'|| chk=='23')  && (document.frmEvt.chSale.checked || document.frmEvt.chCoupon.checked)){ //�������� �϶� ������ ǥ��
		     document.all.spSale.style.display = "";
		     if (document.frmEvt.chSale.checked) {
		        document.all.spSale.style.color = "red";
		      }else{
		        document.all.spSale.style.color = "green";
		      }
		}else{
		     document.frmEvt.sSP.value ="";
		     document.all.spSale.style.display = "none"; 
		}
	}
	 

//-- jsLastEvent : ���� �̺�Ʈ �ҷ����� --//
	function jsLastEvent(){
	  var winLast,eKind;
	  eKind = document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value;
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}


	//���޸� ǥ��
	function jsSetPartner(){
		if(document.frmEvt.chkEscope[0].checked&&document.frmEvt.chkEscope[1].checked) {
			document.all.spanP.style.display ="";
			document.frmEvt.eventscope.value="1";
		} else if(document.frmEvt.chkEscope[0].checked) {
			document.all.spanP.style.display ="none";
			document.frmEvt.eventscope.value="2";
		} else if(document.frmEvt.chkEscope[1].checked) {
			document.all.spanP.style.display ="";
			document.frmEvt.eventscope.value="3";
		} else {
			document.all.spanP.style.display ="none";
			document.frmEvt.eventscope.value="";
		}
	}

	 
	
	function jsGetID(sType, iCid, sUserID){
		var openWorker = window.open('PopWorkerList.asp?sType='+sType+'&department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function jsDelID(sType){ 
		eval("document.frmEvt.s"+sType+"Id").value = "";
		eval("document.frmEvt.s"+sType+"Nm").value = ""; 
	}
	
	//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}


	function jsSetImg(sFolder, sImg, sName, sSpan){ 
		var winImg;
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsManageEventImage(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/V2/eventManageDir.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	function jsManageEventImageNew(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/V2/eventManageDir_new.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}
 
 	function jsAddGroup(eCode,gCode, smode, eChannel){ 
		var winG 
		var vYear = "<%=vYear%>";  
		winG = window.open('/admin/eventmanage/event/v3/pop_eventitem_group.asp?yr='+vYear+'&eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel,'popG','width=800, height=800,scrollbars=yes,resizable=yes');
		winG.focus();
	}

	function jsAddGroup2(eCode,gCode, smode, eChannel,EType){ 
		var winG 
		var vYear = "<%=vYear%>";  
		winG = window.open('/admin/eventmanage/event/v3/pop_eventitem_group.asp?yr='+vYear+'&eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel+'&etype='+EType,'popG','width=800, height=800,scrollbars=yes,resizable=yes');
		winG.focus();
	}
	
	function jsAddProcGroup(eCode, smode, sModeType,eChannel)
	{ 
	    document.frmG.target="ifrmProc";
	    document.frmG.mode.value = smode;
	    document.frmG.eCh.value = eChannel;
	    document.frmG.eMT.value = sModeType
	    document.frmG.submit();
	}

	function jsGroupImg(eCode,gCode,eChannel){
		var vYear = "<%=vYear%>";
		var winG = window.open('pop_eventitem_groupImage.asp?yr='+vYear+'&eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel,'popG','width=700, height=600,scrollbars=yes,resizable=yes');
		winG.focus();
	}

	function jsChangeFrm(iVal,sType){  
		if (sType =="P"){
		    $("div[id^='divFrm']").hide();  
    		$("#divGM").hide(); 
			//$("#w_slide").hide();
			$("#w_slide1").prop("disabled", true);
			$("#w_slide2").prop("disabled", true);
			$("#w_slide3").prop("disabled", true);
    		
    		if(iVal == 3 || iVal == 7){  
    			$("#divGM").show(); 
    			$("#divFrm3").show();
				//$("#w_slide").show();
				$("#w_slide1").prop("disabled", false);
				$("#w_slide2").prop("disabled", false);
				$("#w_slide3").prop("disabled", false);
    		}else if(iVal == 5 || iVal == 6 ){
    			//iframG.location.href = "about;blank"; 
    			$("#divFrm5").show(); 
    		}else{
    			//iframG.location.href = "about;blank"; 
    			$("#divFrm1").show();
				//$("#w_slide").show();
				$("#w_slide1").prop("disabled", false);
				$("#w_slide2").prop("disabled", false);
				$("#w_slide3").prop("disabled", false);
    		} 
    	}else if (sType=="M") {
    	    $("div[id^='divMFrm']").hide();  
    	    $("#divGM_mo").hide();
			//$("#m_slide").hide();
			$("#m_slide1").prop("disabled", true);
			$("#m_slide2").prop("disabled", true);
			$("#m_slide3").prop("disabled", true);
    		if(iVal == 3 || iVal == 7){  
    			$("#divGM_mo").show();  
    			$("#divMFrm3").show();
				//$("#m_slide").show();
				$("#m_slide1").prop("disabled", false);
				$("#m_slide2").prop("disabled", false);
				$("#m_slide3").prop("disabled", false);
    		}else if(iVal == 5 || iVal == 6 ){ 
    			$("#divMFrm5").show();
    		}else{  
    			$("#divMFrm1").show();
				//$("#m_slide").show();
				$("#m_slide1").prop("disabled", false);
				$("#m_slide2").prop("disabled", false);
				$("#m_slide3").prop("disabled", false);
    		} 
    	}else if (sType=="DG1") {
			if(iVal==""){
				document.frmEvt.designerstatus[0].value = "";
			} else {
				document.frmEvt.designerstatus[0].value = "20";
			}
    	}else if (sType=="DG2") {
			if(iVal==""){
				document.frmEvt.designerstatus[1].value = "";
			} else {
				document.frmEvt.designerstatus[1].value = "20";
			}
    	}
	}
	
	//����� �ؽ�ƮŸ��
	function jsChkTitle(sType){  
		if(sType=="g") { 
		 	if (document.frmEvt.chGift.checked){
				document.all.dvGf.style.display ="";
		 	}else{
		 		document.all.dvGf.style.display ="none";
			}	 
		}else if (sType=="i"){
			if (document.frmEvt.chItemps.checked){
				<% If etype<>"80" Then %>
				document.all.dvIps.style.display ="";
				<% end if %>
				document.all.evalinfo.style.display ="";
		 	}else{
				<% If etype<>"80" Then %>
		 		document.all.dvIps.style.display ="none";
				<% end if %>
				document.all.evalinfo.style.display ="none";
			}
		}else if (sType=="b"){
			if (document.frmEvt.chBookingsell.checked){
				document.all.dvBS.style.display ="";
		 	}else{
		 		document.all.dvBS.style.display ="none";
			}	  
		}else if (sType=="c"){	
			if (document.frmEvt.chComm.checked){
				<% If etype<>"80" Then %>
				document.all.dvCmt.style.display ="";
				<% end if %>
				document.all.comminfo.style.display ="";
		 	}else{
				<% If etype<>"80" Then %>
		 		document.all.dvCmt.style.display ="none";
				<% end if %>
				document.all.comminfo.style.display ="none";
			}	 
		}
	}
	
	function popRegItem(eCode, gCode,eChannel){
	var wImgView;
	wImgView = window.open('eventitem_regist.asp?eC='+eCode+'&selG='+gCode+'&eCh='+eChannel,'pImg','width=1400,height=800,scrollbars=yes,resizable=yes');
	wImgView.focus();
}

	function jsAddByte(obj){ 
	var realText = obj.value; 
	 var textBit = '';
	 var textLen = 0;
	 
	 for (var i = 0 ; i < realText.length ; i++) {
	  textBit = realText.charAt(i); 
	  if(escape(textBit).length > 4) {
	   textLen = textLen + 2;
	  } else {
	   textLen = textLen + 1;
	  }
	  
	  if (textLen >= 70){
	    realText = realText.substr(0,i);
	    obj.value = realText;
	    break;
	  }
	  
	 }
	
    document.frmEvt.subSize.value = textLen;  

	}
	
	// ��α�URL�±� �˻�(�ڸ�Ʈ�� üũ�� �Ǿ��־�� ����)
	function jsChkBlogEnable() {
		if($('#isblogurl').prop('checked') == true) {
			if($('#chComm').prop('checked') == false) {
				alert("��α�URL����� �ڸ�Ʈ�� �־�߸� ��밡���մϴ�. �ڸ�Ʈ���θ� �������ּ���.");
				$('#isblogurl').prop('checked',false);
			}
		}
	}
	// ��ǰ���� ����Ʈ�˾�
	function jsItemcopylist(){
		var winLast,eKind;
		winLast = window.open('pop_event_itemlist.asp?menupos=<%=menupos%>&eC=<%=eCode%>','pLast','width=800,height=600, scrollbars=yes')
		winLast.focus();
	}
	
	
	function jsChkChannel(sChannel){ 
	    if (sChannel =="P"){
	        if(document.frmEvt.blnWeb.checked){
	            document.all.divPC1.style.display="";
	            document.all.divPC2.style.display="";
	        }else{
	            document.all.divPC1.style.display="none";
	            document.all.divPC2.style.display="none";
	        }
	    }
	    if (sChannel =="M" || sChannel =="A"){
	        if(document.frmEvt.blnMobile.checked || document.frmEvt.blnApp.checked){
	            document.all.divMA1.style.display="";
	            document.all.divMA2.style.display=""; 
	        }else{
	            document.all.divMA1.style.display="none";
	            document.all.divMA2.style.display="none"; 
	        }
	    }
	}
	 
	function jsChkSale(){
	    var frm = document.frmEvt; 
	    if(( frm.eventkind.options[frm.eventkind.selectedIndex].value == 1 ||  frm.eventkind.options[frm.eventkind.selectedIndex].value == 23 )   && (frm.chSale.checked)){
			document.all.spSale.style.color = "red";
			document.all.spSale.style.display = "";
	    }else{
	        document.frmEvt.sSP.value ="";
	        document.all.spSale.style.display = "none";
	    }
	}

	function jsChkCSale(){
	    var frm = document.frmEvt; 
	    if(( frm.eventkind.options[frm.eventkind.selectedIndex].value == 1 ||  frm.eventkind.options[frm.eventkind.selectedIndex].value == 23 )   && (frm.chCoupon.checked)){
			document.all.spCSale.style.color = "green";
			document.all.spCSale.style.display = "";
	    }else{
			document.frmEvt.sCSP.value ="";
			document.all.spCSale.style.display = "none";
	    }
	}
  
   function jsPubHelp(){ 
	   	var winPop = window.open("pop_publishing.asp","popHelp","width=500,height=500,scrollbars=yes,resizable=yes");
		winPop.focus();
	}    
	
    function jsChkMBReq(){ 
	    if(document.frmEvt.chkMB.checked){  
	         document.frmEvt.sWorkTag.value = "�ڡ�" + document.frmEvt.sWorkTag.value; 
	    }else{
	          document.frmEvt.sWorkTag.value =  document.frmEvt.sWorkTag.value.replace("�ڡ�", "");
	    }
	}
		// ��ǰ �ʱ�ȭ
	function jsItemclear(){
		var frm = document.frmitemclear;

		if(confirm("��ǰ �ʱ�ȭ�� �Ͻðڽ��ϱ�?\n\n��ǰ �ʱ�ȭ�� ������ ������ �Ұ��� �մϴ�.")){
			frm.target = "FrameCKP";
			//frm.target = "blank";
			frm.action = "/admin/eventmanage/event/event_process.asp";
			frm.submit();
		}
	}
	      
		function popCommentXLS(ecd) {
		 var wCmtXls = window.open('/admin/eventmanage/event/pop_event_Comment_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	//2015.05.19 ���¿�(Ǫ�������� �̺�Ʈ������ �ӽ� ����-�̺�Ʈ ������ ��������)
	function popCommentXLS2(ecd) {
		 var wCmtXls = window.open('/admin/eventmanage/event/pop_event_Comment_xls_2.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	function popBBSXLS(ecd) {
		 var wBBSXls = window.open('/admin/eventmanage/event/pop_event_board_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wBBSXls.focus();
	}
	   
	   
	function jsCmtStyle(sName){  
	    if (eval("document.frmEvt.chk"+sName).checked){ 
	         eval("document.frmEvt."+sName+"MT").value = ""; 
	         eval("document.frmEvt."+sName+"MT").className = "textarea_ro";
	         eval("document.frmEvt."+sName+"MT").disabled  = true;
	        if (sName =="eCmt" || sName == "eIps" ) {
	         eval("document.frmEvt."+sName+"ST").value = "";
	         eval("document.frmEvt."+sName+"ST").className = "textarea_ro";
	         eval("document.frmEvt."+sName+"ST").disabled  = true; 
	        }
	    }else{
	         eval("document.frmEvt."+sName+"MT").className = "textarea"; 
	         eval("document.frmEvt."+sName+"MT").disabled  = false;
	         if (sName =="eCmt" || sName == "eIps" ) {
             eval("document.frmEvt."+sName+"ST").disabled  = false; 
             eval("document.frmEvt."+sName+"ST").className = "textarea";
            }
	    } 
	}

	function jstagchk(v){
		var taglength = document.frmEvt.ietag.length;
		for ( i = 1 ; i<=taglength ; i++ )
		{
			if (v == "1" || v == "2" )
			{
				document.frmEvt.ietagval.readOnly = false;
			}else{
				document.frmEvt.ietagval.readOnly = true;
			}
		}
	}
	
	//�����̵� üũ
	function slidechk(d){
		if (d == "w"){
			if($('input:radio[name=slide_w_flag]').is(':checked'))
			{
				var winpop = window.open('/admin/eventmanage/event/v2/template/slide/pop_pcweb_slide.asp?eC=<%=eCode%>','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
				winpop.focus();
			}else{ alert("PCWEB �����̵��� üũ Ȯ��");}
		}else{
			if($('input:radio[name=slide_m_flag]').is(':checked'))
			{
				var winpop = window.open('/admin/eventmanage/event/v2/template/slide/pop_mobile_slide.asp?eC=<%=eCode%>','winpop','width=1200,height=850,scrollbars=yes,resizable=yes');
				winpop.focus();
			}else{ alert("MOBILE �����̵��� üũ Ȯ��");}
		}
	}
	
	//�̸�����
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes");
	    }
	}

	function poppcaddimg(){
		var winPopaddimg;
		winPopaddimg = window.open('/admin/eventmanage/event/v3/template/addbanner/pop_pc_addbanner.asp?eC=<%=eCode%>','pCal','width=1450,height=800,scrollbars=yes,resizable=yes');
		winPopaddimg.focus();
	}

	function popmoaddimg(){
		var winPopaddimg;
		winPopaddimg = window.open('/admin/eventmanage/event/v2/template/addbanner/pop_mobile_addbanner.asp?eC=<%=eCode%>','pCal','width=1450,height=800,scrollbars=yes,resizable=yes');
		winPopaddimg.focus();
	}

	// �̺�Ʈ ��ǰ �ִ� ������ ����
	function fnGetMaxSalevalue(saildiv) {
		var evtcd = document.frmEvt.eC.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetEvtMaxItemSalePer.asp",
			data: "eC="+evtcd+"&saildiv="+saildiv,
			cache: false,
			success: function(message) {
				if(message) {
					if(saildiv=="S"){
						document.frmEvt.sSP.value=message;
					}else{
						document.frmEvt.sCSP.value=message;
					}
					
				} else {
					alert("�̺�Ʈ�� ��ǰ�� ���ų� �������� ��ǰ�� �����ϴ�.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	function TnChecketcitemType(type){
		if (type =="C"){
            document.all.etcitemtype1.style.display="";
			document.all.etcitemtype2.style.display="none";
        }else{
			document.all.etcitemtype1.style.display="none";
            document.all.etcitemtype2.style.display="";
	    }
	}
	
	function TnSelectTheme(objval){
		if(objval!=1){
			document.all.bnimg.style.display="";
			document.all.colorselect2.style.display="";
			document.all.colorselect1.style.display="none";
		}else{
			document.all.bnimg.style.display="none";
			document.all.colorselect2.style.display="none";
			document.all.colorselect1.style.display="";
		}
		if(objval==2){
			document.all.textbgcolor.style.display="";
		}else{
			document.all.textbgcolor.style.display="none";
		}
		if(objval==2){
			//$('input:radio[name="bntype"][value="D"]').prop("checked",true);
			//$('input:radio[name="bntype"][value="T"]').attr("disabled",true);
			//$('input:radio[name="bntype"][value="D"]').attr("disabled",false);
			document.frmEvt.bntype.value="D";
			document.frmEvt.themetype.value="3";
		}else if(objval==3){
			//$('input:radio[name="bntype"][value="T"]').prop("checked",true);
			//$('input:radio[name="bntype"][value="D"]').attr("disabled",true);
			//$('input:radio[name="bntype"][value="T"]').attr("disabled",false);
			document.frmEvt.bntype.value="T";
		}else if(objval==4){
			//$('input:radio[name="bntype"][value="D"]').prop("checked",true);
			//$('input:radio[name="bntype"][value="T"]').attr("disabled",true);
			//$('input:radio[name="bntype"][value="D"]').attr("disabled",false);
			document.frmEvt.themetype.value="2";
		}
	}

	function TnSelectThemeMo(objval){
		if(objval!=1){
			document.all.bnimgmo.style.display="";
			document.all.mcolorselect2.style.display="";
			document.all.mcolorselect1.style.display="none";
		}else{
			document.all.bnimgmo.style.display="none";
			document.all.mcolorselect2.style.display="none";
			document.all.mcolorselect1.style.display="";
			document.all.videomo.style.display="none";
		}
		if(objval==2){
			document.all.textbgcolormo.style.display="";
			document.all.videomo.style.display="";
		}else{
			document.all.textbgcolormo.style.display="none";
			document.all.videomo.style.display="none";
		}
		if(objval==2){
			//$('input:radio[name="bntypemo"][value="D"]').prop("checked",true);
			//$('input:radio[name="bntypemo"][value="T"]').attr("disabled",true);
			//$('input:radio[name="bntypemo"][value="D"]').attr("disabled",false);
			document.frmEvt.bntypemo.value="D";
			document.frmEvt.themetypemo.value="3";
		}else if(objval==3){
			//$('input:radio[name="bntypemo"][value="T"]').prop("checked",true);
			//$('input:radio[name="bntypemo"][value="D"]').attr("disabled",true);
			//$('input:radio[name="bntypemo"][value="T"]').attr("disabled",false);
			document.frmEvt.bntypemo.value="T";
		}else if(objval==4){
			//$('input:radio[name="bntypemo"][value="D"]').prop("checked",true);
			//$('input:radio[name="bntypemo"][value="T"]').attr("disabled",true);
			//$('input:radio[name="bntypemo"][value="D"]').attr("disabled",false);
			document.frmEvt.themetypemo.value="2";
		}
	}

	//�׸� ��� ��� üũ
	/*function TnThemeBannerRegCheck(d){
		if (d == "w"){
			if($('input:radio[name=bntype]').is(':checked'))
			{
				if($(':radio[name="bntype"]:checked').val()=="T"){
					var winpop = window.open('/admin/eventmanage/event/V3/template/addbanner/pop_theme_addItems.asp?eC=<%=eCode%>&stype=w','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
				}else{
					var winpop = window.open('/admin/eventmanage/event/v3/template/slide/pop_pcweb_themeslide.asp?eC=<%=eCode%>&bgubun='+document.frmEvt.themetype.value,'winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
				}
				winpop.focus();
			}else{ alert("��� �̹��� ��� Ÿ�� üũ Ȯ��");}
		}else{
			if($('input:radio[name=bntypemo]').is(':checked'))
			{
				if($(':radio[name="bntypemo"]:checked').val()=="T"){
					var winpop = window.open('/admin/eventmanage/event/V3/template/addbanner/pop_theme_addItems.asp?eC=<%=eCode%>&stype=m','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
				}else{
					var winpop = window.open('/admin/eventmanage/event/v3/template/slide/pop_mobile_themeslide.asp?eC=<%=eCode%>&bgubun='+document.frmEvt.themetypemo.value,'winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
				}
				winpop.focus();
			}else{ alert("��� �̹��� ��� Ÿ�� üũ Ȯ��");}
		}
	}*/

	//�׸� ��� ��� üũ
	function TnThemeBannerRegCheck(d){
		if (d == "w"){
			if($('#bntype').val()=="T"){
				var winpop = window.open('/admin/eventmanage/event/V3/template/addbanner/pop_theme_addItems.asp?eC=<%=eCode%>&stype=w','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
			}else{
				var winpop = window.open('/admin/eventmanage/event/v3/template/slide/pop_pcweb_themeslide.asp?eC=<%=eCode%>&bgubun='+document.frmEvt.themetype.value,'winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
			}
			winpop.focus();
		}else{
			if($('#bntypemo').val()=="T"){
				var winpop = window.open('/admin/eventmanage/event/V3/template/addbanner/pop_theme_addItems.asp?eC=<%=eCode%>&stype=m','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
			}else{
				var winpop = window.open('/admin/eventmanage/event/v3/template/slide/pop_mobile_themeslide.asp?eC=<%=eCode%>&bgubun='+document.frmEvt.themetypemo.value,'winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
			}
			winpop.focus();
		}
	}

	//�����ڵ� ����
	function selColorChip(cd) {
		var i;
		frmEvt.DFcolorCD.value= cd;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//�����ڵ� ����
	function selColorChip2(cd) {
		var i;
		frmEvt.DFcolorCD.value= cd;
		for(i=1;i<=11;i++) {
			document.all("clinet"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("clinet0").bgColor='#DD3300';
		else document.all("clinet"+cd).bgColor='#DD3300';
	}

	//�����ڵ� ����
	function selBackGroundColorChip(cd) {
		var i;
		frmEvt.DFcolorCD2.value= cd;
		for(i=1;i<=2;i++) {
			document.all("bcline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("bcline0").bgColor='#DD3300';
		else document.all("bcline"+cd).bgColor='#DD3300';
	}

	//�����ڵ� ����
	function selMoColorChip(cd) {
		var i;
		frmEvt.DFcolorCDMo.value= cd;
		for(i=1;i<=11;i++) {
			document.all("mocline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("mocline0").bgColor='#DD3300';
		else document.all("mocline"+cd).bgColor='#DD3300';
	}

	//�����ڵ� ����
	function selMoColorChip2(cd) {
		var i;
		frmEvt.DFcolorCDMo.value= cd;
		for(i=1;i<=11;i++) {
			document.all("moclinet"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("moclinet0").bgColor='#DD3300';
		else document.all("moclinet"+cd).bgColor='#DD3300';
	}

	//�����ڵ� ����
	function selMoBackGroundColorChip(cd) {
		var i;
		frmEvt.DFcolorCDMo2.value= cd;
		for(i=1;i<=2;i++) {
			document.all("mobcline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("mobcline0").bgColor='#DD3300';
		else document.all("mobcline"+cd).bgColor='#DD3300';
	}

	function TnThemeItemBannerDel(idx){
		document.ibfrm.target="ifrmProc";
	    document.ibfrm.idx.value=idx;
	    document.ibfrm.action="/admin/eventmanage/event/v3/template/addbanner/delitembanner.asp";
	    document.ibfrm.submit();
	}

	function TnDelSlideBanner(idx){
		document.ibfrm.target="ifrmProc";
	    document.ibfrm.idx.value=idx;
	    document.ibfrm.action="/admin/eventmanage/event/v3/template/slide/delslidebanner.asp";
	    document.ibfrm.submit();
	}

	function TnGiftUsingNum(objval){
		if (objval == "1"){
            document.all.gifttxt1.style.display="";
			document.all.giftimg1.style.display="";
			document.all.gifttxt2.style.display="none";
			document.all.giftimg2.style.display="none";
			document.all.gifttxt3.style.display="none";
			document.all.giftimg3.style.display="none";
        }else if (objval == "2"){
            document.all.gifttxt1.style.display="";
			document.all.giftimg1.style.display="";
			document.all.gifttxt2.style.display="";
			document.all.giftimg2.style.display="";
			document.all.gifttxt3.style.display="none";
			document.all.giftimg3.style.display="none";
        }else if (objval == "3"){
            document.all.gifttxt1.style.display="";
			document.all.giftimg1.style.display="";
			document.all.gifttxt2.style.display="";
			document.all.giftimg2.style.display="";
			document.all.gifttxt3.style.display="";
			document.all.giftimg3.style.display="";
        }else{
			document.all.gifttxt1.style.display="none";
			document.all.giftimg1.style.display="none";
			document.all.gifttxt2.style.display="none";
			document.all.giftimg2.style.display="none";
			document.all.gifttxt3.style.display="none";
			document.all.giftimg3.style.display="none";
	    }
	}

	function TnUsingInfoNum(objval){
		if (objval == "1"){
            document.all.usingtxt1.style.display="";
			document.all.usingcont1.style.display="";
			document.all.usingtxt2.style.display="none";
			document.all.usingcont2.style.display="none";
			document.all.usingtxt3.style.display="none";
			document.all.usingcont3.style.display="none";
        }else if (objval == "2"){
            document.all.usingtxt1.style.display="";
			document.all.usingcont1.style.display="";
			document.all.usingtxt2.style.display="";
			document.all.usingcont2.style.display="";
			document.all.usingtxt3.style.display="none";
			document.all.usingcont3.style.display="none";
        }else if (objval == "3"){
            document.all.usingtxt1.style.display="";
			document.all.usingcont1.style.display="";
			document.all.usingtxt2.style.display="";
			document.all.usingcont2.style.display="";
			document.all.usingtxt3.style.display="";
			document.all.usingcont3.style.display="";
        }else{
			document.all.usingtxt1.style.display="none";
			document.all.usingcont1.style.display="none";
			document.all.usingtxt2.style.display="none";
			document.all.usingcont2.style.display="none";
			document.all.usingtxt3.style.display="none";
			document.all.usingcont3.style.display="none";
	    }
	}

	function TnFavSearchTxt(){
		var winpop = window.open("http://61.252.133.17:5601/app/kibana#/dashboard/5c9d9970-ef60-11e6-9fb4-f3d99fd9206d?_g=(refreshInterval:(display:Off,pause:!f,value:0),time:(from:now-5h%2Fh,mode:quick,to:now))&_a=(filters:!(),options:(darkTheme:!f),panels:!((col:1,id:ca566510-ef5f-11e6-9fb4-f3d99fd9206d,panelIndex:1,row:1,size_x:3,size_y:5,type:visualization),(col:1,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4(MOB)',panelIndex:2,row:6,size_x:3,size_y:5,type:visualization),(col:1,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4(APP)',panelIndex:3,row:11,size_x:3,size_y:5,type:visualization),(col:4,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4-%EC%8B%9C%EA%B0%84%EB%8C%80%EB%B3%84(MOB)',panelIndex:4,row:6,size_x:9,size_y:5,type:visualization),(col:4,id:d06ee1e0-ef62-11e6-9fb4-f3d99fd9206d,panelIndex:5,row:1,size_x:9,size_y:5,type:visualization),(col:4,id:c7604a10-1aa2-11e7-b3b2-cb4977e75f0e,panelIndex:6,row:11,size_x:9,size_y:5,type:visualization)),query:(query_string:(analyze_wildcard:!t,query:'*')),title:'0005.%20%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4',uiState:(P-1:(vis:(params:(sort:(columnIndex:!n,direction:!n)))),P-2:(vis:(params:(sort:(columnIndex:!n,direction:!n)))),P-3:(vis:(params:(sort:(columnIndex:!n,direction:!n))))))",'winpop2','width=1450,height=800,scrollbars=yes,resizable=yes');
		winpop.focus();
	}

	function TnReloadThisPage(){
		location.reload();
	}

	$(document).ready(function(){
		$('#nocate').on('click',function(){
			if($("#nocate").is(":checked")){
				$("#disp1").attr("disabled", true);
				$("#disp2").attr("disabled", true);
			}else{
				$("#disp1").attr("disabled", false);
				$("#disp2").attr("disabled", false);
			}
		});
		$('#eventtype').on('change',function(){
			var etype="<%=etype%>";
			var objval=$('#eventtype').val();
			if(etype!=80 && objval==80){
				if(confirm("MD ����� ���� ������ �������� ���� �����ʹ� �Ҹ� �˴ϴ�.\n������ ���� �ϰڽ��ϱ�?")){
					location.href="/admin/eventmanage/event/v3/event_modify.asp?eC=<%=eCode%>&menupos=1739&eventtype=80";
				}
			}else if(etype==80 && objval!=80){
				if(confirm("MD ����� ���� ������ �������� ���� �����ʹ� �Ҹ� �˴ϴ�.\n������ ���� �ϰڽ��ϱ�?")){
					location.href="/admin/eventmanage/event/v3/event_modify.asp?eC=<%=eCode%>&menupos=1739&eventtype="+objval;
				}
			}
		});
		$('#sEN').keyup(function(){
			$("#title_pc").val($("#sEN").val());
			$("#title_mo").val($("#sEN").val());
		});
	});

$(document).ready(function() {
	//���ڸ� üũ
	$("#sSP").keyup(function(){
		/* 48~57:�Ϲ� ����Ű �ڵ�, 96~105:����Ű�е� ����Ű �ڵ� , 8 : backspace, 9 : tab ,46 : delete, 37 : ���ʹ���Ű , 39 : �����ʹ���Ű */
		var keyID = event.which;

		if ((keyID >= 48 && keyID <= 57) || (keyID >= 96 && keyID <= 105) || keyID == 8 || keyID == 9 || keyID == 46 || keyID == 37 || keyID == 39) {
			if (this.value > 100 || this.value < 0) {
				alert("0���� 100 ������ ���� �Է����ּ���");
				this.value = ''; //������ �����ִ� �� �� �����ش�.
			};
		}
		else {
			alert("���ڸ� �Է� �����մϴ�.");
			this.value = this.value.replace(/[^0-9\.]/g, ''); //���ڸ� ������ ���ڸ� �����ش�.
		}
	});
	$("#sCSP").keyup(function(){
		/* 48~57:�Ϲ� ����Ű �ڵ�, 96~105:����Ű�е� ����Ű �ڵ� , 8 : backspace, 9 : tab ,46 : delete, 37 : ���ʹ���Ű , 39 : �����ʹ���Ű */
		var keyID = event.which;

		if ((keyID >= 48 && keyID <= 57) || (keyID >= 96 && keyID <= 105) || keyID == 8 || keyID == 9 || keyID == 46 || keyID == 37 || keyID == 39) {
			if (this.value > 100 || this.value < 0) {
				alert("0���� 100 ������ ���� �Է����ּ���");
				this.value = ''; //������ �����ִ� �� �� �����ش�.
			};
		}
		else {
			alert("���ڸ� �Է� �����մϴ�.");
			this.value = this.value.replace(/[^0-9\.]/g, ''); //���ڸ� ������ ���ڸ� �����ش�.
		}
	});
});
</script>
<form name="frmitemclear" method="post">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="IC">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<form name="frmG" method="post" action="eventgroup_process.asp">
  <input type="hidden" name="menupos" value="<%=menupos%>">  
  <input type="hidden" name="eC" value="<%=eCode%>">
  <input type="hidden" name="mode" value="">
  <input type="hidden" name="eCh" value="">
  <input type="hidden" name="eMT" value="">
</form>
<form name="frmEvt" method="post" action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="U">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="strparm" value="<%=strparm%>">  
<input type="hidden" name="banMoList" value="<%=ebimgMo2014%>">
<input type="hidden" name="icon" value="<%=eicon%>"> 
<input type="hidden" name="gift" value="<%=gimg%>"> 
<input type="hidden" name="etcitemban" value="<%=eEtcitemimg%>">
<input type="hidden" name="DFcolorCD" value="<%=themecolor%>">
<input type="hidden" name="DFcolorCD2" value="<%=textbgcolor%>">
<input type="hidden" name="DFcolorCDMo" value="<%=themecolormo%>">
<input type="hidden" name="DFcolorCDMo2" value="<%=textbgcolormo%>">
<input type="hidden" name="themetype" value="3">
<input type="hidden" name="themetypemo" value="3">
<input type="hidden" name="upback" value="N">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" > 
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ �⺻ ����</B></td>
		   	</tr>
			 <tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
		   			<tr>
		   				<td>
							<%=eCode%>
							<input type="button" value="��ǰ ����" onclick="jsItemcopylist();" class="button"/>
							<input type='button' value='��ǰ�ʱ�ȭ' onclick='jsItemclear();' class='button' /> 
						</td>
		   				<td>
						<%
							'�̺�Ʈ ������ ���� ����Ʈ��ũ ������ ����
							Select Case ekind
								Case "7"		'��Ŭ���ڵ�
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case "13"		'��ǰ �̺�Ʈ
								    Response.Write "<td> �̸�����:" 
									Response.Write "<a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & eitemid & "' target='_blank'>PC-Web</a>"
									Response.Write "&nbsp;<a href= ""javascript:jsOpen('" & vmobileUrl & "/category/category_itemPrd.asp?itemid=" & eitemid & "','M');"">Mobile</a>" 
									Response.Write"</td>"
								Case "14"		'��ǳ���±�
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case "16"		'�귣�� �������
									Response.Write "<td><a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & ebrand & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>�̸�����</a></td>"
								Case "22"		'DAY&(���̾ص�)
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case "26"		'�����
									Response.Write "<td><a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'>�̸�����</a></td>"
								Case Else		'�������� �� ��Ÿ
								    Response.Write "<td> �̸�����:" 
									Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'> PC-Web</a>" 
									Response.Write "&nbsp;<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "','M');"">Mobile</a>" 
									Response.Write "&nbsp;<a href= ""javascript:jsOpen('" & vmobileUrl & "/apps/appCom/wish/web2014/event/eventmain.asp?eventid=" & eCode & "','M');"">App</a>"
								  
									Response.Write"</td>"
							End Select

							'//�ν�Ÿ�׷� ���� ��ư (�����ø��ϴ�)
							If session("ssBctId") = "motions" Or session("ssBctId") = "greenteenz" Or session("ssBctId") = "bjh2546" Or session("ssBctId") = "djjung" Or session("ssBctId") = "ksy92630" Or session("ssBctId") = "ppono2" Or session("ssBctId") = "thensi7" Or session("ssBctId") = "corpse2"  Or session("ssBctId") = "aasd1991" Then
								Response.write "<td><a href=""/admin/etc/only_sys/10x10instagram.asp?eventid="&eCode&""" target='_blank'>10x10instagram</a></td>"
							End If 
						%>
		   				</td>
		   				<td align="right">
		   				<% If sKind = "2" Then %>
		   					<input type="button" value="�Ѹ���List" onClick="window.open('/admin/eventmanage/oneline/?eC=<%=eCode%>&esday=<%=esday%>','oneline','width=600,height=500,scrollbars=yes');">
		   					<img src="/images/icon_excel_reply.gif" alt="�ڸ�Ʈ ������ Excel�ٿ�ε�" onClick="location.href='/admin/eventmanage/oneline/oneline_excel.asp?eC=<%=eCode%>&esday=<%=esday%>';" style="cursor:pointer" align="absmiddle">
		   				<% Else %>
		   					<img src="/images/icon_excel_reply.gif" alt="�ڸ�Ʈ ������ Excel�ٿ�ε�" onClick="popCommentXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   					<img src="/images/icon_excel_bbs.gif" alt="�Խ��� ������ Excel�ٿ�ε�" onClick="popBBSXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   				<% End If %>
		   					<img src="/images/icon_excel_vote.gif" alt="���� ������ Excel�ٿ�ε�" onClick="window.open('/admin/eventmanage/event/pop_event_votelist_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle" title ="xls �ٿ�ε� ȸ�����">
							<img src="/images/icon_excel_vote.gif" alt="���� ������ Excel�ٿ�ε� ��ȸ��"  title ="xls �ٿ�ε� ��ȸ��" onClick="window.open('/admin/eventmanage/event/pop_event_votelist_guest_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">
							<img src="/images/icon_excel_vote.gif" alt="���� ������ Excel�ٿ�ε� Lite����"  title ="xls �ٿ�ε� Lite����" onClick="window.open('/admin/eventmanage/event/pop_event_votelist_lite_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">
		   				</td>
		   			</tr>
		   			</table>
		   		</td>
		   	</tr>
		    <tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�������</td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="radio" name="using" value="Y" <%IF eusing="Y" THEN%>checked<%END IF%>>Yes <input type="radio" name="using" value="N" <%IF eusing="N" THEN%>checked<%END IF%>>No
		   		</td>
		   	</tr> 
			<tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">ä��<B style="color:red">*</B></td>
		   		<td bgcolor="#FFFFFF">
		   			 <label><input type="checkbox" name="blnWeb" value="1" <%IF isWeb THEN%>checked<%END IF%> onClick="jsChkChannel('P');"> PC-Web</label>
		   			 <label><input type="checkbox" name="blnMobile" value="1" <%IF  isMobile  THEN%>checked<%END IF%> onClick="jsChkChannel('M');"> Mobile</label>
		   			 <label><input type="checkbox" name="blnApp" value="1"  <%IF  isApp  THEN%>checked<%END IF%> onClick="jsChkChannel('A');"> APP</label>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��ü</td>
		   		<td bgcolor="#FFFFFF">
		   		 	 <%=fnGetCommCodeArrDesc(arreventmanager,eman)%> 
		   		 	 <input type="hidden" name="eventmanager" value="<%=eman%>">
		   		</td>
		   	</tr> 
		   <tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">����<B style="color:red">*</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventkind",ekind,False,"onChange=javascript:jsChkSubj(this.value);"%>
		   			<% If ekind = "29" Then %>
		   			<strong> �� <font color="blue" size="3">���� �� �ڵ� �۾��� �ִ� ���</font> <font color="red" size="3">�ݵ�� �۾��ڸ� �����ؾ��մϴ�.</font></strong>
		   			<% End If %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ Ÿ��<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">  
		   		<input type="checkbox" name="chSale" <%IF esale   THEN%>checked<%END IF%> value="1"  onClick="jsChkSale();">����
		   		<input type="checkbox" name="chGift" <%IF egift  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('g');">����ǰ
		   		<input type="checkbox" name="chCoupon" <%IF ecoupon   THEN%>checked<%END IF%> value="1" onClick="jsChkCSale();">����
		   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten   THEN%>checked<%END IF%> value="1">Only-TenByTen
		   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone  THEN%>checked<%END IF%> value="1">1+1
				<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery   THEN%>checked<%END IF%> value="1">������
				<input type="checkbox" name="chBookingsell" <%IF eBookingsell  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('b');">�����Ǹ�
				<input type="checkbox" name="chDiary" <%IF eDiary  THEN%>checked<%END IF%> value="1">DiaryStory
				<input type="checkbox" name="chNew" <%IF eNew  THEN%>checked<%END IF%> value="1">��Ī
		   		</td>
			</tr>
		   	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ���<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chComm" <%IF ecomment   THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('c');">�ڸ�Ʈ
		   		<input type="checkbox" name="chBbs" <%IF ebbs   THEN%>checked<%END IF%> value="1" >�Խ���
		   		<input type="checkbox" name="chItemps" <%IF eitemps  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('i');">��ǰ�ı�
		   		<input type="checkbox" name="isblogurl" <%IF eisblogurl THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">Blog URL
		   		<!--<input type="checkbox" name="chApply" <%IF eapply = 1 THEN%>checked<%END IF%> value="1" >����-->
		   		</td>
		   </tr>
		   <tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ����<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventtype",etype,True,""%>
		   			<a href="#" onclick="window.open('popEventType.html','popViewEvtType','width=550,height=480, scrollbars=yes');return false;" style="margin-left:10px;color:#A38;">[�̺�Ʈ ��������]</a>
		   			<span id="lyrEvtConfirm" style="<%=chkIIF(etype="50","","display:none;")%>margin-left:10px;">
		   			<% if isConfirm then %>
		   				<input type="hidden" name="blnCnfm" value="1">
		   				<font color="#AA2244">�� �̺�Ʈ ������ ���εǾ����ϴ�.</font>
		   			<% else %>
		   				<label><input type="checkbox" name="blnCnfm" value="1" <%=chkIIF(session("ssAdminLsn")<="3","","readonly")%>> �̺�Ʈ ���� ����</label>
		   			<% end if %>
		   			</span>
		   		</td>
		   </tr>
		   <tr id="eNameTr_A" style="display:<% if ekind="16"  then Response.Write "none" %>;">
		   		<td bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" id="sEN" maxlength="120" size="60" value="<%=ename%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:<% if ekind<>"16"  then Response.Write "none" %>;">
		   	<%
		   		'// �귣�������̸� ������ �������� ǥ��
		   		dim arrEname
				arrEname = Split(ename,"|")
				if Ubound(arrEname)<2 then
					arrEname = ename & "|0|0"
					arrEname = Split(arrEname,"|")
				end if

				If enameEng <> "" then
					Dim arrEnameEng
					arrEnameEng = Split(enameEng,"|")
					if Ubound(arrEnameEng)<2 then
						arrEnameEng = enameEng & "|0|0"
						arrEnameEng = Split(arrEnameEng,"|")
					end If
				End If
		   	%>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��<br>�� ������<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
					�̺�Ʈ��: <input type="text" name="sEDN" size="60" maxlength="60" value="<%=arrEname(0)%>"><br>
					<% If enameEng <> "" Then %>
		   			�����̺�Ʈ��: <input type="text" name="sEDNEng" size="60" maxlength="60" value="<%=arrEnameEng(0)%>"><br>
					<% End If %>
		   			������: ���� <input type="text" name="sSDc" size="4" value="<%=arrEname(1)%>" style="text-align:right;">% ~
		   			�ְ� <input type="text" name="sMDc" size="4" value="<%=arrEname(2)%>" style="text-align:right;">%<br>
		   			<font color=gray>�غ귣�� ��Ʈ��Ʈ�� ������ �������Դϴ�. ������ ��ǰ���� ������� ������ ��ǰ���� ���� ������ �������ּ���.
		   		<br>�̺�Ʈ ��ũ�� �귣�� ��Ʈ��Ʈ�� ����Ǵ� �ݵ�� �󼼳��뿡 �귣�带 �������ּ���.</font> 
		   		</td>
		   	</tr>
		   	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�Ⱓ<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   		<%IF estate = 9 THEN%>
		   			������ : <%=esday%><input type="hidden" name="sSD" size="10" value="<%=esday%>">
		   			~ ������ : <%=eeday%> <input type="hidden" name="sED" value="<%=eeday%>" size="10" >
		   		<%ELSE%>
		   			������ : <input type="text" id="termSdt" name="sSD" size="10" value="<%=esday%>" />
							<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkStart_trigger" onclick="return false;" />
		   			~ ������ : <input type="text" id="termEdt" name="sED" size="10" value="<%=eeday%>" />
							<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "termSdt", trigger    : "ChkStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(eeday<>"",", max: " & replace(eeday,"-",""),"")%>
						});
						var CAL_End = new Calendar({
							inputField : "termEdt", trigger    : "ChkEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(eeday<>"",", min: " & replace(esday,"-",""),"")%>
						});
					</script>
		   		<%END IF%>
				&nbsp;&nbsp;<input type="checkbox" name="dateview"  value="1"  <%IF eDateView THEN%>checked<%END IF%>> �̺�Ʈ �Ⱓ ���� ����
				&nbsp;&nbsp;<input type="checkbox" name="endlessview"  value="Y"  <%IF endlessView="Y" THEN%>checked<%END IF%>> <a title="��ó��� ������ �Ⱓ�� ���� �̺�Ʈ�� �̺�Ʈ ���� �ȳ� ���̾� �ȶߵ��� ����">��ó���</a>
		   		</td>
		   	</tr>
		   	<tr>
				<td bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��<b style="color:red">*</b></td>
		   		<td  bgcolor="#FFFFFF">
		   			<input type="text" id="priceDt" name="sPD" size="10" value="<%=epday%>" />
					<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkPrc_trigger" onclick="return false;" />
					(��÷�ڰ� �ִ� ��쿡�� ���)
					<script type="text/javascript">
						var CAL_Prcdt = new Calendar({
							inputField : "priceDt", trigger    : "ChkPrc_trigger",
							onSelect: function() {
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
		   		</td>
		   </tr>
		   <tr id="spSale" style="display:<%if not ((ekind="1" or ekind="23") and (esale)) then %>none<%end if%>;<%if esale then%>color:red;<%end if%>">
		   		<td bgcolor="<%= adminColor("tabletop") %>"><font color="black">��ǰ ������</font><b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			    </font>~[<input type="text" name="sSP" id="sSP" value="<%=salePer%>" size="10" class="text" />]%&nbsp;&nbsp;&nbsp;<input type="button" class="button" value="�ִ밪 ��������" onclick="fnGetMaxSalevalue('S')" /> (���ڸ� �Է� ���� Ư������ �ڵ��Է�)
		   			    
		   		</td>
		   	</tr>
			<tr id="spCSale" style="display:<%if not ((ekind="1" or ekind="23") and (ecoupon)) then %>none<%end if%>;<%if ecoupon then%>color:green;<%end if%>">
		   		<td bgcolor="<%= adminColor("tabletop") %>"><font color="black">���� ������</font><b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			    </font>~[<input type="text" name="sCSP" id="sCSP" value="<%=saleCPer%>" size="10" class="text" />]&nbsp;&nbsp;&nbsp;<input type="button" class="button" value="�ִ밪 ��������" onclick="fnGetMaxSalevalue('C')" /> (���ڸ� �Է� ���� Ư������ �ڵ��Է�)
		   		</td>
		   	</tr>
		</table>
	</td>
</tr>
<tr>
	<td> 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ ���� ����</B></td>
		   	</tr>
		   	<tr>
		   		<td width="100"  bgcolor="<%= adminColor("tabletop") %>">ī�װ�<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   		<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->&nbsp;<input type="checkbox" name="nocate" id="nocate" value="Y"<% If nocate="Y" Then Response.write " checked"%>>ī�װ� ���� ����
		   		</td>
		   	</tr>
		   <tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�귣�� �ٷΰ���</td>
		   		<td bgcolor="#FFFFFF">
		   			<% drawSelectBoxDesignerwithName "ebrand", ebrand %>
		   		</td>
		   	</tr>
		    
		   	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ���Ĺ��<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "itemsort",eisort,False,""%>
		   		</td>
		   	</tr>
			<tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">Tag</td>
		   		<td bgcolor="#FFFFFF">
		   			(250�� �̳�)		   			<Br>
		   			<textarea name="eTag" rows="2" style="width:80%;"><%=etag%></textarea>&nbsp;<input type="button" class="button" value="�α� �˻��� ����" onclick="TnFavSearchTxt()" />
		   		</td>
		   	</tr>
		  	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">����<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<%   sbGetOptStatusCodeSort "eventstate",estate,false,"" 
		   				''if ekind="22" then
		   				''	'//���̾ص�� ��������Ʈ�� ����ؼ� �������
		   				''	sbGetOptStatusCodeValue "eventstate",estate,false,""
		   				''else
		   				''	sbGetOptStatusCodeAuth "eventstate",estate,"M",""
		   				''end if
		   			%>
		   			<input type="hidden" name="eOD" value="<%=dopendate%>">
		   			<input type="hidden" name="eCD" value="<%=dclosedate%>">
		   			<input type="hidden" name="eIRD" value="<%=dImgregdate%>">
		   			<%IF not isnull(dopendate) THEN%><span style="padding-left:10px;">  ����ó���� : <%=dopendate%>  </span><%END IF%>
		   			<%IF not isnull(dclosedate) THEN%>/ <span style="padding-left:10px;">  ����ó���� : <%=dclosedate%>  </span><%END IF%>
		   			<%IF not isnull(dImgregdate) THEN%>/ <span style="padding-left:10px;">  �̹�����Ͽ�û�� : <%=dImgregdate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�߿䵵<b style="color:red">*</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
		   		</td>
		   	</tr>
			<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">���� �׸� ����</td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="checkbox" name="chkItemprice"  value="1"  <%IF blnitemprice THEN%>checked<%END IF%>>����/���ΰ��� ����(üũ�� �������)&nbsp;&nbsp;
					<input type="checkbox" name="chkIteminfo"  value="1"  <%IF blnIteminfo THEN%>checked<%END IF%>>��ǰ����
		   		</td>
		   	</tr>
		</table>
		<div id="divDE" style="display:none;"> 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">���Ĺ�ȣ</td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sortNo" value="0" size="6" maxlength="5" style="text-align:right;" />
		   			(�ؼ��ڰ� Ŭ���� �켱ǥ�� �˴ϴ�. / Day&:ȸ��)
		   		</td>
		   	</tr> 
		</table>
		</div>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>����� ����</B></td>
		   	</tr>
		   	<!-- <tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ۺ���</td>
		   		<td bgcolor="#FFFFFF"><input type="checkbox" name="chkReqP" value="1" <%if blnReqPublish then%>checked<%end if%>>  �ۺ��� ��û <img src="/images/admin_help.gif" style="cursor:hand;" onClick="jsPubHelp();"></td>
		   	</tr> -->
			<tr>
				<td bgcolor="<%= adminColor("tabletop") %>" width="100">��ȹ��</td>
				<td bgcolor="#FFFFFF">
					<input type="hidden" name="sMdId" value="<%=emdid%>">
					<input type="name" name="sMdNm" value="<%=eMDnm%>"class="text_ro" readonly size="10">
					<input type="button" class="button" value="����" onClick="jsGetID('Md','<%=idepartmentid%>','<%=emdid%>');">
					<input type="button" value="&times"  class="button" onClick="jsDelID('Md');" title="����� �����" />
				</td>
			</tr>
			<tr>
				<td bgcolor="<%= adminColor("tabletop") %>">�����̳�(PC)</td>
				<td bgcolor="#FFFFFF">
					<%sbGetDesignerid "sDgId",edgid,"onchange=""jsChangeFrm(this.value,'DG1');"""%>
					<%sbGetOptEventCodeValue "designerstatus",edgstat1,True,""%>
				</td>
			</tr>
			<tr>
				<td bgcolor="<%= adminColor("tabletop") %>">�����̳�(Mobile)</td>
				<td bgcolor="#FFFFFF">
					<%sbGetDesignerid "sDgId2",edgid2,"onchange=""jsChangeFrm(this.value,'DG2');"""%>
					<%sbGetOptEventCodeValue "designerstatus",edgstat2,True,""%>
				</td>
			</tr>
			<tr>
				<td bgcolor="<%= adminColor("tabletop") %>">�ۺ���</td>
				<td bgcolor="#FFFFFF">
					<input type="hidden" name="sPsId" value="<%=epsid%>">
					<input type="name" name="sPsNm" value="<%=epsnm%>"class="text_ro" readonly size="10">
					<input type="button" class="button" value="����"  onClick="jsGetID('Ps','157','<%=epsid%>');">
					<input type="button" value="&times"  class="button" onClick="jsDelID('Ps');" title="����� �����" />
					&nbsp;&nbsp;<input type="checkbox" name="chkReqP" value="1" <%if blnReqPublish then%>checked<%end if%>>  �ۺ��� ��û <img src="/images/admin_help.gif" style="cursor:hand;" onClick="jsPubHelp();">
				</td>
			</tr>
			<tr>
				<td bgcolor="<%= adminColor("tabletop") %>">������</td>
				<td bgcolor="#FFFFFF">
					<input type="hidden" name="sDpId" value="<%=edpid%>">
					<input type="name" name="sDpNm" value="<%=edpnm%>" class="text_ro" readonly size="10">
					<input type="button" class="button" value="����" onClick="jsGetID('Dp','130','<%=edpid%>');">
					<input type="button" value="&times"  class="button" onClick="jsDelID('Dp');" title="����� �����" />
				</td>
			</tr>
			<!-- <tr>
				<td>���߰˼���</td>
				<td>
					<input type="hidden" name="sCCId" value="">
					<input type="name" name="sCCNm" value="" class="text_ro" readonly size="10">
					<input type="button" class="button" value="����" onClick="jsGetID('CC','130','<%=edpid%>');">
					<input type="button" value="&times"  class="button" onClick="jsDelID('CC');" title="����� �����" />
				</td>
			</tr> -->
		   	<tr>
		   		<td bgcolor="<%= adminColor("tabletop") %>">�����̳� �۾�����</td>
		   		<td bgcolor="#FFFFFF"><input type="text" name="sWorkTag" size="20" maxlength="16" class="text" value="<%= sWorkTag %>">
		   		    <input type="checkbox" name="chkMB"  onClick="jsChkMBReq();" <%if left(sWorkTag,4) ="�ڡ�" then%>checked<%end if%>> ����Ϲ�� ��û�� üũ   
		   		    </td>
		   	</tr> 
		 <!--����  2015.02.05
		 	<tr>
		   		<td width="100" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
		   		<td bgcolor="#FFFFFF">
		   			(200�� �̳�)		   			<Br>
		   			<textarea name="eCT" rows="2" style="width:100%;"></textarea>
		   		</td>
		   	</tr>-->
		</table> 
	</td>
</tr> 
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>��� ��� �̹��� ����</B></td>
		   	</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�⺻ ���<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
					<input type="radio" name="etcitemtype" value="C" onclick="TnChecketcitemType('C');" <%IF eEtcitemid<>"0" THEN%>checked<%END IF%>>��ǥ��ǰ�ڵ� <input type="radio" name="etcitemtype" value="I" onclick="TnChecketcitemType('I');" <%IF eEtcitemimg<>"" THEN%>checked<%END IF%>>�̹��� ���ε�
		   			<div id="etcitemtype1" style="display:<%IF eEtcitemid<>"0" THEN%><% Else %>none<%END IF%>"><input type="text" name="etcitemid" value="<%=eEtcitemid%>"></div>
					<div id="etcitemtype2" style="display:<%IF eEtcitemimg<>"" THEN%><% Else %>none<%END IF%>"><input type="button" name="etcitem" value="��ǰ��ǥ���" onClick="jsSetImg('<%=eFolder%>','<%=eEtcitemimg%>','etcitemban','etciitem')" class="button"></div>
					<div id="etciitem" style="padding: 5 5 5 5">
						<%IF eEtcitemimg <> "" THEN %>
						<img  src="<%=eEtcitemimg%>" border="0">
						<a href="javascript:jsDelImg('etcitemban','etciitem');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">���̵� ���</td>
				<td bgcolor="#FFFFFF">
				<input type="button" name="btnMoBan2014" value="���̵� ���" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo2014%>','banMoList','spanbanMoList')" class="button">
					 <div id="spanbanMoList" style="padding: 5 5 5 5">
						<%IF ebimgMo2014 <> "" THEN %>
						<img  src="<%=ebimgMo2014%>" border="0">
						<a href="javascript:jsDelImg('banMoList','spanbanMoList');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
					<p style="color:#602030;font-size:11px;">[ ���� �̹��� : JPEG, 60%, 750px �� 406px ]</p>
				</td>
			</tr>
			<!-- <tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ �̹��� </td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="btnicon" value="����ǰ�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=gimg%>','gift','spangift')" class="button">
					<div id="spangift" style="padding: 5 5 5 5">
						<%IF gimg <> "" THEN %>
						<a href="javascript:jsImgView('<%=gimg%>')"><img  src="<%=gimg%>" width="400" border="0"></a>
						<a href="javascript:jsDelImg('gift','spangift');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr> -->
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">PC ������</td>
				<td bgcolor="#FFFFFF">
					<input type="button" value="PC �߰� ���(<%=chkiif(evt_pc_addimg_cnt<>"",evt_pc_addimg_cnt,"0")%>)" onClick="poppcaddimg();" class="button"> (�ִ� 5��)
				</td>
			</tr>  
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����� ������</td>
				<td bgcolor="#FFFFFF">
					<input type="button" value="����� �߰� ���(<%=chkiif(evt_m_addimg_cnt<>"",evt_m_addimg_cnt,"0")%>)" onClick="popmoaddimg();" class="button"> (�ִ� 5��)
				</td>
			</tr>  
		</table> 
	</td>
</tr>
</table>
<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0">
			<tr>
				<td valign="top">
				    <div id="divPC1" style="display:<%if not isweb then%>none<%end if%>;">
					<% If etype<>"80" Then %>
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="evt_type1" style="display:<% If etype="80" Then %>none<% End If %>">
						<tr>
							<td bgcolor="#FAECC5" colspan="2"><b>PC �̺�Ʈ ���� ����</b></td>
						</tr>
						<tr>
					   		<td width="100"  bgcolor="#FAECC5">ȭ�鱸��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="radio" name="chkFull" value="0" <%IF not blnFull and not blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkWide.checked=false;"> �⺻��&nbsp;&nbsp;
					   			<input type="radio" name="chkFull" value="1" <%IF  blnFull and not blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkWide.checked=false;"> Ǯ��&nbsp;&nbsp;
					   			<input type="radio" name="chkWide" value="1" <%IF blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkFull[0].checked=false;chkFull[1].checked=false;"> ���̵� 
					   		</td>
						</tr>
						<tr> 
							<td bgcolor="#FAECC5">����ī��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">  
					   			<textarea name="subcopyK" style="width:100%; height:80px;" value="<%=subcopyK%>"><%=chkiif(subcopyK="","",subcopyK)%></textarea>
					   			<!-- <textarea name="subcopyE" style="width:100%; height:80px;" onclick="if(this.value=='����')this.value='';" onblur="if(this.value=='')this.value='����';" value="<%=subcopyE%>"><%=chkiif(subcopyE="","����",subcopyE)%></textarea> -->
					   		</td>
						</tr>
						<tr>
							<td bgcolor="#FAECC5">�۾����޻���</td>
							<td bgcolor="#FFFFFF"> 
								<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
							</td>
						</tr>
						<tr>
							<td bgcolor="#FAECC5">�����̵� ���<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<input type="radio" id="w_slide1" name="slide_w_flag" value="Y" <% If etemp <> 1 And etemp <> 3 And etemp <> 7 then %>disabled<% End If %> <%=chkiif(slide_w_flag="Y","checked","")%>>��� <input type="radio" id="w_slide2" name="slide_w_flag" value="" <% If etemp <> 1 And etemp <> 3 And etemp <> 7 then %>disabled<% End If %> <%=chkiif(slide_w_flag="","checked","")%>>������
								&nbsp;&nbsp;<input type="button" value="�����̵� ����" onclick="slidechk('w');return false;" id="w_slide3" <% If etemp <> 1 And etemp <> 3 And etemp <> 7 then %>disabled<% End If %>/>
							</td>
						</tr>
						<tr>
							<td width="100" height="50"  bgcolor="#FAECC5">ȭ�����ø�<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF"><%sbGetOptEventCodeValue "eventview",etemp,false,"onchange=""jsChangeFrm(this.value,'P');"""%>
								<span id="divGM" style="display:<%if etemp <> 3 and etemp <> 7 and etemp <> 9 then%>none<%end if%>;">
									<input type="button" value="�׷����" onClick="jsAddGroup('<%=eCode%>','','I','P');" class="button" style="color:blue;width:80" >
									<span  style="float:right;"><input type="checkbox" value="1" name="sgroup_w"  <%=chkiif(sgroup_w=true," checked","")%>> �ֻ��� ��������</span>
									  <%IF not blngroup THEN%>  
									  <div style="padding:5 0 5 0px;display:;" id="divForm" >
									   <input type="button" value="Tab1+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','1','P');" class="button">, 
									   <input type="button" value="Tab2+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','2','P');" class="button">,
									   <input type="button" value="Tab3+����5  �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','3','P');" class="button">  
									  </div> 
									  <%END IF%> 
								</span>
							</td>
						</tr> 
						<tr>
							<td bgcolor="#FAECC5" width="100" align="Center">�̹���<br>&<br>HTML<b style="color:red">*</b></td>
							<td bgcolor="#ffffff">
								<!-- 1.���� ž-->
					   			<div id="divFrm1" style="display:<%if etemp <> 1 then%>none<%end if%>;">
					   				<input type="hidden" name="main" value="<%=emimg%>">
						   			<input type="button" name="btnMain" value="����TOP�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=emimg%>','main','spanmain')" class="button">
						   			<div id="spanmain" style="padding: 5 5 5 5">
						   				<%IF emimg <> "" THEN %>
						   				<a href="javascript:jsImgView('<%=emimg%>')"><img  src="<%=emimg%>" width="400" border="0"></a>
						   				<a href="javascript:jsDelImg('main','spanmain');"><img src="/images/icon_delete2.gif" border="0"></a>
						   				<%END IF%>
						   			</div>
								   	<hr>
									<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('1');">PC-WEB ����</span> 
									<div id="notice1" style="display:block">
									&lt;map name="Mainmap"&gt;<br>
									<font color="blue">��ǰ������ ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('<font color="blue">��ǰ��ȣ</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain('<font color="blue">�̺�Ʈ�ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�̺�Ʈ �׷� �������� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventGroupMain('<font color="blue">�̺�Ʈ�ڵ�</font>','<font color="blue">�׷��ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�̺�Ʈ ����ǰ �˾� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:popShowGiftImg('<font color="blue">�̺�Ʈ�ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">�귣�������� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:GoToBrandShop('<font color="blue">�귣����̵�</font>');" onfocus="this.blur();"&gt;<br>
									&lt;/map&gt;<br>
									<font color="blue">���帮�� ���� ��ũ��</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain_New('<font color="blue">�̺�Ʈ�ڵ�</font>');" onfocus="this.blur();"&gt;<br>
									&lt;/map&gt;
									</map> <br>
									<font color="blue">������ Ÿ��Ʋ �̹����� ��ũ��</font><br>
									&lt;area shape="circle" coords="186,250,144" href="#event_namelink1" onfocus="this.blur();"&gt;<br>
									href="#event_namelink2" href="#event_namelink3" ��� href�� ���ڸ� �ٲ���. &lt;area������ ĭ�� ���������� �� ����.<br>
									</div> 
									<br>
									<b>�̹��� ��� http://<font color="RED">webimage.</font>10x10.co.kr/event/XXX/</b> �� ����Ǿ����ϴ�.<br>
									<textarea name="tHtml" rows="20" style="width:100%;font-size:11px;"><%=ehtml%></textarea>
								</div>
								<!-- 3.�׷���-->
								<div id="divFrm3" style="display:<%if not ( etemp = 3 or etemp = 7 Or blngroup) then%>none<%end if%>;"> 
									<%IF isArray(arrGroup) THEN %>
									<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
									<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
										<td>�׷��ڵ�</td>					
										<td>�����׷�</td>
										<td>�׷��</td>
										<td>���ļ���</td>					
										<td>�̹���</td>
										<td>���ÿ���</td>
										<td>����</td>
									</tr>
									<%FOR intg = 0 To UBound(arrGroup,2)%>				   						
									<tr <%if not arrGroup(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
										<td  ><%IF arrGroup(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup(0,intg)%></td>						
										<td  align="center"><%IF isnull(arrGroup(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrGroup(5,intg)%>]<%=db2html(arrGroup(7,intg))%><%END IF%></td>	
										<td  align="center"><%=db2html(arrGroup(1,intg))%></td>	
										<td  align="center"><%=arrGroup(2,intg)%></td>									   									
										<td  align="center">  
											<a href="javascript:jsImgView('<%=arrGroup(3,intg)%>');"><img src="<%=arrGroup(3,intg)%>" width="50" border="0"></a> 
										</td>	
										<td  align="center"><%if arrGroup(8,intg) then%>Y<%else%>N<%end if%></td>						   									
										<td  align="center">
											<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup(0,intg)%>','P')" class="button">
											<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup(0,intg)%>')"  class="button">-->
											<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrGroup(0,intg)%>','P')"  class="button">
											<% IF arrGroup(5,intg) = 0 THEN %>
											<% 		Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup(0,intg) &"' target='_blank'>�̸�����</a>"
											 %>
											<% END IF %>
										</td>					   									
									</tr>
									<%NEXT%>
									</table>
									<%END IF%>	 
								</div>
								<!-- /3.�׷���-->
								<!-- 5.���۾�-->
								<div id="divFrm5" style="display:<%if  not ( etemp = 5 or etemp = 6) then%>none<%end if%>;">
									<table border="0" cellpadding="1" cellspacing="3" class="a">
										<tr>
											<td>
											    <!-- <input type="button" value="�̹�������"  onclick="TnFtpUpload('D:/home/cube1010/imgstatic/event/','/event/');" class="input_b"> -->
											    <input type="button" value="�̹�������(��)"  onclick="jsManageEventImageNew('<%=eCode%>')" class="input_b">
											    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											    <input type="button" value="�̹�������(��)"  onclick="jsManageEventImage('<%=eCode%>')" class="input_b">
											    <p>
											    <b>�̹�������(��)</b> : ������ ����� �̹��� ����<br>
											    <b>�̹�������(��)</b> : ������ ����� �̹��� ����Ʈ(�̹����߰�����. ���ο� �̹��� �߰��� �̹�������(��)������.)<br>
											    �� �̺�Ʈ �̹��� �ý��� ���� �������� eventIMG ��� ���ο� ������ �̺�Ʈ����Ҵ ������ �߰��Ͽ� �� �ȿ� �̺�Ʈ�ڵ庰 ������ �����ϰ� �˴ϴ�.<br>
											    ���� ����� ���� �ڿ� �̹�������(��)�� ����� �����ʰ� �̹�������(��)�� ����ϰ� �˴ϴ�.<br>
											    �׶������� ��������� �����ô��� �ý��۰��� ������ ���� ��ġ�̹Ƿ� ���عٶ��ϴ�.
											</td>
										</tr>
										<tr>
										    <td><b>�̹��� ��� http://<font color="RED">webimage.</font>10x10.co.kr/eventIMG/�̺�Ʈ����Ҵ/XXX/</b> �� ����Ǿ����ϴ�.</td>
										</tr>
										<tr>
											<td><textarea name="tHtml5" rows="25" style="width:100%;font-size:11px;"><%=ehtml5%></textarea></td>
										</tr> 
									</table>
								</div>
								<!-- /5.���۾�-->
							</td>
						</tr>
						<tr>
                		    <td bgcolor="#FAECC5" width="100" align="Center">Exec File<b style="color:red">*</b>
                		        <br/><span style="color:#602030;font-size:11px;">[ ���� ��������]</span>
                		        </td>
                			<td bgcolor="#ffffff"  >
                		         <input type="radio" name="rdoEF" value="0" <%if not blnExec then%>checked<%end if%>>����� 
						        <input type="radio" name="rdoEF" value="1" <%if blnExec then%>checked<%end if%>>���� <input type="text" name="sEFP" size="60" class="text" value="<%=eExecFile%>"> 
                		    </td>
                		</tr>
					  	<tr id="eNameTr_BL" style="display:<%if ekind<>16 then%>none<%end if%>;"> 
					   		<td  bgcolor="#FAECC5">�귣���̺�Ʈ ��ũ</td>
					   		<td bgcolor="#FFFFFF"> 
					   		 <input type="hidden" name="elType" value="<%=chkiif(eKind=16,"I","E")%>"> 
					   		 <input type="text" id="elUrl" name="elUrl" size="60" maxlength="128" value="<%= elkurl %>" class="text" > 
					   		</td>
					   	</tr> 
					</table>
					<% End If %>
					<% If etype="80" Then %>
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="evt_type2" style="display:<% If etype<>"80" Then %>none<% Else %><% End If %>">
						<tR>
							<td bgcolor="#FAECC5" colspan="2"><b>PC �̺�Ʈ ���� ����</b></td>
						</tr>
						<tr> 
							<td bgcolor="#FAECC5">�׸�����<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
					   			<select name="mdtheme" class="select" onChange="TnSelectTheme(this.value)">
									<option value="1" <% If mdtheme = "1" Then %>selected<% End If %>>�ؽ�Ʈ �׸�</option>
									<option value="2" <% If mdtheme = "2" Then %>selected<% End If %>>�̹��� �׸�</option>
									<option value="3" <% If mdtheme = "3" Then %>selected<% End If %>>��ǰ �׸�</option>
									<!-- <option value="4" <% If mdtheme = "4" Then %>selected<% End If %>>���̵� �׸�</option> -->
								</select>
					   		</td>
						</tr>
						<tr> 
							<td bgcolor="#FAECC5">�̺�Ʈ��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
								<input type="text" name="title_pc" id="title_pc" size="70" maxlength="64" value="<%=title_pc%>">
					   		</td>
						</tr>
						<tr> 
							<td bgcolor="#FAECC5">����ī��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
					   			<textarea name="subcopyK" style="width:100%; height:80px;" value="<%=subcopyK%>"><%=chkiif(subcopyK="","",subcopyK)%></textarea>
								<input type="hidden" name="chkFull" value="1">
								<input type="hidden" name="eventview" value="9">
					   		</td>
						</tr>
						<tr>
							<td bgcolor="#FAECC5">�׸� ����<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<table id="colorselect1" style="display:<% If mdtheme<>"1" Then Response.write "none"%>">
									<tr>
										<td><table id='cline11' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="11" Or themecolor="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#848484' style="font-size:8px"><a href='javascript:selColorChip(11)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="1" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#ed6c6c' style="font-size:8px"><a href='javascript:selColorChip(1)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="2" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f385af' style="font-size:8px"><a href='javascript:selColorChip(2)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline3' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="3" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f3a056' style="font-size:8px"><a href='javascript:selColorChip(3)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline4' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="4" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#e7b93c' style="font-size:8px"><a href='javascript:selColorChip(4)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline5' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="5" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#8eba4a' style="font-size:8px"><a href='javascript:selColorChip(5)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline6' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="6" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#43a251' style="font-size:8px"><a href='javascript:selColorChip(6)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline7' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="7" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#50bdd1' style="font-size:8px"><a href='javascript:selColorChip(7)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline8' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="8" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#5aa5ea' style="font-size:8px"><a href='javascript:selColorChip(8)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline9' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="9" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2672bf' style="font-size:8px"><a href='javascript:selColorChip(9)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='cline10' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="10" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2c5a85' style="font-size:8px"><a href='javascript:selColorChip(10)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										
										<td>*�����Ͻ� �׸� ������ ���� ���� ������ ���õ˴ϴ�.</td>
									</tr>
								</table>
								<table id="colorselect2" style="display:<% If mdtheme="1" Then Response.write "none"%>">
									<tr>
										<td><table id='clinet11' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="11" Or themecolor="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#848484' style="font-size:8px"><a href='javascript:selColorChip2(11)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="1" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#ed6c6c' style="font-size:8px"><a href='javascript:selColorChip2(1)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="2" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f385af' style="font-size:8px"><a href='javascript:selColorChip2(2)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet3' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="3" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f3a056' style="font-size:8px"><a href='javascript:selColorChip2(3)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet4' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="4" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#e7b93c' style="font-size:8px"><a href='javascript:selColorChip2(4)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet5' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="5" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#8eba4a' style="font-size:8px"><a href='javascript:selColorChip2(5)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet6' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="6" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#43a251' style="font-size:8px"><a href='javascript:selColorChip2(6)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet7' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="7" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#50bdd1' style="font-size:8px"><a href='javascript:selColorChip2(7)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet8' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="8" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#5aa5ea' style="font-size:8px"><a href='javascript:selColorChip2(8)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet9' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="9" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2672bf' style="font-size:8px"><a href='javascript:selColorChip2(9)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='clinet10' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolor="10" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2c5a85' style="font-size:8px"><a href='javascript:selColorChip2(10)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										
										<td>*�����Ͻ� �׸� ������ ���� ���� ������ ���õ˴ϴ�.</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr id="textbgcolor" style="display:<% If mdtheme<>"2" Then Response.write "none"%>"> 
							<td bgcolor="#FAECC5">�ؽ�Ʈ ���<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
								<table>
								<tr>
									<td>
										<table>
											<tr>
												<td><table id='bcline1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If textbgcolor="1" Or textbgcolor="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><a href='javascript:selBackGroundColorChip(1)' onfocus='this.blur()'><img src='http://testwebimage.10x10.co.kr/color/colorchip/chip11.gif' alt='���' width='12' height='12' hspace='2' vspace='2' border='0'></a></td></tr></table></td>
												<td><table id='bcline2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If textbgcolor="2" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><a href='javascript:selBackGroundColorChip(2)' onfocus='this.blur()'><img src='http://testwebimage.10x10.co.kr/color/colorchip/chip13.gif' alt='����' width='12' height='12' hspace='2' vspace='2' border='0'></a></td></tr></table></td>
											</tr>
										</table>
									</td>
									<td>*�̹��� �߾��� ī�� ���� ��濡 ������ ó���� ������ �������ּ���.</td>
								</tr>
								</table>
					   		</td>
						</tr>
						<tr id="bnimg" style="display:<% If mdtheme=1 Or mdtheme=0 Then %>none<% End If %>"> 
							<td bgcolor="#FAECC5">��� �̹���<%=mdbntype%><b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
								<table width="100%">
									<tr>
										<td>
											<input type="button" value="��� ���" onclick="TnThemeBannerRegCheck('w');return false;" />
											<input type="hidden" name="bntype" id="bntype" value="<%=mdbntype%>">
										</td>
									</tr>
									<tr>
										<td>
											<% If oitembanner.FresultCount > 0 And mdbntype="T" Then %>
											<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
											<tr>
												<td bgcolor="<%= adminColor("tabletop") %>">����</td>
												<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
												<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
												<td bgcolor="<%= adminColor("tabletop") %>">����</td>
											</tr>
											<input type="hidden" name="pcitemcheck" id="pcitemcheck" value="Y">
											<% For i=0 To oitembanner.FresultCount-1 %>
											<tr>
												<td bgcolor="#FFFFFF"><%= oitembanner.FItemList(i).Fviewidx %></td>
												<td bgcolor="#FFFFFF"><%= oitembanner.FItemList(i).Fitemid %></td>
												<td bgcolor="#FFFFFF"><%= oitembanner.FItemList(i).Fitemname %></td>
												<td bgcolor="#FFFFFF"><a href="javascript:TnThemeItemBannerDel(<%= oitembanner.FItemList(i).Fidx %>);">����</a></td>
											</tr>
											<% Next %>
											</table>
											<% Else %>
											<input type="hidden" name="pcitemcheck" id="pcitemcheck" value="N">
											<% End If %>
											<% If mdbntype="D" Then %>
											<%
											sqlStr = "SELECT idx , slideimg , linkurl , sorting , isusing " + vbcrlf
											sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"' and device = 'W'" + vbcrlf
											sqlStr = sqlStr & " order by sorting asc , idx asc" 
											rsget.Open sqlStr,dbget,1
											%>
											<% if Not(rsget.EOF or rsget.BOF) Then %>
											<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
											<tr>
												<td bgcolor="<%= adminColor("tabletop") %>">����</td>
												<td bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
												<td bgcolor="<%= adminColor("tabletop") %>">����</td>
											</tr>
											<input type="hidden" name="pcbannercheck" id="pcbannercheck" value="Y">
											<% Do Until rsget.eof %>
											<tr>
												<td bgcolor="#FFFFFF"><%=rsget("sorting")%></td>
												<td bgcolor="#FFFFFF"><%=rsget("slideimg")%></td>
												<td bgcolor="#FFFFFF"><a href="javascript:TnDelSlideBanner(<%=rsget("idx")%>);">����</a></td>
											</tr>
											<%
											rsget.movenext
											Loop
											%>
											</table>
											<% Else %>
											<input type="hidden" name="pcbannercheck" id="pcbannercheck" value="N">
											<% End If %>
											<%rsget.close%>
											<% End If %>
										</td>
									</tr>
								</table>
					   		</td>
						</tr>
						<tr>
							<td bgcolor="#FAECC5" width="100">�׷� ����<b style="color:red">*</b></td>
							<td bgcolor="#ffffff">
								<!-- 1.���� ž-->
								<!-- 9.MD�����-->
									<input type="button" value="�׷����" onClick="jsAddGroup2('<%=eCode%>','','I','P','MD');" class="button" style="color:blue;width:80" >
									<span>&nbsp;&nbsp;&nbsp;<input type="checkbox" value="1" name="sgroup_w"  <%=chkiif(sgroup_w=true," checked","")%>> �ֻ��� ��������</span>
								<div id="divFrm3">
									
									  <%IF not blngroup THEN%>  
									  <div style="padding:5 0 5 0px;display:;" id="divForm" >
									   <input type="button" value="Tab1+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','1','P');" class="button">, 
									   <input type="button" value="Tab2+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','2','P');" class="button">,
									   <input type="button" value="Tab3+����5  �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','3','P');" class="button">  
									  </div> 
									  <%END IF%>
									<%IF isArray(arrGroup) THEN %>
									<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
									<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
										<td>�׷��ڵ�</td>					
										<td>�����׷�</td>
										<td>�׷��</td>
										<td>���ļ���</td>					
										<td>���ÿ���</td>
										<td>����</td>
									</tr>
									<%FOR intg = 0 To UBound(arrGroup,2)%>				   						
									<tr <%if not arrGroup(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
										<td  ><%IF arrGroup(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup(0,intg)%></td>						
										<td  align="center"><%IF isnull(arrGroup(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrGroup(5,intg)%>]<%=db2html(arrGroup(7,intg))%><%END IF%></td>	
										<td  align="center"><%=db2html(arrGroup(1,intg))%></td>	
										<td  align="center"><%=arrGroup(2,intg)%></td>									   									
										<td  align="center"><%if arrGroup(8,intg) then%>Y<%else%>N<%end if%></td>						   									
										<td  align="center">
											<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup(0,intg)%>','P')" class="button">
											<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup(0,intg)%>')"  class="button">-->
											<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrGroup(0,intg)%>','P')"  class="button">
											<% IF arrGroup(5,intg) = 0 THEN %>
											<% 		Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup(0,intg) &"' target='_blank'>�̸�����</a>"
											 %>
											<% END IF %>
										</td>					   									
									</tr>
									<%NEXT%>
									</table>
									<%END IF%>	

									
								</div>
								<!-- /9.MD�����-->
							</td>
						</tr>
						<tr>
							<td bgcolor="#FAECC5">�۾����޻���</td>
							<td bgcolor="#FFFFFF"> 
								<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
							</td>
						</tr>
						<tr style="display:none">
                		    <td bgcolor="#FAECC5" width="100">Exec File<b style="color:red">*</b>
                		        <br/><span style="color:#602030;font-size:11px;">[ ���� ��������]</span>
                		        </td>
                			<td bgcolor="#ffffff"  >
                		         <input type="radio" name="rdoEF" value="0" <%if not blnExec then%>checked<%end if%>>����� 
						        <input type="radio" name="rdoEF" value="1" <%if blnExec then%>checked<%end if%>>���� <input type="text" name="sEFP" size="60" class="text" value="<%=eExecFile%>"> 
                		    </td>
                		</tr>
					  	<tr id="eNameTr_BL" style="display:<%if ekind<>16 then%>none<%end if%>;"> 
					   		<td  bgcolor="#FAECC5">�귣���̺�Ʈ ��ũ</td>
					   		<td bgcolor="#FFFFFF"> 
					   		 <input type="hidden" name="elType" value="<%=chkiif(eKind=16,"I","E")%>"> 
					   		 <input type="text" id="elUrl" name="elUrl" size="60" maxlength="128" value="<%= elkurl %>" class="text" > 
					   		</td>
					   	</tr> 
					</table>
					<% End If %>
				    </div>
				</td>
			</tr>
			<tr>
				<td valign="top">
				    <div id="divMA1" style="display:<%if not (isMobile or isApp) then%>none<%end if%>;">
					<% If etype<>"80" Then %>
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="evt_type3" style="display:<% If etype="80" Then %>none<% End If %>">
						<tR>
							<td bgcolor="#e3f1fb"  colspan="2"><b>Mobile �̺�Ʈ ���� ����</b></td>
						</tr>
						<tr>
							<td  bgcolor="#e3f1fb">����Ʈ ����<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>> ������&nbsp;&nbsp;
								<input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>> ����Ʈ��&nbsp;&nbsp;
								<input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>> BIG��
							</td>
						</tr>
						<tr>
							<td bgcolor="#e3f1fb">����ī��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF"> <input type="text" name="subsEN" size="70" maxlength="70" value="<%=enamesub%>"  OnKeyUp="jsAddByte(this);"> 
					   		    <input type="text" name="subSize" size="3" value="" class="text_ro" style="text-align:right" readonly>Byte
					   		     <p style="color:#602030;font-size:11px;"> [ �ִ� 70byte���� ��ϰ����մϴ�. ]</p>
					   		    <script type="text/javascript">
					   		        jsAddByte(frmEvt.subsEN);
					   		     </script>
					   		 </td>
					   	</tr>
						<tr>
							<td bgcolor="#e3f1fb">�۾����޻���</td>
							<td bgcolor="#FFFFFF"> 
								<textarea name="tFwdMo" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd_mo%></textarea>
							</td>
						</tr>
						<tr>
							<td bgcolor="#e3f1fb">�����̵� ���<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<input type="radio" id="m_slide1" name="slide_m_flag" value="Y" <% If etemp <> 1 And etemp <> 3 then %>disabled<% End If %> <%=chkiif(slide_m_flag="Y","checked","")%>>��� <input type="radio" id="m_slide2" name="slide_m_flag" value="" <% If etemp <> 1 And etemp <> 3 then %>disabled<% End If %> <%=chkiif(slide_m_flag="","checked","")%>>������
								&nbsp;&nbsp;<input type="button" value="�����̵� ����" onclick="slidechk('m');return false;" id="m_slide3" <% If etemp <> 1 And etemp <> 3 then %>disabled<% End If %>/>
							</td>
						</tr>
						<tr>
							<td width="100" height="50"  bgcolor="#e3f1fb">ȭ�����ø�<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF"><%sbGetOptEventCodeValue "eventview_mo",etemp_mo,false,"onchange=""jsChangeFrm(this.value,'M');"""%>
								<span id="divGM_mo" style="display:<%if etemp_mo <> 3 and etemp_mo <> 7 and etemp_mo <> 9 then%>none<%end if%>;">
									<input type="button" value="�׷����" onClick="jsAddGroup('<%=eCode%>','','I','M');" class="button" style="color:blue;width:80" >
									<span style="float:right;"><input type="checkbox" value="1" name="sgroup_M" <%=chkiif(sgroup_m=true," checked","")%>> �ֻ��� ��������</span>
									<%IF not blngroup_mo THEN%>
									<div style="padding:3 0 3 0px;display:;" id="divForm_mo">
									    <input type="button" value="Tab1+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','1','M');" class="button">, 
									    <input type="button" value="Tab2+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','2','M');" class="button">,
									    <input type="button" value="Tab3+����5  �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','3','M');" class="button">   
									</div>    
									<%END IF%> 
								</span> 
							</td>
						</tr> 
						<tr>
							<td bgcolor="#e3f1fb" width="100" align="Center">�̹���<br>&<br>HTML<b style="color:red">*</b></td>
							<td bgcolor="#ffffff" valign="top">
									<!-- 1.���� ž-->
								<div id="divMFrm1" style="display:<%if etemp_mo <> 1 then%>none<%end if%>;">
									<input type="hidden" name="main_mo" value="<%=emimg_mo%>">
						   			<input type="button" name="btnMain_mo" value="����TOP�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=emimg_mo%>','main_mo','spanmain_mo')" class="button">
						   			<div id="spanmain_mo" style="padding: 5 5 5 5">
						   				<%IF emimg_mo <> "" THEN %>
						   				<a href="javascript:jsImgView('<%=emimg_mo%>')"><img  src="<%=emimg_mo%>" width="400" border="0"></a>
						   				<a href="javascript:jsDelImg('main_mo','spanmain_mo');"><img src="/images/icon_delete2.gif" border="0"></a>
						   				<%END IF%>
						   			</div>
								  	<hr>
									<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('2');">Mobile-WEB ����</span>||<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('3');">APP ����</span>
									<div id="notice2" style="display:">
										<font color="blue">��ǰ������ ��ũ��</font><br>
										&lt;a href="/category/category_itemprd.asp?itemid=<span style="color:red">��ǰ�ڵ�</span>"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
										&lt;a href="/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>"&gt; �̺�Ʈ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ �׷� �������� ��ũ��</font><br>
										&lt;a href="/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>&eGc=<span style="color:red">�׷��ڵ�</span>"&gt; �̺�Ʈ �׷� ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�귣�������� ��ũ��</font><br>
										&lt;a href="/street/street_brand.asp?makerid=<span style="color:red">�귣���ڵ�</span>"&gt; �귣�������� ��ũ &lt;/a&gt;<br>
									</div>
									<div id="notice3" style="display:none">
										�������������� �̵��Ҷ���<br/>
										<font color="blue">��ǰ������ ��ũ��</font><br>
										&lt;a href="/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<span style="color:red">��ǰ�ڵ�</span>"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
										&lt;a href="/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>"&gt; �̺�Ʈ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ �׷� �������� ��ũ��</font><br>
										&lt;a href="/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<span style="color:red">�̺�Ʈ�ڵ�</span>&eGc=<span style="color:red">�׷��ڵ�</span>"&gt; �̺�Ʈ �׷� ������ ��ũ &lt;/a&gt;<br>
										<br>
										���˾����� ������ ������<br/>
										�ؼ��۾� iframe�߰��Ҷ� �϶� <span style="color:blue">parent.</span> �Լ������� �߰���<br/>
										ex) &lt;a href="#" onclick="<span style="color:blue">parent.</span>fnAPPpopupProduct('<span style="color:red">��ǰ�ڵ�</span>'); return false;"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">��ǰ������ ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupProduct('<span style="color:red">��ǰ�ڵ�</span>'); return false;"&gt; ��ǰ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�̺�Ʈ�������� ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupEvent('<span style="color:red">�̺�Ʈ�ڵ�</span>'); return false;"&gt; �̺�Ʈ������ ��ũ &lt;/a&gt;<br>
										<font color="blue">�귣�������� ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupBrand('<span style="color:red">�귣���</span>'); return false;"&gt; �귣�� ��ũ &lt;/a&gt;<br>
										<font color="blue">ī�װ� ��ũ��</font><br>
										&lt;a href="#" onclick="fnAPPpopupCategory('<span style="color:red">ī�װ���ȣ</span>'); return false;"&gt; ī�װ� ��ũ &lt;/a&gt;<br>
									</div>
									<br>
									<b>�̹��� ��� http://<font color="RED">webimage.</font>10x10.co.kr/event/XXX/</b> �� ����Ǿ����ϴ�.<br>
									<textarea name="tHtml_mo" rows="20" style="width:100%;font-size:11px;"><%=ehtml_mo%></textarea>
								</div>
								<!-- 3.�׷���-->
								<div id="divMFrm3" style="display:<%if not ( etemp_mo = 3 or etemp_mo = 7 Or blngroup_mo) then%>none<%end if%>;">
									<%IF isArray(arrGroup_mo) THEN %>
									<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
									<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
										<td>�׷��ڵ�</td>					
										<td>�����׷�</td>
										<td>�׷��</td>
										<td>���ļ���</td>					
										<td>�̹���</td>
										<td>���ÿ���</td>
										<td>����</td>
									</tr>
									<%
									FOR intg = 0 To UBound(arrGroup_mo,2)
									 sumi= 0
									%>				   						
									<tr <%if not arrGroup_mo(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
										<td  ><%IF arrGroup_mo(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup_mo(0,intg)%> 
										     <% if intg < UBound(arrGroup_mo,2)  then 
                                				 for i = 1 to (UBound(arrGroup_mo,2)-intg)%> 
                                				<%if arrGroup_mo(9,intg) = arrGroup_mo(9,intg+i) then
                                					sumi = sumi + 1 
                                				 %>
                                				 + <%=arrGroup_mo(0,intg+i)%>
                                				<%else 
                                					exit for
                                				 end if 
                                				next
                                			end if    %>
										 </td>						
										<td  align="center"><%IF isnull(arrGroup_mo(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrGroup_mo(5,intg)%>]<%=db2html(arrGroup_mo(7,intg))%><%END IF%></td>	
										<td  align="center"><%=db2html(arrGroup_mo(1,intg))%></td>	
										<td  align="center"><%=arrGroup_mo(2,intg)%></td>									   									
										<td  align="center">  
											<a href="javascript:jsImgView('<%=arrGroup_mo(3,intg)%>');"><img src="<%=arrGroup_mo(3,intg)%>" width="50" border="0"></a> 
										</td>			
										<td  align="center"><%if arrGroup_mo(8,intg) then%>Y<%else%>N<%end if%></td>				   									
										<td  align="center">
											<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup_mo(0,intg)%>','M')" class="button">
											<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup_mo(0,intg)%>')"  class="button">-->
											<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrGroup_mo(0,intg)%>','M')"  class="button">
											<% IF arrGroup_mo(5,intg) = 0 THEN %>
											<% 		Response.Write "<a href=""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup_mo(0,intg) &"','M');"">�̸�����</a>"
											 %>
											<% END IF %>
										</td>					   									
									</tr>
									<%
									     intg = intg+sumi
									NEXT%>
									</table>
									<%END IF%> 
								</div>
								<!-- /3.�׷���-->
								<!-- 5.���۾�-->
								<div id="divMFrm5" style="display:<%if not ( etemp_mo = 5 or etemp_mo = 6) then%>none<%end if%>;">
									<table border="0" cellpadding="1" cellspacing="3" class="a">
										<tr>
											<td>
											    <!-- <input type="button" value="�̹�������"  onclick="TnFtpUpload('D:/home/cube1010/imgstatic/event/','/event/');" class="input_b"> -->
											    <input type="button" value="�̹�������(��)"  onclick="jsManageEventImageNew('<%=eCode%>')" class="input_b">
											    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											    <input type="button" value="�̹�������(��)"  onclick="jsManageEventImage('<%=eCode%>')" class="input_b">
											    <p>
											    <b>�̹�������(��)</b> : ������ ����� �̹��� ����<br>
											    <b>�̹�������(��)</b> : ������ ����� �̹��� ����Ʈ(�̹����߰�����. ���ο� �̹��� �߰��� �̹�������(��)������.)<br>
											    �� �̺�Ʈ �̹��� �ý��� ���� �������� eventIMG ��� ���ο� ������ �̺�Ʈ����Ҵ ������ �߰��Ͽ� �� �ȿ� �̺�Ʈ�ڵ庰 ������ �����ϰ� �˴ϴ�.<br>
											    ���� ����� ���� �ڿ� �̹�������(��)�� ����� �����ʰ� �̹�������(��)�� ����ϰ� �˴ϴ�.<br>
											    �׶������� ��������� �����ô��� �ý��۰��� ������ ���� ��ġ�̹Ƿ� ���عٶ��ϴ�.
											</td>
										</tr>
										<tr>
										    <td><b>�̹��� ��� http://<font color="RED">webimage.</font>10x10.co.kr/eventIMG/�̺�Ʈ����Ҵ/XXX/</b> �� ����Ǿ����ϴ�.</td>
										</tr>
										<tr>
											<td><textarea name="tHtml5_mo" rows="25" style="width:100%;font-size:11px;"><%=ehtml5_mo%></textarea></td>
										</tr> 
									</table>
								</div>
								<!-- /5.���۾�-->
							</td>
						</tr>
						<tr>
						    <td bgcolor="#e3f1fb" width="100" align="Center">Exec File<b style="color:red">*</b><br/> <span style="color:#602030;font-size:11px;">[ ���� ��������]</span></td>
							<td bgcolor="#ffffff"  >
						        <input type="radio" name="rdoEF_mo" value="0" <%if not blnExec_mo then%>checked<%end if%>>����� 
						        <input type="radio" name="rdoEF_mo" value="1" <%if blnExec_mo then%>checked<%end if%>>���� 
						        <input type="text" name="sEFP_mo" size="60" class="text" value="<%=eExecFile_mo%>"> 
						    </td>
						</tr>							 
						<tr>
							<td  bgcolor="#e3f1fb">�ؽ�Ʈ Ÿ��Ʋ</td>  
							<td bgcolor="#FFFFFF">
							    <a href="javascript:jsOpen('<%=vmobileUrl%>/event/eventmain.asp?eventid=<%=eCode%>','M');">  + �̸����� </a>  
								<div id="dvTxT">
								<table border="0" cellpadding="3" cellspacing="1" class="a" width="100%">  
								<tr><%dim chkeCmt, chkeIps, chkeGf, chkeBS%>
									<td colspan="2">
										<% if rdCmt="" THEN rdCmt=1%>
										<div id="dvCmt" style="display:<%IF not ecomment THEN %>none<%end if%>;"> 
										   <table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">
											        <input type="radio" name="rdCmt" value="1" <%if rdCmt="1" THEN%>checked<%END IF%>>�ڸ�Ʈ
											        <input type="radio" name="rdCmt" value="2" <%if rdCmt="2" THEN%>checked<%END IF%>>�׽��� �ڸ�Ʈ
											        <input type="checkbox" name="chkeCmt" value="0" <%if chkeCmt="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eCmt');">������(�̹����� ���)
                                                </th>
        										<tr>
        											<td bgcolor="#e3f1fb">����</td><td bgcolor="#FFFFFF"><textarea type="text" name="eCmtMT" cols="50" rows="3" <%if chkeCmt="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=eCmtMT%></textarea> <span style="color:#602030;font-size:11px;">[200�� �̳�]</span></td>
        										</tr>
        										<tR >
        											<td bgcolor="#e3f1fb">��÷�ڼ�/<br/>����ǰ</td><td bgcolor="#FFFFFF"><textarea cols="70" rows="3" name="eCmtST"  <%if chkeCmt="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=db2Html(eCmtST)%></textarea></td>
        										</tr>
        							        </table> 
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="dvIps" style="display:<%IF not eitemps THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">
												    ��ǰ�ı� 
												    <input type="checkbox" name="chkeIps" value="0" <%if chkeIps="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eIps');">������(�̹����� ���)
												  </th> 
												<tr>
													<td bgcolor="#e3f1fb">����</td><td bgcolor="#FFFFFF"><textarea  name="eIpsMT" cols="50" rows="3"  <%if chkeIps="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=eIpsMT%></textarea> <span style="color:#602030;font-size:11px;">[200�� �̳�]</span></td>
												</tr>
												<tR>
													<td bgcolor="#e3f1fb">��÷�ڼ�/<br/>����ǰ</td><td bgcolor="#FFFFFF"><textarea cols="70" rows="3" name="eIpsST"  <%if chkeIps="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=db2Html(eIpsST)%></textarea></td>
												</tr>
											 </table>
										</div>
									</td>
								</tr>
								 <tr>
									<td colspan="2">
										<div id="dvGf"  style="display:<%IF not egift THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%"  bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">
												    ����ǰ 
												    <input type="checkbox" name="chkeGf" value="0" <%if chkeGf="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eGf');">������(�̹����� ���)
												</th> 
												<tr>
													<td bgcolor="#FFFFFF"><textarea  name="eGfMT" cols="50"  rows="3" <%if chkeGf="0" then%>class="textarea_ro" disabled<%else%> class="textarea"<%end if%>><%=eGfMT%></textarea> <span style="color:#602030;font-size:11px;">[200�� �̳�]</span></td>
												</tr> 
											 </table>
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="dvBS" style="display:<%IF not eBookingsell THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left"  bgcolor="#e3f1fb"> 
												    �����Ǹ� 
												    <input type="checkbox" name="chkeBS" value="0" <%if chkeBS="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eBS');">������(�̹����� ���)
												</th> 
												<tr>
													<td bgcolor="#FFFFFF"><textarea name="eBSMT" cols="50"  rows="3"  <%if chkeBs="0" then%>class="textarea_ro" disabled<%else%> class="textarea"<%end if%>><%=eBSMT%></textarea> <span style="color:#602030;font-size:11px;">[200�� �̳�]</span></td>
												</tr> 
											 </table>
										</div>
									</td>
								</tr> 
								</table>
								</div>
							</td>
						</tr>
						<!-- ��ǰ �̺�Ʈ -->
						<tr id="itemevt" style="display:<%=chkiif(ekind="13","","none")%>;">
							<td bgcolor="#e3f1fb" align="center" colspan="2">
								<div>
								<table border="0" cellpadding="3" cellspacing="1" class="a" width="100%">
								<tr>
									<td bgcolor="#e3f1fb" align="center"><strong>��ǰ�̺�Ʈ</strong></td>
								</tr>
								<tr>
									<td>
										<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
										<tr>
											<td bgcolor="#e3f1fb"><b>�±�</B></td>
											<td bgcolor="#FFFFFF">
												<input type="radio" name="ietag" value="7" <%=chkiif(tmpietag="7","checked","")%> onclick="jstagchk(this.value);"/> ���� <input type="radio" name="ietag" value="2" <%=chkiif(tmpietag="2","checked","")%> onclick="jstagchk(this.value);"/> ���� <input type="text" size="5" name="ietagval" value="<%=tmpietagval%>" <%=chkiif(tmpietag="7" Or tmpietag = "2" ,"","readOnly")%> class="text_ro"/> <input type="radio" name="ietag" value="1" <%=chkiif(tmpietag="1","checked","")%> onclick="jstagchk(this.value);"/> GiFT <input type="radio" name="ietag" value="4" <%=chkiif(tmpietag="4","checked","")%> onclick="jstagchk(this.value);"/> ������ <input type="radio" name="ietag" value="5" <%=chkiif(tmpietag="5","checked","")%> onclick="jstagchk(this.value);"/> 1:1 <input type="radio" name="ietag" value="6" <%=chkiif(tmpietag="6","checked","")%> onclick="jstagchk(this.value);"/> 1+1 <input type="radio" name="ietag" value="3" <%=chkiif(tmpietag="3","checked","")%> onclick="jstagchk(this.value);"/> ������
											</td>
										</tr>
										<tr>
											<td bgcolor="#e3f1fb"><b>���θ�� ����</B></td>
											<td bgcolor="#FFFFFF"><input type="text" size="70" name="mcopy" maxlength="50" value="<%=tmpmcopy%>"/><div style="color:#602030;font-size:11px;padding-top:5px;">( ex: ���� �� �Ϸ�, UDH-02 ���ⷻ�� ����! )</div></td>
										</tr>
										<tr>
											<td bgcolor="#e3f1fb"><b>���� ����</B></td>
											<td bgcolor="#FFFFFF"><input type="text" size="70" name="scopy" maxlength="50" value="<%=tmpscopy%>"/><div style="color:#602030;font-size:11px;padding-top:5px;">( ex: ������ 100��/ ���� �� �������� )</div></td>
										</tr>
										</table>
									</td>
								</tr>
								</table>
								</div>
							</td>
						</tr>
					</table>
					<% End If %>
					<% If etype="80" Then %>
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="evt_type4" style="display:<% If etype<>"80" Then %>none<% Else %><% End If %>">
						<tR>
							<td bgcolor="#e3f1fb"  colspan="2"><b>Mobile �̺�Ʈ ���� ����</b></td>
						</tr>
						<tr>
							<td  bgcolor="#e3f1fb">�׸� ����<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<select name="mdthememo" class="select" onChange="TnSelectThemeMo(this.value)">
									<option value="1" <% If mdthememo = "1" Then %>selected<% End If %>>�ؽ�Ʈ �׸�</option>
									<option value="2" <% If mdthememo = "2" Then %>selected<% End If %>>�̹��� �׸�</option>
									<option value="3" <% If mdthememo = "3" Then %>selected<% End If %>>��ǰ �׸�</option>
								</select>
							</td>
						</tr>
						<tr>
							<td  bgcolor="#e3f1fb">����Ʈ ����<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>> ������&nbsp;&nbsp;
								<input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>> ����Ʈ��&nbsp;&nbsp;
								<input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>> BIG��
							</td>
						</tr>
						<tr>
							<td bgcolor="#e3f1fb">�̺�Ʈ��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF">
								<input type="text" name="title_mo" id="title_mo" size="70" maxlength="64" value="<%=title_mo%>">
					   		 </td>
					   	</tr>
						<tr>
							<td bgcolor="#e3f1fb">����ī��<b style="color:red">*</b></td>
					   		<td bgcolor="#FFFFFF"> <input type="text" name="subsEN" size="70" maxlength="70" value="<%=enamesub%>"  OnKeyUp="jsAddByte(this);"> 
					   		    <input type="text" name="subSize" size="3" value="" class="text_ro" style="text-align:right" readonly>Byte
					   		     <p style="color:#602030;font-size:11px;"> [ �ִ� 70byte���� ��ϰ����մϴ�. ]</p>
					   		    <script type="text/javascript">
					   		        jsAddByte(frmEvt.subsEN);
					   		    </script>
								<input type="hidden" name="eventview_mo" value="9">
					   		 </td>
					   	</tr>
						<tr>
							<td  bgcolor="#e3f1fb">�׸� ����<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<table id="mcolorselect1" style="display:<% If mdthememo<>"1" Then Response.write "none"%>">
									<tr>
										<td><table id='mocline11' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="11" Or themecolormo="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#848484' style="font-size:8px"><a href='javascript:selMoColorChip(11)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="1" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#ed6c6c' style="font-size:8px"><a href='javascript:selMoColorChip(1)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="2" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f385af' style="font-size:8px"><a href='javascript:selMoColorChip(2)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline3' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="3" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f3a056' style="font-size:8px"><a href='javascript:selMoColorChip(3)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline4' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="4" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#e7b93c' style="font-size:8px"><a href='javascript:selMoColorChip(4)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline5' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="5" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#8eba4a' style="font-size:8px"><a href='javascript:selMoColorChip(5)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline6' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="6" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#43a251' style="font-size:8px"><a href='javascript:selMoColorChip(6)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline7' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="7" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#50bdd1' style="font-size:8px"><a href='javascript:selMoColorChip(7)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline8' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="8" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#5aa5ea' style="font-size:8px"><a href='javascript:selMoColorChip(8)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline9' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="9" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2672bf' style="font-size:8px"><a href='javascript:selMoColorChip(9)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='mocline10' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="10" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2c5a85' style="font-size:8px"><a href='javascript:selMoColorChip(10)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										
										<td>*�����Ͻ� �׸� ������ ���� ���� ������ ���õ˴ϴ�.</td>
									</tr>
								</table>
								<table id="mcolorselect2" style="display:<% If mdthememo="1" Then Response.write "none"%>">
									<tr>
										<td><table id='moclinet11' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="11" Or themecolormo="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#848484' style="font-size:8px"><a href='javascript:selMoColorChip2(11)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="1" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#ed6c6c' style="font-size:8px"><a href='javascript:selMoColorChip2(1)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="2" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f385af' style="font-size:8px"><a href='javascript:selMoColorChip2(2)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet3' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="3" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#f3a056' style="font-size:8px"><a href='javascript:selMoColorChip2(3)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet4' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="4" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#e7b93c' style="font-size:8px"><a href='javascript:selMoColorChip2(4)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet5' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="5" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#8eba4a' style="font-size:8px"><a href='javascript:selMoColorChip2(5)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet6' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="6" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#43a251' style="font-size:8px"><a href='javascript:selMoColorChip2(6)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet7' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="7" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#50bdd1' style="font-size:8px"><a href='javascript:selMoColorChip2(7)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet8' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="8" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#5aa5ea' style="font-size:8px"><a href='javascript:selMoColorChip2(8)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet9' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="9" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2672bf' style="font-size:8px"><a href='javascript:selMoColorChip2(9)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td><table id='moclinet10' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If themecolormo="10" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#2c5a85' style="font-size:8px"><a href='javascript:selMoColorChip2(10)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
										<td>*�����Ͻ� �׸� ������ ���� ���� ������ ���õ˴ϴ�.</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr id="textbgcolormo" style="display:<% If mdthememo<>"2" Then Response.write "none"%>">
							<td  bgcolor="#e3f1fb">�ؽ�Ʈ ���<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<table>
								<tr>
									<td>
										<table>
											<tr>
												<td><table id='mobcline1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If textbgcolormo="1" Or textbgcolormo="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><a href='javascript:selMoBackGroundColorChip(1)' onfocus='this.blur()'><img src='http://testwebimage.10x10.co.kr/color/colorchip/chip11.gif' alt='���' width='12' height='12' hspace='2' vspace='2' border='0'></a></td></tr></table></td>
												<td><table id='mobcline2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If textbgcolormo="2" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><a href='javascript:selMoBackGroundColorChip(2)' onfocus='this.blur()'><img src='http://testwebimage.10x10.co.kr/color/colorchip/chip13.gif' alt='����' width='12' height='12' hspace='2' vspace='2' border='0'></a></td></tr></table></td>
											</tr>
										</table>
									</td>
									<td>*�̹��� �߾��� ī�� ���� ��濡 ������ ó���� ������ �������ּ���.</td>
								</tr>
								</table>
							</td>
						</tr>
						<tr id="videomo" style="display:<% If mdthememo<>"2" Then Response.write "none"%>"> 
							<td bgcolor="#e3f1fb">������</td>
					   		<td bgcolor="#FFFFFF">
								<input type="text" name="videoLink" id="videoLink" size="70" maxlength="128" value="<%=videoLink%>">
					   		</td>
						</tr>
						<tr id="bnimgmo" style="display:<% If mdthememo=1 Or mdthememo=0 Then %>none<% End If %>">
							<td  bgcolor="#e3f1fb">��� �̹���<b style="color:red">*</b></td>
							<td bgcolor="#FFFFFF">
								<table>
								<tr>
									<td>
										<input type="button" value="��� ���" onclick="TnThemeBannerRegCheck('m');return false;" />
										<input type="hidden" name="bntypemo" id="bntypemo" value="<%=mdbntypemo%>">
									</td>
								</tr>
								<tr>
									<td>
										<% If moitembanner.FresultCount > 0 And mdbntypemo="T" Then %>
										<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
										<tr>
											<td bgcolor="<%= adminColor("tabletop") %>">����</td>
											<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
											<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
											<td bgcolor="<%= adminColor("tabletop") %>">����</td>
										</tr>
										<input type="hidden" name="moitemcheck" id="moitemcheck" value="Y">
										<% For i=0 To moitembanner.FresultCount-1 %>
										<tr>
											<td bgcolor="#FFFFFF"><%= moitembanner.FItemList(i).Fviewidx %></td>
											<td bgcolor="#FFFFFF"><%= moitembanner.FItemList(i).Fitemid %></td>
											<td bgcolor="#FFFFFF"><%= moitembanner.FItemList(i).Fitemname %></td>
											<td bgcolor="#FFFFFF"><a href="javascript:TnThemeItemBannerDel(<%= moitembanner.FItemList(i).Fidx %>);">����</a></td>
										</tr>
										<% Next %>
										</table>
										<% Else %>
										<input type="hidden" name="moitemcheck" id="moitemcheck" value="N">
										<% End If %>
										<% If mdbntypemo="D" Then %>
										<%
										sqlStr = "SELECT idx , slideimg , linkurl , sorting , isusing " + vbcrlf
										sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"' and device = 'M' " 
										sqlStr = sqlStr & " order by sorting asc , idx asc " 
										rsget.Open sqlStr,dbget,1
										%>
										<% if Not(rsget.EOF or rsget.BOF) Then %>
										<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
										<tr>
											<td bgcolor="<%= adminColor("tabletop") %>">����</td>
											<td bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
											<td bgcolor="<%= adminColor("tabletop") %>">����</td>
										</tr>
										<input type="hidden" name="mobannercheck" id="mobannercheck" value="Y">
										<% Do Until rsget.eof %>
										<tr>
											<td bgcolor="#FFFFFF"><%=rsget("sorting")%></td>
											<td bgcolor="#FFFFFF"><%=rsget("slideimg")%></td>
											<td bgcolor="#FFFFFF"><a href="javascript:TnDelSlideBanner(<%=rsget("idx")%>);">����</a></td>
										</tr>
										<%
										rsget.movenext
										Loop
										%>
										</table>
										<% Else %>
											<input type="hidden" name="mobannercheck" id="mobannercheck" value="N">
										<% End If %>
										<%rsget.close%>
										<% End If %>
									</td>
								</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td bgcolor="#e3f1fb" width="100">�׷� ����<b style="color:red">*</b></td>
							<td bgcolor="#ffffff" valign="top">
								<!-- 9.MD�����-->
								<input type="button" value="�׷����" onClick="jsAddGroup2('<%=eCode%>','','I','M','MD');" class="button" style="color:blue;width:80" >
								<span style="float:right;"><input type="checkbox" value="1" name="sgroup_M" <%=chkiif(sgroup_m=true," checked","")%>> �ֻ��� ��������</span>
								<div id="divMFrm3"> 
									
									<%IF not blngroup_mo THEN%>
									<div style="padding:3 0 3 0px;display:;" id="divForm_mo">
									    <input type="button" value="Tab1+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','1','M');" class="button">, 
									    <input type="button" value="Tab2+����5 �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','2','M');" class="button">,
									    <input type="button" value="Tab3+����5  �׷����" onClick="jsAddProcGroup('<%=eCode%>','F','3','M');" class="button">   
									</div>    
									<%END IF%> 
									<%IF isArray(arrGroup_mo) THEN %>
									<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
									<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
										<td>�׷��ڵ�</td>					
										<td>�����׷�</td>
										<td>�׷��</td>
										<td>���ļ���</td>					
										<td>���ÿ���</td>
										<td>����</td>
									</tr>
									<%
									FOR intg = 0 To UBound(arrGroup_mo,2)
									 sumi= 0
									%>				   						
									<tr <%if not arrGroup_mo(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
										<td  ><%IF arrGroup_mo(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup_mo(0,intg)%> 
										     <% if intg < UBound(arrGroup_mo,2)  then 
                                				 for i = 1 to (UBound(arrGroup_mo,2)-intg)%> 
                                				<%if arrGroup_mo(9,intg) = arrGroup_mo(9,intg+i) then
                                					sumi = sumi + 1 
                                				 %>
                                				 + <%=arrGroup_mo(0,intg+i)%>
                                				<%else 
                                					exit for
                                				 end if 
                                				next
                                			end if    %>
										 </td>						
										<td  align="center"><%IF isnull(arrGroup_mo(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrGroup_mo(5,intg)%>]<%=db2html(arrGroup_mo(7,intg))%><%END IF%></td>	
										<td  align="center"><%=db2html(arrGroup_mo(1,intg))%></td>	
										<td  align="center"><%=arrGroup_mo(2,intg)%></td>									   											
										<td  align="center"><%if arrGroup_mo(8,intg) then%>Y<%else%>N<%end if%></td>				   									
										<td  align="center">
											<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup_mo(0,intg)%>','M')" class="button">
											<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup_mo(0,intg)%>')"  class="button">-->
											<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrGroup_mo(0,intg)%>','M')"  class="button">
											<% IF arrGroup_mo(5,intg) = 0 THEN %>
											<% 		Response.Write "<a href=""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup_mo(0,intg) &"','M');"">�̸�����</a>"
											 %>
											<% END IF %>
										</td>					   									
									</tr>
									<%
									     intg = intg+sumi
									NEXT%>
									</table>
									<%END IF%>
								</div>
								<!-- /9.MD�����-->
							</td>
						</tr>
						<tr>
							<td bgcolor="#e3f1fb">�۾����޻���</td>
							<td bgcolor="#FFFFFF"> 
								<textarea name="tFwdMo" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd_mo%></textarea>
							</td>
						</tr>
						<tr style="display:none">
						    <td bgcolor="#e3f1fb" width="100">Exec File<b style="color:red">*</b><br/> <span style="color:#602030;font-size:11px;">[ ���� ��������]</span></td>
							<td bgcolor="#ffffff"  >
						        <input type="radio" name="rdoEF_mo" value="0" <%if not blnExec_mo then%>checked<%end if%>>����� 
						        <input type="radio" name="rdoEF_mo" value="1" <%if blnExec_mo then%>checked<%end if%>>���� 
						        <input type="text" name="sEFP_mo" size="60" class="text" value="<%=eExecFile_mo%>"> 
						    </td>
						</tr>							 
					</table>
					<% End If %>
				</td>
			</tr> 
		</table>	 
	</td>
</tr>
</table>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" style="display:<% If ecomment Then %><% Else %>none<% End If %>" id="comminfo">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>�ڸ�Ʈ �ؽ�Ʈ �ڽ� ����</B></td>
		   	</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">��뿩��<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
					<input type="radio" name="comm_isusing" value="Y" <%IF comm_isusing="Y" THEN%>checked<%END IF%>>��� <input type="radio" name="comm_isusing" value="N" <%IF comm_isusing="" Or comm_isusing="N" THEN%>checked<%END IF%>>������
				</td>
			</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF">
					<textarea name="comm_text" rows="5" style="width:100%;font-size:11px;"><%=comm_text%></textarea>
				</td>
			</tr>
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ �̹���</td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="btnicon" value="����ǰ�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=freebie_img%>','freebie_img','spanfreebie_img')" class="button">
					<input type="hidden" name="freebie_img" value="<%=freebie_img%>">
					<div id="spanfreebie_img" style="padding: 5 5 5 5">
						<%IF freebie_img <> "" THEN %>
						<a href="javascript:jsImgView('<%=freebie_img%>')"><img  src="<%=freebie_img%>" border="0"></a>
						<a href="javascript:jsDelImg('freebie_img','spanfreebie_img');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ڸ�Ʈ �ۼ��Ⱓ<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
				<%IF estate = 9 THEN%>
		   			������ : <%=comm_start%><input type="hidden" name="comm_start" size="10" value="<% If comm_start <> "" Then %><%=comm_start%><% Else %><%=esday%><% End If %>">
		   			~ ������ : <%=comm_end%> <input type="hidden" name="comm_end" value="<% If comm_end <> "" Then %><%=comm_end%><% Else %><%=eeday%><% End If %>" size="10" >
		   		<%ELSE%>
		   			������ : <input type="text" id="commsdt" name="comm_start" size="10" value="<% If comm_start <> "" Then %><%=comm_start%><% Else %><%=esday%><% End If %>" />
							<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="CommStart_trigger" onclick="return false;" />
		   			~ ������ : <input type="text" id="commedt" name="comm_end" size="10" value="<% If comm_end <> "" Then %><%=comm_end%><% Else %><%=eeday%><% End If %>" />
							<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="CommEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "commsdt", trigger    : "CommStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(comm_end<>"",", max: " & replace(comm_end,"-",""),"")%>
						});
						var CAL_End = new Calendar({
							inputField : "commedt", trigger    : "CommEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(comm_end<>"",", min: " & replace(comm_start,"-",""),"")%>
						});
					</script>
		   		<%END IF%>
				</td>
			</tr>  
		</table> 
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" style="display:<% If eitemps Then %><% Else %>none<% End If %>" id="evalinfo">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>��ǰ�ı� �ؽ�Ʈ �ڽ� ����</B></td>
		   	</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">��뿩��<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
					<input type="radio" name="eval_isusing" value="Y" <%IF eval_isusing="Y" THEN%>checked<%END IF%>>��� <input type="radio" name="eval_isusing" value="N" <%IF eval_isusing="" Or eval_isusing="N" THEN%>checked<%END IF%>>������
				</td>
			</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF">
					<textarea name="eval_text" rows="5" style="width:100%;font-size:11px;"><%=eval_text%></textarea>
				</td>
			</tr>
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ �̹���</td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="btnicon" value="����ǰ�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=eval_freebie_img%>','eval_freebie_img','spanfreebie_img')" class="button">
					<input type="hidden" name="eval_freebie_img" value="<%=eval_freebie_img%>">
					<div id="spaneval_freebie_img" style="padding: 5 5 5 5">
						<%IF eval_freebie_img <> "" THEN %>
						<a href="javascript:jsImgView('<%=eval_freebie_img%>')"><img  src="<%=eval_freebie_img%>" border="0"></a>
						<a href="javascript:jsDelImg('eval_freebie_img','spaneval_freebie_img');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ı� �ۼ��Ⱓ<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
				<%IF estate = 9 THEN%>
		   			������ : <%=eval_start%><input type="hidden" name="eval_start" size="10" value="<% If eval_start <> "" Then %><%=eval_start%><% Else %><%=esday%><% End If %>">
		   			~ ������ : <%=eval_end%> <input type="hidden" name="eval_end" value="<% If eval_end <> "" Then %><%=eval_end%><% Else %><%=eeday%><% End If %>" size="10" >
		   		<%ELSE%>
		   			������ : <input type="text" id="evalsdt" name="eval_start" size="10" value="<% If eval_start <> "" Then %><%=eval_start%><% Else %><%=esday%><% End If %>" />
							<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="EvalStart_trigger" onclick="return false;" />
		   			~ ������ : <input type="text" id="evaledt" name="eval_end" size="10" value="<% If eval_end <> "" Then %><%=eval_end%><% Else %><%=eeday%><% End If %>" />
							<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="EvalEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "evalsdt", trigger    : "EvalStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(null2blank(eval_end)<>"",", max: " & replace(null2blank(eval_end),"-",""),"")%>
						});
						var CAL_End = new Calendar({
							inputField : "evaledt", trigger    : "EvalEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(null2blank(eval_end)<>"",", min: " & replace(null2blank(eval_start),"-",""),"")%>
						});
					</script>
		   		<%END IF%>
				</td>
			</tr>  
		</table> 
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>����Ʈ �ؽ�Ʈ �ڽ� ����</B></td>
		   	</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">��뿩��<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
					<select name="gift_isusing" onchange="TnGiftUsingNum(this.value);">
						<option value="1"<% If gift_isusing=1 Then %> selected<% End If %>>1�� ���</option>
						<option value="2"<% If gift_isusing=2 Then %> selected<% End If %>>2�� ���</option>
						<option value="3"<% If gift_isusing=3 Then %> selected<% End If %>>3�� ���</option>
						<option value="0"<% If gift_isusing=0 Then %> selected<% End If %>>��� ����</option>
					</select>
				</td>
			</tr>
			<tr style="display:" id="gifttxt1">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����Ʈ1 ����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="gift_text1" value="<%=gift_text1%>" size="100" maxlength="64">
				</td>
			</tr>
			<tr style="display:" id="giftimg1">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����Ʈ1 �̹���</td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="btnicon" value="�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=gift_img1%>','gift_img1','spangift_img1')" class="button">
					<input type="hidden" name="gift_img1" value="<%=gift_img1%>">
					<div id="spangift_img1" style="padding: 5 5 5 5">
						<%IF gift_img1 <> "" THEN %>
						<a href="javascript:jsImgView('<%=gift_img1%>')"><img  src="<%=gift_img1%>" border="0"></a>
						<a href="javascript:jsDelImg('gift_img1','spangift_img1');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr>
			<tr style="display:<% If gift_isusing<2 Then %>none<% End If %>" id="gifttxt2">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����Ʈ2 ����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="gift_text2" value="<%=gift_text2%>" size="100" maxlength="64">
				</td>
			</tr>
			<tr style="display:<% If gift_isusing<2 Then %>none<% End If %>" id="giftimg2">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����Ʈ2 �̹���</td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="btnicon" value="�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=gift_img2%>','gift_img2','spangift_img2')" class="button">
					<input type="hidden" name="gift_img2" value="<%=gift_img2%>">
					<div id="spangift_img2" style="padding: 5 5 5 5">
						<%IF gift_img2 <> "" THEN %>
						<a href="javascript:jsImgView('<%=gift_img2%>')"><img  src="<%=gift_img2%>" border="0"></a>
						<a href="javascript:jsDelImg('gift_img2','spangift_img2');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr>
			<tr style="display:<% If gift_isusing<3 Then %>none<% End If %>" id="gifttxt3">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">����Ʈ3 ����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="gift_text3" value="<%=gift_text3%>" size="100" maxlength="64">
				</td>
			</tr>
			<tr style="display:<% If gift_isusing<3 Then %>none<% End If %>" id="giftimg3">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����Ʈ3 �̹���</td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="btnicon" value="�̹��� ���" onClick="jsSetImg('<%=eFolder%>','<%=gift_img3%>','gift_img3','spangift_img3')" class="button">
					<input type="hidden" name="gift_img3" value="<%=gift_img3%>">
					<div id="spangift_img3" style="padding: 5 5 5 5">
						<%IF gift_img3 <> "" THEN %>
						<a href="javascript:jsImgView('<%=gift_img3%>')"><img  src="<%=gift_img3%>" border="0"></a>
						<a href="javascript:jsDelImg('gift_img3','spangift_img3');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
				</td>
			</tr>
		</table> 
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
		   		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>����� ���� �ؽ�Ʈ �ڽ� ����</B></td>
		   	</tr>
			<tr>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">��뿩��<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
					<select name="usinginfo" onchange="TnUsingInfoNum(this.value);">
						<option value="1"<% If usinginfo=1 Then %> selected<% End If %>>1�� ���</option>
						<option value="2"<% If usinginfo=2 Then %> selected<% End If %>>2�� ���</option>
						<option value="3"<% If usinginfo=3 Then %> selected<% End If %>>3�� ���</option>
						<option value="0"<% If usinginfo=0 Then %> selected<% End If %>>��� ����</option>
					</select>
				</td>
			</tr>
			<tr style="display:" id="usingtxt1">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ1 ����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="using_text1" value="<%=using_text1%>" size="40" maxlength="20">
				</td>
			</tr>
			<tr style="display:" id="usingcont1">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ1 ����</td>
				<td bgcolor="#FFFFFF">
					<textarea name="using_contents1" rows="5" style="width:100%;font-size:11px;"><%=using_contents1%></textarea>
				</td>
			</tr>
			<tr style="display:<% If usinginfo<2 Then %>none<% End If %>" id="usingtxt2">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ2 ����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="using_text2" value="<%=using_text2%>" size="40" maxlength="20">
				</td>
			</tr>
			<tr style="display:<% If usinginfo<2 Then %>none<% End If %>" id="usingcont2">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ2 ����</td>
				<td bgcolor="#FFFFFF">
					<textarea name="using_contents2" rows="5" style="width:100%;font-size:11px;"><%=using_contents2%></textarea>
				</td>
			</tr>
			<tr style="display:<% If usinginfo<3 Then %>none<% End If %>" id="usingtxt3">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ3 ����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="using_text3" value="<%=using_text3%>" size="40" maxlength="20">
				</td>
			</tr>
			<tr style="display:<% If usinginfo<3 Then %>none<% End If %>" id="usingcont3">
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ3 ����</td>
				<td bgcolor="#FFFFFF">
					<textarea name="using_contents3" rows="5" style="width:100%;font-size:11px;"><%=using_contents3%></textarea>
				</td>
			</tr>
		</table> 
	</td>
</tr>
<tr>
	<td width="100%" align="right" >
		<input type="image" src="/images/icon_save.gif">
		<a href="index.asp?menupos=<%=menupos%>&<%=strParm%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->