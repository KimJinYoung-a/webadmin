<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
Response.AddHeader "Content-Disposition","attachment;filename=이벤트리스트_" & date & hour(now) & minute(now) & ".xls"
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	'변수선언
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory, sCateMid ,sState,sKind,esale,egift,ecoupon,ebrand,eonlyten,etype_pc, etype_mo,isConfirm,eMng
	Dim strparm
	Dim edgid, edgid2,edgstat1,edgstat2, emdid, epsid, edpid, edgnm, edgnm2, emdnm, epsnm, edpnm, eDiary
	dim eopo,efd,ebs,enew
	dim blnWeb, blnMobile, blnApp
	dim dispCate, maxDepth
	dim blnReqPublish ,sSort
	dim isResearch, mdtheme

	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	maxDepth = 2
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## 검색 #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),60)

	sCategory	= requestCheckVar(Request("selC"),10) 		'카테고리
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'카테고리(중분류)
	dispCate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리
	sState		= requestCheckVar(Request("eventstate"),4)	'이벤트 상태
	 
	sKind 		= requestCheckVar(Request("eventkind"),32)	'이벤트종류
	edgid  		= requestCheckVar(Request("sDgId"),32)		'담당 디자이너
''	edgid2 		= requestCheckVar(Request("sDg2Id"),32)		'서브 디자이너
	emdid  		= requestCheckVar(Request("sMdId"),32)		'담당 MD
	epsid  		= requestCheckVar(Request("sPsId"),32)		'담당 퍼블리셔
	edpid  		= requestCheckVar(Request("sDpId"),32)		'담당 개발자
	
	edgnm  		= requestCheckVar(Request("sdgnm"),32)		'담당 디자이너
''	edgnm2 		= requestCheckVar(Request("sdg2nm"),32)		'서브 디자이너
	emdnm  		= requestCheckVar(Request("smdnm"),32)		'담당 MD
	epsnm  		= requestCheckVar(Request("spsnm"),32)		'담당 퍼블리셔
	edpnm  		= requestCheckVar(Request("sdpnm"),32)		'담당 개발자

	if Request("designerstatus")<>"" AND Request("designerstatus") <> "," then
		edgstat1	= requestCheckVar(Request("designerstatus")(1),2)		'담당 디자이너 상태
		edgstat2	= requestCheckVar(Request("designerstatus")(2),2)		'서브 디자이너 상태
	end if

	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무
	eDiary		= requestCheckVar(Request("chDiary"),2)	'다이어리 유무
	eopo		= requestCheckVar(Request("chopo"),1)	'원플러스원
	efd		= requestCheckVar(Request("chfd"),1)	'무료배송
	ebs		= requestCheckVar(Request("chbs"),1)	'예약판매
	enew		= requestCheckVar(Request("chnew"),1)	'new
	
	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)
	
	dispCate 	= requestCheckvar(request("disp"),16)
	blnReqPublish= requestCheckvar(request("chkPus"),1)
	sSort       = requestCheckvar(request("sSort"),2)

	etype_pc	= requestCheckvar(request("eventtype_pc"),4)
	etype_mo	= requestCheckvar(request("eventtype_mo"),4)
	eMng    = requestCheckvar(request("eventmanager"),4)
	isConfirm	= requestCheckvar(request("blnCnfm"),1)
	mdtheme  	= requestCheckVar(Request("mdtheme"),1)		'MD등록 이벤트 테마

	if isResearch="0" and sKind="" then
		skind="1,12,13,23,27,28,29,31"
	end if

	'이벤트 첫페이지 관심항목이 보이도록 
	IF (sKind="" and isResearch="0") or sKind="1,12" THEN
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") and (not ( session("ssBctId")="fotoark" or session("ssBctId")="arlejk" or session("ssBctId")="barbie8711")) then
			'MD부서라면 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너) - 최이령(fotoark), 이주경(arlejk), 차선화(barbie8711) 제외
			sKind = "1,12,13,16,17,23,24"
		else
			'기타 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너,모바일,브랜드Week)
			sKind = "1,12,13,16,17,23,24,19,25,26,31"
		end if
	end if
	'#######################################
 	if sSort = "" then sSort = "CD"
 	if blnReqPublish= "" then blnReqPublish = False     


dim strSearch, DesignID, strSort, strSubSort, strSql

	if edgnm<>"" then
		DesignID = fnGetWorkerNameToID(edgnm)
	end if

	strSearch = ""

	'//정렬
	IF sSort = "SD" then
	    strSubSort = " A.evt_state Desc , A.evt_code desc "
	    strSort = " evt_state Desc , evt_code desc "
	ELSEIF sSort = "SA" then
	    strSubSort = " A.evt_state Asc ,  A.evt_code desc  "
	    strSort = " evt_state Asc , evt_code desc  "
	ELSEIF sSort = "DD" then
	    strSubSort = " A.evt_startdate Desc ,  A.evt_code desc  "
	     strSort = " evt_startdate Desc , evt_code desc  "
	ELSEIF sSort = "DA" then
	    strSubSort = "  A.evt_startdate Asc ,  A.evt_code desc  "
	    strSort = " evt_startdate Asc , evt_code desc  "
	ELSEIF sSort = "ID" then
	    strSubSort = " A.evt_imgregdate Desc ,  A.evt_code desc  "
	     strSort = " evt_imgregdate Desc , evt_code desc  "
	ELSEIF sSort = "IA" then
	    strSubSort = "  A.evt_imgregdate Asc ,  A.evt_code desc  "
	    strSort = " evt_imgregdate Asc , evt_code desc  "
	ELSEIF sSort = "CA" then
	    strSubSort = "  A.evt_code Asc "
	    strSort = " evt_code Asc "
	ELSE
	    strSubSort = " A.evt_code DESC "
	    strSort = " evt_code DESC "
    END IF

	'//검색조건
	If sSdate <> ""  or sEdate <> "" THEN
		if CStr(sDate) = "S" THEN
			strSearch  = strSearch & " and  datediff(day, '"&sSdate&"', evt_startdate) >= 0 and  datediff(day,'"&sEdate&"', evt_startdate) <=0  "
		elseif CStr(sDate) = "E" THEN
			strSearch  = strSearch & " and  datediff(day,'"&sSdate&"',evt_enddate) >= 0 and  datediff(day,'"&sEdate&"',evt_enddate) <=0  "
		elseif CStr(sDate) = "O" THEN
			strSearch  = strSearch & " and  datediff(day,'"&sSdate&"',opendate) >= 0 and  datediff(day,'"&sEdate&"',opendate) <=0  "
		end if
	END IF
	If strTxt <> "" THEN
		IF Cstr(sEvt) = "evt_code" THEN
			'이벤트 코드 검색
			If chkWord(strTxt,"[^-0-9 ]") = "False" Then
				Alert_return("이벤트코드는 숫자만 입력하세요")
				response.end
			End If
			strSearch  = strSearch &  " and A.evt_code = "&strTxt
		ElseIF Cstr(sEvt) = "evt_tag" THEN
			'이벤트 태그 검색
			strSearch  = strSearch &  " and B.evt_tag like '%"&strTxt&"%'"
		ElseIF Cstr(sEvt) = "evt_sub" THEN
			'이벤트 서브카피 검색
			strSearch  = strSearch &  " and  (A.evt_subcopyK like '%"&strTxt&"%' or A.evt_subname like '%"&strTxt&"%') "
		ELSE
			'이벤트명 + 작업태그 검색
			strSearch  = strSearch &  " and  (A.evt_name like '%"&strTxt&"%' or B.workTag like '%"&strTxt&"%') "
		END IF
	End If

	If sState <> "" THEN
		IF sState = "9" THEN	'종료
			strSearch  = strSearch & " and   (evt_state = 9 or  datediff(day,getdate(),evt_enddate)< 0 )"
		ELSEIF sState = "7" THEN	'오픈예정
		    strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)> 0 and  datediff(day,getdate(),evt_enddate)>=0 "
		ELSEIF sState = "6" THEN	'오픈진행중
			strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0 and datediff(day,getdate(),evt_enddate) >= 0  "
		ELSEIF sState = "1^3" THEN
		    strSearch  = strSearch & " and  ( evt_state = 1 or  evt_state = 3 ) and     datediff(day,getdate(),evt_enddate)>=0"
		ELSEIF sState = "6^9" THEN
		    strSearch  = strSearch & " and  (( evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0   ) or  evt_state = 9  or   datediff(day,getdate(),evt_enddate)< 0)        "
		ELSE
			strSearch  = strSearch & " and  evt_state = "&sState & " and  datediff(day,getdate(),evt_enddate)>=0"
		END IF
	End If

	If etype_pc <> "" THEN strSearch  = strSearch &  " and  eventtype_pc='" & etype_pc & "'"
	If etype_mo <> "" THEN strSearch  = strSearch &  " and  eventtype_mo='" & etype_mo & "'"

	If sCategory <> "" THEN strSearch  = strSearch &  " and  evt_category = "&sCategory
	If sCateMid <> "" THEN strSearch  = strSearch &  " and  evt_cateMid = "&sCateMid
	If dispCate<>"" then	strSearch  = strSearch &  " and  evt_dispcate like '"& dispCate & "%'"

	IF sKind <> "" THEN
		strSearch  = strSearch &  " and evt_kind in ("& sKind & ") "
	END IF

	IF edgid <> "" THEN
		strSearch  = strSearch &  " and (B.designerid = '"&edgid&"' or B.designerid2 = '"&edgid&"')"
	END IF

	IF emdid <> "" THEN
		strSearch  = strSearch &  " and B.partMDid = '"&emdid&"'"
	END IF

	IF epsid <> "" THEN
		strSearch  = strSearch &  " and B.publisherid = '"&epsid&"'"
	END IF

	IF edpid <> "" THEN
		strSearch  = strSearch &  " and B.developerid = '"&edpid&"'"
	END IF


	IF DesignID <> "" THEN
		strSearch  = strSearch &  " and ("
		strSearch  = strSearch &  " B.designerid='"&DesignID&"' or "
		strSearch  = strSearch &  " B.designerid2='"&DesignID&"')"
	END IF

	IF emdnm <> "" THEN
		strSearch  = strSearch &  " and D.username = '"&emdnm&"'"
	END IF

	IF epsnm <> "" THEN
		strSearch  = strSearch &  " and E.username = '"&epsnm&"'"
	END IF

	IF edpnm <> "" THEN
		strSearch  = strSearch &  " and F.username = '"&edpnm&"'"
	END IF

	IF ebrand <> "" THEN
		strSearch  = strSearch & " and brand = '"&ebrand&"'"
	END If

	IF mdtheme <> "" THEN
		strSearch  = strSearch & " and (B.mdtheme='"&mdtheme&"' or B.mdthememo='"&mdtheme&"')"
	END IF

	if edgstat1<>"" then strSearch  = strSearch & " and dsn_state1=" & edgstat1
	if edgstat2<>"" then strSearch  = strSearch & " and dsn_state2=" & edgstat2

	if isConfirm="1" then strSearch  = strSearch & " and evt_type=50 and isConfirm=1 "
	if eMng <> "" then strSearch = strSearch & " and evt_manager ="&eMng
	IF esale = "1" THEN strSearch  = strSearch & " and issale = 1 "
	IF egift = "1" THEN strSearch  = strSearch & " and isgift = 1 "
	IF ecoupon = "1" THEN strSearch  = strSearch & " and iscoupon = 1 "
	IF eonlyten = "1" THEN strSearch  = strSearch & " and isOnlyTen = 1 "
	IF eDiary = "1" THEN strSearch  = strSearch & " and isDiary = 1 "
	IF eopo   = "1" THEN strSearch  = strSearch & " and isoneplusone = 1 "
	IF efd = "1" THEN strSearch  = strSearch & " and isfreedelivery = 1 "
	IF ebs = "1" THEN strSearch  = strSearch & " and isbookingsell = 1 "
	IF enew = "1" THEN strSearch  = strSearch & " and isNew = 1 "

	if Not(blnWeb="" and blnMobile="" and blnApp="") then
		IF blnWeb = "1" then
			strSearch = strSearch & " and isWeb = 1 "
		else
			strSearch = strSearch & " and isWeb = 0 "
		end IF
		IF blnMobile = "1" then
			strSearch = strSearch & " and isMobile=1 "
		else
			strSearch = strSearch & " and isMobile=0 "
		end if
		IF blnApp = "1" then
			strSearch = strSearch & " and isApp=1 "
		else
			strSearch = strSearch & " and isApp=0 "
		end if
	end if

	IF blnReqPublish ="1" then strSearch = strSearch & " and isReqPublish = 1 "

	strSql = "SELECT A.evt_code, A.evt_kind, A.evt_manager, A.evt_scope, A.evt_name, A.evt_startdate, A.evt_enddate, A.evt_level  "&_
			"		,evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9  "&_
			"		When A.evt_state = 7 and DateDiff(day,getdate(),evt_startdate)  <= 0 Then 6 "&_
			"		ELSE A.evt_state  "&_
			"		end  "&_
			"		,A.evt_regdate,B.evt_bannerimg, isNull(C.username,'') as designername "&_
			"		,(SELECT code_nm from [db_item].[dbo].tbl_Cate_large WHERE code_large = B.evt_category) categoryname "&_
			"		, A.evt_prizedate , B.brand, B.issale, B.isgift, B.iscoupon "&_
			"		, (SELECT COUNT(sale_code) FROM [db_event].[dbo].[tbl_sale] WHERE evt_code = A.evt_code and sale_using =1) as sale_count  "&_
			"		, (SELECT COUNT(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = A.evt_code and gift_using ='y') as gift_count "&_
			"		, A.prizeyn  "&_
			"		, (Case When A.evt_kind=13 Then (Select top 1 itemid from [db_event].[dbo].[tbl_eventitem] where evt_code=A.evt_code) else 0 end) as itemid  "&_
			"		,(select top 1 code_nm from db_item.dbo.tbl_Cate_mid where code_large=b.evt_category and code_mid=b.evt_cateMid) as code_nm  "&_
			"		, D.username as mdname "&_
			"		, isNull(B.evt_bannerimg2010,'') AS evt_bannerimg2010, B.workTag  "&_
			"		,(select top 1 catename from db_item.dbo.tbl_display_cate where catecode=left(b.evt_dispcate,3)) as dispcate_nm ,B.evt_itemsort "&_
			"		, E.username as psname, F.username as dpname , A.isWeb, A.isMobile, A.isApp, B.isDiary ,etc_itemimg ,evt_mo_listbanner, evt_imgregdate, B.evt_mo_listbannerTXT "&_
			"		, G.username as ccname, A.evt_type "&_
			"		, isNull(B.designerid2,'') as designerid2, isNull(C2.username,'') as designername2, isNull(B.dsn_state1,'') as dsn_state1, isNull(B.dsn_state2,'') as dsn_state2, B.eventtype_pc, B.eventtype_mo "&_
			"		, B.iscomment, B.isbbs, B.isitemps, B.isGetBlogURL "&_
			"	FROM [db_event].[dbo].[tbl_event] as A  "&_
			"		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B ON A.evt_code = B.evt_code "&_
			"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as C ON C.userid = B.designerid   and b.designerid is not null and b.designerid <> '' "&_
			"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as C2 ON C2.userid = B.designerid2   and b.designerid2 is not null and b.designerid2 <> '' "&_
			"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as D ON D.userid = B.partMDid  and b.partMDid is not null and b.partMDid <> '' "&_
			"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as E ON E.userid = B.publisherid and b.publisherid is not null and b.publisherid <> '' "&_
			"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as F ON F.userid = B.developerid and b.developerid is not null and b.developerid <> '' "&_
			"		LEFT OUTER JOIN db_partner.dbo.tbl_user_tenbyten as G ON G.userid = B.codecheckerid and b.codecheckerid is not null and b.codecheckerid <> '' "&_
			"	WHERE evt_using ='Y'  " &strSearch &_
			" order by "&strSort

''		response.Write strSql
	rsget.Open strSql,dbget,0
		IF not rsget.EOF THEN
			arrList = rsget.getRows()
		End IF
	rsget.Close

	Dim arreventlevel, arreventstate, arreventkind, arreventtype, arrdsnStat,arreventmanager
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arreventlevel = fnSetCommonCodeArr("eventlevel",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)
	arreventkind= fnSetCommonCodeArr("eventkind",False)
	arreventtype= fnSetCommonCodeArr("eventtype",False)
	arrdsnStat = fnSetCommonCodeArr("designerstatus",False)
	arreventmanager = fnSetCommonCodeArr("eventmanager",False)
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
br { mso-data-placement:same-cell; }
</style>
</head>
<body>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="100">채널</td>
			<td width="100">주체</td>
			<td width="150">이벤트유형(PC)</td>
			<td width="150">이벤트유형(모바일)</td>
			<td width="100"><b>이벤트코드</b></td>
			<td width="100"><b>진행상태</b></td>
			<td>이벤트명</td>
			<td width="200">카테고리</td>
			<td width="150"><b>시작일</b></td>
			<td width="150">종료일</td>
			<td width="100">디자이너</td>
    </tr>
    <%IF isArray(arrList) THEN
		Dim itemSortvalue
		Dim strURL
		Dim isMobile, isApp, isWeb
		 dim tmpename, ename,eSalePer
		
    	For intLoop = 0 To UBound(arrList,2) 
		
		'2014-08-27 김진영 / 변수에 순서값 저장
		Select Case arrList(27,intLoop)
			Case "1"	itemSortvalue = "sitemid"
			Case "2"	itemSortvalue = "slsell"
			Case "3"	itemSortvalue = "sevtitem"
			Case "4"	itemSortvalue = "sbest"
			Case "5"	itemSortvalue = "shsell"
		End Select
		
		isWeb = False
		isMobile = False
		isApp = False
		
		IF isNull(arrList(30,intLoop)) and isNull(arrList(31,intLoop)) and isNull(arrList(32,intLoop)) then
			if arrList(1,intLoop) = "19" THEN
				isWeb = False
				isMobile = True
				isApp = True
			ELSEIF arrList(1,intLoop) = "25"  THEN
				isWeb = False
				isMobile = False
				isApp = True
			ELSEIF arrList(1,intLoop) = "26"  THEN	
				isWeb = False
				isMobile = True
				isApp = False
			ELSE
				isWeb = True
				isMobile = False
				isApp = False	
			END IF
		END IF	
		IF 	 not isNull(arrList(30,intLoop))  THEN	
			isWeb = arrList(30,intLoop)
		END IF	
		IF 	 not isNull(arrList(31,intLoop)) THEN
			 isMobile = arrList(31,intLoop)
		END IF	 
		IF 	 not isNull(arrList(32,intLoop)) THEN
			isApp = arrList(32,intLoop)	
		END IF	
		
		 
    %>
    <tr align="center" bgcolor="#FFFFFF">
			<td>
				<%IF isWeb THEN %>Web<%END IF%>
				<%=chkIIF(isMobile,"<br /><font color=""blue"">Mobile</font>","")%>
				<%=chkIIF(isApp,"<br /><font color=""red"">App</font>","")%>
			</td>
			<td><%=fnGetCommCodeArrDesc(arreventmanager,arrList(2,intLoop))%></td>
			<td>
				<% If arrList(45,intLoop)>0 Then %>
				<%=fnGetCommCodeArrDesc(arreventtype,arrList(44,intLoop))%>
				<% Else %>
				<%=fnGetCommCodeArrDesc(arreventtype,arrList(39,intLoop))%>
				<% End If %>
			</td>
			<td>
				<% If arrList(45,intLoop)>0 Then %>
				<%=fnGetCommCodeArrDesc(arreventtype,arrList(45,intLoop))%>
				<% Else %>
				<%=fnGetCommCodeArrDesc(arreventtype,arrList(39,intLoop))%>
				<% End If %>
			</td>
			<td><%=arrList(0,intLoop)%></td>
			<td><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></td>
			<td align="left">
				<%=chkIIF(Not(arrList(25,intLoop)="" or isNull(arrList(25,intLoop))),"["&arrList(25,intLoop)&"] ","")%>
				<%   ename =  arrList(4,intLoop)  
				eSalePer = ""
				if  (arrList(15,intLoop) or arrList(17,intLoop)) then 
				tmpename = Split(ename,"|")  
				if Ubound(tmpename)>0 then
				ename = tmpename(0)
				eSalePer = tmpename(1)
				end if

				end if
				%>   
				<%=db2html(ename)%>
			</td>
			<td>
				<%=arrList(12,intLoop)%>
				<%
				if arrList(22,intLoop) <> "" then
				response.write "(" & arrList(22,intLoop) &")"
				end if
				'전시카테고리
				if arrList(26,intLoop)<>"" then
				response.write chkIIF(arrList(12,intLoop)<>"","<br/>","") & "<font color='#4030A0'>" & arrList(26,intLoop) & "</font>"
				end if
				%>
			</td>

			<td><%=arrList(5,intLoop)%></td>
			<td><%=arrList(6,intLoop)%></td> 
			<td><%=arrList(11,intLoop)%></td>
    </tr>
<% Next %>
<% end if %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->