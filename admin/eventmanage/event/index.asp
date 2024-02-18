<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/index.asp
' Description :  이벤트 등록 - 화면설정
' History : 2007.02.07 정윤정 생성
'           2012.02.13 허진원 - 미니달력 교체
'						2014.03.10 정윤정 - 관심항목 최이령(fotoark), 이주경(arlejk) 예외사항 설정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

	'변수선언
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory, sCateMid ,sState,sKind,esale,egift,ecoupon,ebrand,eonlyten, dispCate
	Dim strparm
	Dim edid, emid, eDiary

	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

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
	edid  		= requestCheckVar(Request("selDId"),32)		'담당 디자이너
	emid  		= requestCheckVar(Request("selMId"),32)		'담당 MD

	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무
	eDiary	= requestCheckVar(Request("chDiary"),2)	'다이어리 유무

	'이벤트 첫페이지 관심항목이 보이도록
	IF sKind="" or sKind="1,12" THEN
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") and (not ( session("ssBctId")="fotoark" or session("ssBctId")="arlejk" or session("ssBctId")="barbie8711")) then
			'MD부서라면 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너) - 최이령(fotoark), 이주경(arlejk), 차선화(barbie8711) 제외
			sKind = "1,12,13,16,17,23,24"
		else
			'기타 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너,모바일)
			sKind = "1,12,13,16,17,23,24,19,25,26"
		end if
	end if
	strparm  = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&selDId="&edid&"&selMId="&emid&_
				"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOnlyTen="&eonlyten&"&disp="&dispCate&"&chDiary="&eDiary
	'#######################################

	'데이터 가져오기
	set cEvtList = new ClsEvent
		cEvtList.FCPage = iCurrpage		'현재페이지
		cEvtList.FPSize = iPageSize		'한페이지에 보이는 레코드갯수

		cEvtList.FSfDate 	= sDate		'기간 검색 기준
		cEvtList.FSsDate 	= sSdate	'검색 시작일
		cEvtList.FSeDate 	= sEdate	'검색 종료일
		cEvtList.FSfEvt 	= sEvt		'검색 이벤트명 or 이벤트코드
		cEvtList.FSeTxt 	= strTxt	'검색어
		cEvtList.FScategory = sCategory	'검색 카테고리
		cEvtList.FScateMid	= sCateMid	'검색 카테고리(중분류)
		cEvtList.FEDispCate	= dispCate	'검색 전시카테고리
		cEvtList.FSstate 	= sState	'검색 상태
		cEvtList.FSedid   	= edid
		cEvtList.FSemid   	= emid
		cEvtList.FSkind 	= sKind
		cEvtList.FEBrand 	= ebrand
		cEvtList.FSisSale 	= esale
		cEvtList.FSisGift 	= egift
		cEvtList.FSisCoupon	= ecoupon
		cEvtList.FSisOnlyTen= eonlyten
		cEvtList.FSisDiary = eDiary

 		arrList = cEvtList.fnGetEventList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

	Dim arreventlevel, arreventstate, arreventkind
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arreventlevel = fnSetCommonCodeArr("eventlevel",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)
	arreventkind= fnSetCommonCodeArr("eventkind",False)
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.submit();
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}

		if(frm.selEvt.value== "evt_code"&&frm.sEtxt.value!=""){
			if(!IsDigit(frm.sEtxt.value)){
				alert("이벤트코드는 숫자만 가능합니다.");
				frm.sEtxt.focus();
				return;
			}
		}

		frm.submit();
	}

	function SubmitForm() {
		jsSearch('E');
	}

	function jsSchedule(){
		var winS;
		winS = window.open('pop_event_schedule.asp','popwin','width=1200, height=800, scrollbars=yes');
		winS.focus();
	}




	function jsCodeManage(){
		var winCode;
		winCode = window.open('/admin/eventmanage/code/popManageCode.asp','popCode','width=400,height=600');
		winCode.focus();
	}

	function prize(evt_code){

		 var prize = window.open('/admin/eventmanage/event/pop_event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
		 prize.focus();

	}
	function workerlist()
	{
		var openWorker = null;
		var worker = frmEvt.selMId.value;
		openWorker = window.open('popWorkerList.asp?worker='+worker+'&department_id=7','openWorker','width=700,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function workerDel()
	{
		var frm = document.frmEvt;

		frm.selMId.value = "";
		frm.doc_workername.value = "";
	}
//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch('E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>브랜드:</td>
				<td><% drawSelectBoxDesignerwithName "ebrand", ebrand %></td>
				<td colspan="5">
					관리 <!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					/ 전시 카테고리 : <%=fnDispCateSelectBox(1,"","disp",dispCate,"") %>
				</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">이벤트종류:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><%sbGetOptCommonCodeArr "eventkind", sKind, False,True,"onChange='javascript:document.frmEvt.submit();'"%></td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">코드/명:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selEvt">
			    	<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
			    	<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
			    	</select>
			        <input type="text" name="sEtxt" value="<%=strTxt%>" maxlength="60"></td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" colspan="2">이벤트타입:
			    	<input type="checkbox" name="chSale" <%IF Cstr(esale)="1" THEN%> checked <%END IF%>  value="1">할인
					<input type="checkbox" name="chGift" <%IF Cstr(egift)="1" THEN%> checked<%END IF%>  value="1">사은품
					<input type="checkbox" name="chCoupon" <%IF Cstr(ecoupon)="1" THEN%> checked<%END IF%> value="1">쿠폰
					<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="checkbox" name="chOnlyTen" <%IF Cstr(eonlyten)="1" THEN%> checked<%END IF%> value="1">Only-TenByTen
					<input type="checkbox" name="chDiary" <%IF Cstr(eDiary)="1" THEN%> checked<%END IF%> value="1">DiaryStory
			    </td>
				<td  style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">진행상태:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><%sbGetOptCommonCodeArr "eventstate", sState, True,False,"onChange='javascript:SubmitForm();'"%>	</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">기간:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selDate">
			    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
			    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
			        </select>
			        <input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			        <input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "iSD", trigger    : "iSD_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "iED", trigger    : "iED_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
			    </td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">담당웹디: <%sbGetDesignerid "selDId",edid, "onChange='javascript:document.frmEvt.submit();'"%></td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">담당자:
					<% sbGetwork "selMId",emid,"" %>
			    	<% 'sbGetMDid "selMId",emid, "onChange='javascript:document.frmEvt.submit();'" %>
			    </td>
			    <td  style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
			</tr>
			</table>
        </td>
    		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch('E');">
		</td>
	</tr>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="40" valign="bottom">
        <td align="left">
        	<input type="button" value="새로등록" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>&<%=strParm%>');" class="button">
	    </td>
	    <td align="right">
	       	<input type="button" value="스케쥴" onclick="jsSchedule();"  class="button">
	       <!--	<input type="button" value="통계" onclick=" ">  -->
	       <% if C_ADMIN_AUTH then %><input type="button" value="코드관리" onclick="jsCodeManage();"  class="button"><%END IF%>
        </td>
	</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="13">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td nowrap>이벤트종류</td>
    	<td nowrap>이벤트코드</td>
    	<td nowrap>중요도</td>
      	<td nowrap>진행상태</td>
      	<td nowrap>배너이미지</td>
      	<td nowrap>이벤트명</td>
      	<td nowrap>카테고리</td>
      	<td nowrap>브랜드</td>
      	<td width="60">시작일</td>
      	<td width="60">종료일</td>
      	<td nowrap>담당자</td>
      	<td nowrap>담당웹디</td>
      	<td nowrap>관리</td>
    </tr>

    <%IF isArray(arrList) THEN
		Dim itemSortvalue
    	For intLoop = 0 To UBound(arrList,2)
		'2014-08-27 김진영 / 변수에 순서값 저장
		Select Case arrList(27,intLoop)
			Case "1"	itemSortvalue = "sitemid"
			Case "2"	itemSortvalue = "slsell"
			Case "3"	itemSortvalue = "sevtitem"
			Case "4"	itemSortvalue = "sbest"
			Case "5"	itemSortvalue = "shsell"
		End Select
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></a></a></td>
		<%
			'이벤트 종류에 따른 프론트링크 페이지 선택
			Select Case arrList(1,intLoop)
				Case "7"		'위클리코디
					Response.Write "<td><a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "13"		'상품 이벤트
					Response.Write "<td><a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & arrList(21,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "14"		'소풍가는길
					Response.Write "<td><a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
'				Case "15"		'브랜드데이
'					Response.Write "<td><a href='" & vwwwUrl & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "16"		'브랜드 할인행사
					Response.Write "<td><a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & arrList(14,intLoop) & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "22"		'DAY&(데이앤드)
					Response.Write "<td><a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "26"		'모바일
					Response.Write "<td><a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case Else		'쇼핑찬스 및 기타
					Response.Write "<td><a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
			End Select
		%>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventlevel,arrList(7,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(10,intLoop) <> "" THEN%> <img src="<%=arrList(10,intLoop)%>" width="100" border="0"><%END IF%></a></td>
      	<!--<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(24,intLoop) <> "" THEN%> <img src="<%=arrList(24,intLoop)%>" width="100" border="0"><%END IF%></a><% If arrList(10,intLoop) = "" OR isNull(arrList(10,intLoop)) Then Response.Write "N" Else Response.Write "Y" End If%></td>//-->
      	<td align="left">
      		<a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=chkIIF(Not(arrList(25,intLoop)="" or isNull(arrList(25,intLoop))),"["&arrList(25,intLoop)&"] ","")%>
      		<%=db2html(arrList(4,intLoop))%>
      		<%if arrList(15,intLoop)  then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_sale.gif" border="0"><%end if%>
      		<%if arrList(16,intLoop) then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_gift.gif" border="0"><%end if%>
      		<%if arrList(17,intLoop) then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif" border="0"><%end if%>
      		</a>
      	</td>
      	<td>
      		<a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>">
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
      		</a>
      	</td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(14,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(23,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(11,intLoop)%></a></td>
   		<td align="left" nowrap><input type="button" value="상품" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>&selsort=<%=itemSortvalue%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="당첨" class="button" onClick="jsGoUrl('eventprize_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')"><%END IF%>
      		<%if arrList(15,intLoop)  then%> <input type="button" value="할인(<%=arrList(18,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/sale/salelist.asp?eC=<%=arrList(0,intLoop)%>&menupos=290');"><%end if%>
      		<%if arrList(16,intLoop) then%> <input type="button" value="사은품(<%=arrList(19,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/gift/giftlist.asp?eC=<%=arrList(0,intLoop)%>&menupos=1045');"><%end if%>
      		<!--<%if arrList(17,intLoop) then%> <input type="button" value="쿠폰" class="button" onClick="jsGoUrl('coupon');"><%end if%>	-->
      		<% If arrList(20,intLoop) = "N" Then %>
      		<table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding:3 0 0 0;"><input type="button" class="button" style="width:105;" value="당첨자없음 설정" onclick="prize(<%= arrList(0,intLoop) %>);"></td></tr></table>
      		<% End IF %>
      	</td>
    </tr>
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="11">등록된 내용이 없습니다.</td>
   	</tr>
   <%END IF%>
</table>
<!-- 페이징처리 -->
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td valign="bottom" align="center">
         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(iCurrpage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% else %>[next]<% end if %>
        </td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
