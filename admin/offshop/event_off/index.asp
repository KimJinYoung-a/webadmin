<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 이벤트
' History : 2010.03.09 한용민 생성
'           2012.02.14 허진원 - 미니달력 교체
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
Call fnSetEventCommonCode_off '공통코드 어플리케이션 변수에 세팅

Dim iPageSize, iCurrpage ,iDelCnt ,cEvtAddedShop, j
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt ,isgift ,israck ,isprize
Dim selDate,evt_startdate,evt_enddate, evt_code,strTxt, sCategory, sCateMid ,evt_state,evt_kind,brand
Dim strparm , shopid , i,page , partMDid , issale
	selDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	evt_startdate 	= requestCheckVar(Request("evt_startdate"),10)
	evt_enddate 	= requestCheckVar(Request("evt_enddate"),10)
	evt_code 		= requestCheckVar(Request("evt_code"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),60)
	sCategory	= requestCheckVar(Request("selC"),10) 		'카테고리
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'카테고리(중분류)
	evt_state	= requestCheckVar(Request("evt_state"),4)	'이벤트 상태
	evt_kind 	= requestCheckVar(Request("evt_kind"),32)	'이벤트종류
	partMDid  	= requestCheckVar(Request("partMDid"),32)		'담당 MD
	brand		= requestCheckVar(Request("brand"),32)		'브랜드
	isgift	= requestCheckVar(Request("isgift"),1)
	israck	= requestCheckVar(Request("israck"),1)
	isprize	= requestCheckVar(Request("isprize"),1)
	issale	= requestCheckVar(Request("issale"),1)
	shopid		= requestCheckVar(Request("shopid"),32)		'매장
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1

strparm = "menupos="&menupos&"&selDate="&selDate&"&evt_startdate="&evt_startdate&"&evt_enddate="&evt_enddate
strparm = strparm & "&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&evt_state="&evt_state&"&evt_kind="&evt_kind&""
strparm = strparm & "&partMDid="&partMDid&"&brand="&brand&"&isgift="&isgift&"&israck="&israck&"&isprize="&isprize&"&shopid="&shopid

'데이터 가져오기
dim cEvtList
set cEvtList = new cevent_list
cEvtList.FPageSize = 50
cEvtList.FCurrPage = page
cEvtList.FrectSfDate 	= selDate		'기간 검색 기준
cEvtList.frectevt_startdate 	= evt_startdate	'검색 시작일
cEvtList.frectevt_enddate 	= evt_enddate	'검색 종료일
cEvtList.FrectSfEvt 	= evt_code		'검색 이벤트명 or 이벤트코드
cEvtList.FrectSeTxt 	= strTxt	'검색어
cEvtList.FrectScategory = sCategory	'검색 카테고리
cEvtList.FrectScateMid	= sCateMid	'검색 카테고리(중분류)
cEvtList.frectevt_state 	= evt_state	'검색 상태
cEvtList.frectpartMDid   	= partMDid
cEvtList.frectevt_kind 	= evt_kind
cEvtList.frectbrand 	= brand
cEvtList.frectissale 	= issale
cEvtList.frectisgift 	= isgift
cEvtList.frectisrack 	= israck
cEvtList.frectisprize 	= isprize
cEvtList.frectshopid = 	shopid
cEvtList.fnGetEventList_off()

Dim arreventlevel, arreventstate, arreventkind
'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
arreventstate= fnSetCommonCodeArr_off("evt_state",False)
arreventkind= fnSetCommonCodeArr_off("evt_kind",False)
%>

<script language="javascript">

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.submit();
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


	if(frm.evt_code.value== "evt_code"&&frm.sEtxt.value!=""){
		if(!IsDigit(frm.sEtxt.value)){
			alert("이벤트코드는 숫자만 가능합니다.");
			frm.sEtxt.focus();
			return;
		}
	}

		frm.submit();
	}

	//코드관리
	function jsCodeManage(){
		var winCode;
		winCode = window.open('/admin/offshop/code/popManageCode.asp','popCode','width=400,height=600');
		winCode.focus();
	}

	//당첨자등록
	function prize(evt_code){
		 var prize = window.open('pop_event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
		 prize.focus();
	}

	//수정
	function event_edit(evt_code){
		var event_edit = window.open('event_modify.asp?evt_code='+evt_code,'event_edit','width=1024,height=768,scrollbars=yes,resizable=yes');
		event_edit.focus();
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch('E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td>
				이벤트종류:<%sbGetOptCommonCodeArr_off "evt_kind", evt_kind, true,True,"onChange='javascript:document.frmEvt.submit();'"%>
				진행상태:<%sbGetOptCommonCodeArr_off "evt_state", evt_state, True,False,"onChange='javascript:document.frmEvt.submit();'"%>
				담당MD: <%sbGetMDid_off "partMDid",partMDid, "onChange='javascript:document.frmEvt.submit();'"%>
				매장: <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			</td>
		</tr>
		<tr>
			<td>
		    	이벤트타입:
		    	<!-- 할인<input type="checkbox" name="issale" value="Y" onclick="jsSearch('E');" <% if issale = "Y" then response.write " checked"%>> -->
		    	사은품<input type="checkbox" name="isgift" value="Y" onclick="jsSearch('E');" <% if isgift = "Y" then response.write " checked"%>>
		    	매대<input type="checkbox" name="israck" value="Y" onclick="jsSearch('E');" <% if israck = "Y" then response.write " checked"%>>
		    	당첨<input type="checkbox" name="isprize" value="Y" onclick="jsSearch('E');" <% if isprize = "Y" then response.write " checked"%>>
				<select name="selDate">
		    		<option value="S" <%if Cstr(selDate) = "S" THEN %>selected<%END IF%>>기간(시작일기준)</option>
		    		<option value="E" <%if Cstr(selDate) = "E" THEN %>selected<%END IF%>>기간(종료일기준)</option>
		        </select>
				<input id="evt_startdate" name="evt_startdate" value="<%=evt_startdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="evt_enddate" name="evt_enddate" value="<%=evt_enddate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "evt_startdate", trigger    : "evt_startdate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "evt_enddate", trigger    : "evt_enddate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				<select name="evt_code">
		    		<option value="evt_code" <%if Cstr(evt_code) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
		    		<option value="evt_name" <%if Cstr(evt_code) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
		    	</select>
		        <input type="text" name="sEtxt" value="<%=strTxt%>" maxlength="60">
		        <br>브랜드:<% drawSelectBoxDesignerwithName "brand", brand %>
		        <!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
			</td>
		</tr>

		</table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch('E');">
	</td>
</tr>
</form>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
	<tr height="40" valign="bottom">
    <td align="left">
    	※ 하루에 한번 새벽에 상태값 오픈대기의 경우 오픈으로 자동변경되며, 오픈상태인데 날짜가 지난경우 자동종료 됩니다.
    </td>
    <td align="right">
		<input type="button" value="새로등록" onclick="event_edit('');" class="button">
		<% if C_ADMIN_AUTH then %>
			<input type="button" value="코드관리" onclick="jsCodeManage();"  class="button">
		<%END IF%>
    </td>
	</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<% if cEvtList.FresultCount>0 then %>
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="13">검색결과 : <b><%=cEvtList.FTotalCount%></b>&nbsp;&nbsp;페이지 : <b><%= page %>/ <%= cEvtList.FTotalPage %></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td >이벤트코드</td>
	<td >기본이미지</td>
	<td >매장</td>
	<td >이벤트종류</td>
  	<td >진행상태</td>
  	<td >이벤트명</td>
  	<td >이벤트타입</td>
  	<td width="60">시작일</td>
  	<td width="60">종료일</td>
  	<td >담당MD</td>
  	<td >비고</td>
</tr>
<%
	For i = 0 To cEvtList.FResultCount - 1
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="silver"; onmouseout=this.style.background='#FFFFFF'; align="center">
  	<td>
  		<%=cEvtList.FItemList(i).fevt_code%>
  	</td>
  	<td>
  		<% if cEvtList.FItemList(i).fimg_basic <> "" then %>
  			<img src="<%=cEvtList.FItemList(i).fimg_basic%>" width=50 height=50 onclick="jsImgView('<%=cEvtList.FItemList(i).fimg_basic%>');" alt="누르시면 확대 됩니다">
  		<% end if %>
  	</td>
  	<td width="100">
  		<%
  		if cEvtList.FItemList(i).fshopid = "all" then
  			response.write "전체매장"
  		else
  			response.write cEvtList.FItemList(i).fshopname
  		end if

  		IF cEvtList.FItemList(i).faddShopCnt>0 THEN
  		    rw ""
  		    set cEvtAddedShop = new cevent_list
		    cEvtAddedShop.frectevt_code = cEvtList.FItemList(i).fevt_code
		    cEvtAddedShop.getAddedShopList

		    for j=0 to cEvtAddedShop.FResultCount-1
		        rw cEvtAddedShop.FItemList(j).FShopName
		    next
		    set cEvtAddedShop=Nothing
  		END IF
  		%>
  	</td>
	<td>
		<%=fnGetCommCodeArrDesc_off(arreventkind,cEvtList.FItemList(i).fevt_kind)%>
	</td>
  	<td>
		<%
		'/오픈
		IF cEvtList.FItemList(i).fevt_state = 7 THEN
		%>
			<font color="blue"><%=fnGetCommCodeArrDesc_off(arreventstate,cEvtList.FItemList(i).fevt_state)%></font>
		<%
		'/종료
		elseIF cEvtList.FItemList(i).fevt_state = 9 THEN
		%>
			<font color="gray"><%=fnGetCommCodeArrDesc_off(arreventstate,cEvtList.FItemList(i).fevt_state)%></font>
		<%
		'/오픈요청 , 종료요청
		elseIF cEvtList.FItemList(i).fevt_state = 5 or cEvtList.FItemList(i).fevt_state = 8 THEN
		%>
			<font color="red"><%=fnGetCommCodeArrDesc_off(arreventstate,cEvtList.FItemList(i).fevt_state)%></font>
		<% else %>
			<%=fnGetCommCodeArrDesc_off(arreventstate,cEvtList.FItemList(i).fevt_state)%>
		<% end if %>
  	</td>
  	<td>
  		<%=cEvtList.FItemList(i).fevt_name%>
  	</td>
  	<td>
  		<%
  			if cEvtList.FItemList(i).fissale = "Y" then
  				response.write " <img src='http://fiximage.10x10.co.kr/web2008/category/icon_sale.gif'> "
  			end if
  			if cEvtList.FItemList(i).fisgift = "Y" then
  				response.write " <img src='http://fiximage.10x10.co.kr/web2008/category/icon_gift.gif'> "
  			end if
  			if cEvtList.FItemList(i).fisrack = "Y" then
  				response.write " 매대("&cEvtList.FItemList(i).fisracknum&") "
  			end if

  			if cEvtList.FItemList(i).fisprize = "Y" then
  				response.write " 당첨 "
  			end if
  		%>
  	</td>
  	<td><%=cEvtList.FItemList(i).fevt_startdate%></td>
	<td><%=cEvtList.FItemList(i).fevt_enddate%></td>
  	<td><%=cEvtList.FItemList(i).fmdname%></td>
  	<td>
  		<input type="button" value="수정" onclick="event_edit(<%= cEvtList.FItemList(i).fevt_code %>);" class="button">
  		<input type="button" value="상품(<%= cEvtList.FItemList(i).fitem_count %>)" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?evt_code=<%= cEvtList.FItemList(i).fevt_code %>&<%= strparm %>')">
  		<%' if cEvtList.FItemList(i).fissale = "Y"  then %>
  			<!--<input type="button" value="할인(<%'= cEvtList.FItemList(i).fsale_count%>)" class="button" onClick="jsGoUrl('/admin/offshop/sale/salelist.asp?ec=<%'= cEvtList.FItemList(i).fevt_code %>&menupos=1251');">-->
  		<%' end if %>
  		<% if cEvtList.FItemList(i).fisprize = "Y" then %>
  			<input type="button" value="당첨(<%= cEvtList.FItemList(i).fprize_count %>)" class="button" onClick="jsGoUrl('eventprize_regist.asp?evt_code=<%= cEvtList.FItemList(i).fevt_code %>&<%= strparm %>')">
  		<%END IF%>
  		<% if cEvtList.FItemList(i).fisgift = "Y" then %>
  			<input type="button" onClick="javascript:jsGoUrl('/admin/offshop/gift/giftlist.asp?evt_code=<%= cEvtList.FItemList(i).fevt_code %>&<%= strparm %>');" value="사은품(<%= cEvtList.FItemList(i).fgift_count %>)" class="button">
  		<%end if%>
  		<% if cEvtList.FItemList(i).fisprize = "Y" then %>
  			<!--<input type="button" value="당첨자등록(<%= cEvtList.FItemList(i).fprizeyn %>)" onclick="prize(<%= cEvtList.FItemList(i).fevt_code %>);" class="button">-->
  		<% End IF %>
  	</td>
</tr>
<%	Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if cEvtList.HasPreScroll then %>
			<span class="list_link"><a href="?<%=strparm%>&page=<%=cEvtList.StartScrollPage-1%>&evt_code=<%=evt_code%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + cEvtList.StartScrollPage to cEvtList.StartScrollPage + cEvtList.FScrollCount - 1 %>
			<% if (i > cEvtList.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cEvtList.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?<%=strparm%>&page=<%=i%>&evt_code=<%=evt_code%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if cEvtList.HasNextScroll then %>
			<span class="list_link"><a href="?<%=strparm%>&page=<%=i%>&evt_code=<%=evt_code%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% ELSE %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="11">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
