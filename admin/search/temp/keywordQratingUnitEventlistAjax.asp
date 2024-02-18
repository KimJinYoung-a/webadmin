<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
If not (Request.ServerVariables("REMOTE_ADDR") = "61.252.133.75" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.105" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.106") Then
	dbget.close
	Response.End
End If

Response.CharSet = "euc-kr"

	'파라미터값 받기 & 기본 변수 값 세팅
	Dim cEvtList, page, iPageSize, iPerCnt, isResearch, sDate, sSdate, sEdate, sEvt, strTxt, sKind, dispCate, strparm
	Dim arrList, iTotCnt, iTotalPage, sSort, blnReqPublish, intLoop, arreventkind, maxDepth, iStartPage, iEndPage, ix
	page = NullFillWith(requestCheckVar(Request("page"),10),1)	'현재 페이지 번호
	iPageSize = 10		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	maxDepth = 2
	'response.write page

	isResearch = NullFillWith(requestCheckVar(Request("isResearch"),1),"0")
	sDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)
	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),60)
	sKind 		= requestCheckVar(Request("eventkind"),32)	'이벤트종류
	dispCate 	= requestCheckvar(request("disp"),16)

	arreventkind= fnSetCommonCodeArr("eventkind",False)
	
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
 	    
	'데이터 가져오기
	set cEvtList = new ClsEvent
		cEvtList.FCPage = page		'현재페이지
		cEvtList.FPSize = iPageSize		'한페이지에 보이는 레코드갯수
		cEvtList.FSfDate 	= sDate		'기간 검색 기준
		cEvtList.FSsDate 	= sSdate	'검색 시작일
		cEvtList.FSeDate 	= sEdate	'검색 종료일
		cEvtList.FSfEvt 	= sEvt		'검색 이벤트명 or 이벤트코드
		cEvtList.FSeTxt 	= strTxt	'검색어
		cEvtList.FEDispCate	= dispCate	'검색 전시카테고리
		cEvtList.FSkind 	= sKind
		cEvtList.FIsReqPublish = blnReqPublish
		cEvtList.FSort          = sSort
		
 		arrList = cEvtList.fnGetEventList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<form id="eventfrm" name="eventfrm" method="get" style="margin:0px;">
<input type="hidden" id="page" name="page">
<input type="hidden" name="isResearch" value="1"> 
<input type="hidden" name="sSort" value="<%=sSort%>">
<div class="searchWrap" style="border-top:none;">
	<div class="search">
		<ul>
			<li>
				<label class="formTit">기간 :</label>
				<select class="formSlt" id="selDate" name="selDate" title="옵션 선택">
			    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
			    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
			    	<option value="O" <%if Cstr(sDate) = "O" THEN %>selected<%END IF%>>오픈일 기준</option>
				</select>
				<input type="text" class="formTxt" id="iSD" name="iSD" value="<%=sSdate%>" style="width:100px" placeholder="시작일" maxlength="10" readonly />
				<img src="/images/admin_calendar.png" id="iSD_trigger" alt="달력으로 검색" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				~
				<input type="text" class="formTxt" id="iED" name="iED" value="<%=sEdate%>" style="width:100px" placeholder="종료일" maxlength="10" readonly />
				<img src="/images/admin_calendar.png" id="iED_trigger" alt="달력으로 검색" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</li>
		</ul>
	</div>
	<dfn class="line"></dfn>
	<div class="search">
		<ul>
			<li>
				<p class="formTit">이벤트 종류 :</p>
				<%sbGetOptCommonCodeArr "eventkind", sKind, True,True,"onChange='javascript:document.frmEvt.submit();'"%>
			</li>
			<li>
				<p class="formTit">카테고리 :</p>
				<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
			</li>
		</ul>
	</div>
	<dfn class="line"></dfn>
	<div class="search">
		<ul>
			<li>
				<label class="formTit" for="schWord">검색어 :</label>
				<select class="formSlt" id="selEvt" name="selEvt">
					<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
					<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
					<option value="evt_tag" <%if Cstr(sEvt) = "evt_tag" THEN %>selected<%END IF%>>TAG</option>
					<option value="evt_sub" <%if Cstr(sEvt) = "evt_sub" THEN %>selected<%END IF%>>서브카피</option>
				</select>
				<input type="text" class="formTxt" id="sEtxt" name="sEtxt" value="<%=strTxt%>" maxlength="60" style="width:400px" onKeyUp="jsEventValueCheck();" onKeyPress="if (event.keyCode == 13){ NextPage(1,'event'); return false;}" />
			</li>
		</ul>
	</div>
	<input type="button" id="btnsearh1" class="schBtn" value="검색" onClick="NextPage(1,'event');" />
	<input type="button" id="btnsearh2" style="display:none;" class="schBtn" value="검색" onClick="alert('검색중입니다. 잠시만 기다려주세요.');" />
</div>
</form>
<div class="tbListWrap tMar15">
	<div class="rt pad10">
		<span>검색결과 : <strong><%=iTotCnt%></strong></span> <span class="lMar10">페이지 : <strong><%=page%> / <%=iTotalPage%></strong></span>
	</div>
	<ul class="thDataList">
		<li>
			<p class="cell05"></p>
			<p class="cell12">이벤트 코드</p>
			<p class="cell12">이벤트 종류</p>
			<p class="cell12">배너</p>
			<p>이벤트명</p>
			<p class="cell12">카테고리</p>
			<p class="cell12">시작일</p>
			<p class="cell12">종료일</p>
		</li>
	</ul>
	<ul class="tbDataList" id="contentslist">
    <%IF isArray(arrList) THEN

    	For intLoop = 0 To UBound(arrList,2)
    %>
		<li id="tr<%= arrList(0,intLoop) %>" style="cursor:pointer;">
			<p class="cell05"><input type="checkbox" name="contentsidx<%=arrList(0,intLoop)%>" id="contentsidx<%=arrList(0,intLoop)%>" value="<%=arrList(0,intLoop)%>" onClick="jsThisCheck('<%=arrList(0,intLoop)%>','event');" /></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(0,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><img src="<%=arrList(34,intLoop)%>" width="50" height="50" border="0" /></p>
			<p class="lt" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(4,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(26,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(5,intLoop)%></p>
			<p class="cell12" onClick="jsThisClick('<%=arrList(0,intLoop)%>','event');"><%=arrList(6,intLoop)%></p>
		</li>
   <%	Next
   	END IF
   	
	iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1
	
	If (page mod iPerCnt) = 0 Then
		iEndPage = page
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
   %>
	</ul>
	<div class="ct tPad20 bPad20 cBk1">
		<% if (iStartPage-1 )> 0 then %><a href="javascript:NextPage(<%= iStartPage-1 %>,'event')">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(page) then
		%>
			<a href="javascript:NextPage(<%= ix %>,'event')"><span class="cRd1">[<%=ix%>]</span></a>
		<%		else %>
			<a href="javascript:NextPage(<%= ix %>,'event')">[<%=ix%>]</a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:NextPage(<%= ix %>,'event')">[next]</a>
		<% else %>[next]<% end if %>
	</div>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->