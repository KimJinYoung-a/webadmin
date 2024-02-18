<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/index.asp
' Description :  이벤트 등록 - 화면설정
' History : 2007.02.07 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->
<%
	Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

	'변수선언
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
	Dim strparm

	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

	'## 검색 #############################
	sDate = Request("selDate")  '기간
	sSdate = Request("iSD")
	sEdate = Request("iED")

	sEvt = Request("selEvt")  '이벤트 코드/명 검색
	strTxt = Request("sEtxt")

	sCategory	= Request("selC") '카테고리
	sState	 = Request("eventstate")'이벤트 상태
	sKind = Request("eventkind")	'이벤트종류

	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################

	'데이터 가져오기
	set cEvtList = new ClsEvent
		cEvtList.FCPage = iCurrpage	'현재페이지
		cEvtList.FPSize = iPageSize '한페이지에 보이는 레코드갯수

		cEvtList.FSfDate = sDate '기간 검색 기준
		cEvtList.FSsDate = sSdate '검색 시작일
		cEvtList.FSeDate = sEdate '검색 종료일
		cEvtList.FSfEvt = sEvt '검색 이벤트명 or 이벤트코드
		cEvtList.FSeTxt = strTxt '검색어
		cEvtList.FScategory = sCategory '검색 카테고리
		cEvtList.FSstate = sState '검색 상태
		cEvtList.FSkind = sKind

 		arrList = cEvtList.fnGetEventList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing

	iTotalPage 	=  Int(iTotCnt/iPageSize)	'전체 페이지 수
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.submit();
	}

	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

	function jsSearch(frm, sType){
	if (sType == "A"){
			frm.iSD.value = "";
			frm.iED.value = "";
			frm.eventstate.value = "";
			frm.sEtxt.value = "";
			frm.selC.value = "";
		}

		frm.submit();
	}

	function jsSchedule(){
		var winS;
		winS = window.open('pop_event_schedule.asp','popwin','width=800, height=600, scrollbars=yes');
		winS.focus();
	}

	function jsChSelect(iVal){
		alert(iVal);
		alert(document.frmEvt.eventkind.value);
		alert(document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value);
		document.frmEvt.submit();
	}
	function jsStatistics(iVal){
		window.open('pop_evt_Statistics.asp?eC='+ iVal,'sPop','');
	}
	function jsSettingEvt(iVal){
		window.open('pop_evt_Setting.asp?eC='+ iVal,'sPop','');
	}
//-->
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="iC">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
        	<table border="0"  cellpadding="1" cellspacing="3" class="a">
        	<tr>
        		<td width="30" align="right"></td>
        		<td colspan="2">
        			<!-- 이벤트 종류 -->
        			<%sbGetOptEventCodeValue "eventkind", sKind, True,"onChange='javascript:document.frmEvt.submit();'"%>
        			<!-- 카테고리 -->
        			<% sbGetOptCategoryLarge "selC", sCategory ,"onChange='javascript:document.frmEvt.submit();'" %>
        			<!-- 이벤트 종류 -->
        			<%sbGetOptEventCodeStateValue "eventstate", sState, True,"onChange='javascript:document.frmEvt.submit();'"%>
        		</td>
        	</tr>
        	<tr>
        		<td width="30" align="right"></td>
        		<td><select name="selEvt">
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
        			</select>
        			<input type="text" name="sEtxt" value="<%=strTxt%>">
        		&nbsp;&nbsp;<br>기간:
        	 	 <select name="selDate">
        			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
        			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
        		 </select>
        		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:pointer;">
        		 ~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:pointer;">&nbsp;&nbsp;
        		</td>
        		<td  colspan="2" align="right" valign="bottom">&nbsp;&nbsp;
        			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
        			<input type="button" class="button" value="전체보기" onClick="jsSearch(document.frmEvt, 'A')">
        		</td>
        	</tr>
        	</table>
        </td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="35">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
       	<img src="/images/icon_new_registration.gif" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>');" style="cursor:pointer;">
    	</td>
    	<td align="right">
       	<input type="button" class="button" value="스케쥴" onclick="jsSchedule();">
       <!--	<input type="button" value="통계" onclick=" ">  -->
       <!--	정렬: <select name="selSort">
       	<option value="1">이벤트코드내림차순</option>

       	</select>-->
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">이벤트코드</td>
    	<td width="80">진행상태</td>
      	<td width="100">종류</td>
      	<td>이벤트명</td>
      	<td width="80">카테고리</td>
      	<td width="70">시작일</td>
      	<td width="70">종료일</td>
      	<td width="110">통계</td>
    </tr>
    <%IF isArray(arrList) THEN
    	For intLoop = 0 To UBound(arrList,2)
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="<%=wwwUrl%>/event/eventmain.asp?eventid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventstate",arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventkind",arrList(1,intLoop))%></a></td>
      	<td><a href="event_entryList.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=db2html(arrList(4,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(12,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td>
      	<td align="left"><input type="button" value="화면" class="input_b" onClick="javascript:jsGoUrl('event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<input type="button" value="상품" class="input_b" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="당첨" class="input_b" onClick="jsStatistics('<%=arrList(0,intLoop)%>')"><input type="button" class="input_b" value="통계작성" onclick="jsSettingEvt('<%=arrList(0,intLoop)%>');"><%END IF%>

      	</td>
    </tr>
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="8">등록된 내용이 없습니다.</td>
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
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->