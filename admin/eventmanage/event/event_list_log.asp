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
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
	Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅
	
	'변수선언
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
	
	Dim iIsSale,iIsGift,iIsCoupon,fstchk
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
	
	if sSdate="" then sSdate= dateserial(year(now()),month(now()),day(now()))
	if sEdate="" then sEdate = dateserial(year(now()),month(now()),day(now()))
	sEvt = Request("selEvt")  '이벤트 코드/명 검색
	strTxt = Request("sEtxt")
	
	sCategory	= Request("selC") '카테고리
	sState	 = Request("eventstate")'이벤트 상태	
	sKind = Request("eventkind")	'이벤트종류
	
	if sState ="" then sState="7"
	
	fstchk = request("fstchk")	
	iIsSale = request("iIsSale")
	iIsGift = request("iIsGift")
	iIsCoupon = request("iIsCoupon")
	
	if fstchk="" then 
		
		if iIsGift="" then 
			iIsGift="on"		
		end if 	
	end if
	
	

		
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
		
		cEvtList.FSisSale = iIsSale
		cEvtList.FSisGift = iIsGift
		cEvtList.FSisCoupon = iIsCoupon
		
 		arrList = cEvtList.fnGetEventList_LOG	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing
 	
	iTotalPage 	=  Int(iTotCnt/iPageSize)	'전체 페이지 수
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1		
		
	function setChkStr(st)
		if st="on" then setChkStr="checked"
	end function
	
	function getGiftItems(evtcode)
		dim sql
		sql =" select top 10 gift_itemname " &_
				" from db_event.dbo.tbl_gift " &_
				" where evt_code='"&evtcode&"'"
				
		rsget.open sql,dbget,1
		
		if not rsget.eof then
			response.write rsget("gift_itemname")
		end if
		rsget.close
	end function

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
//-->

function DrawReceiptPrintobj_TEC(elementid,printname){
        var objstring = "";
        var e;
        objstring = '<OBJECT name="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
        objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
        objstring = objstring + ' width=0 ';
        objstring = objstring + ' height=0 ';
        objstring = objstring + ' align=center ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + ' > ';
        objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
        objstring = objstring + ' </OBJECT>';
        
        document.write(objstring);
}

function eventindexprint(ievt_code, ievt_name01, ievt_name02, ievt_startdate, ievt_enddate, ievt_gift01, ievt_gift02, ievt_gift03){
	var X = 1;
	var Y = 1;
	var F = 1;
	
	// TEC_DO3 : 452
	if (TEC_DO3.IsDriver == 1){
           X = 1.05;
           Y = 1.05;
           F = 1.2;
           
			TEC_DO3.SetPaper(900,600);
			TEC_DO3.OffsetX = 20;
			TEC_DO3.OffsetY = 20;
			TEC_DO3.PrinterOpen();
			
			
			
			

			TEC_DO3.PrintText(50*X, 50*Y, "HY견고딕", 30*F, 0, 0, "[시작일]");
			TEC_DO3.PrintText(250*X, 50*Y, "HY견고딕", 30*F, 1, 0, ievt_startdate);
			
			TEC_DO3.PrintText(50*X, 100*Y, "HY견고딕", 30*F, 0, 0, "[종료일]");
			TEC_DO3.PrintText(250*X, 100*Y, "HY견고딕", 30*F, 1, 0, ievt_enddate);
			
			TEC_DO3.PrintText(500*X, 30*Y, "Arial Bold", 150*F, 0, 0, ievt_code);
			
			TEC_DO3.PrintText(50*X, 175*Y, "HY견고딕", 30*F, 0, 0, "[이벤트명]");
			TEC_DO3.PrintText(50*X, 225*Y, "HY견고딕", 30*F, 1, 0, ievt_name01);
			TEC_DO3.PrintText(50*X, 275*Y, "HY견고딕", 30*F, 1, 0, ievt_name02);
			
			TEC_DO3.PrintText(50*X, 350*Y, "HY견고딕", 30*F, 0, 0, "[사은품명]");
			TEC_DO3.PrintText(50*X, 400*Y, "HY견고딕", 30*F, 1, 0, ievt_gift01);	<!-- 한줄에 한글 24자까지 넘으면 아래로 -->
			TEC_DO3.PrintText(50*X, 450*Y, "HY견고딕", 30*F, 1, 0, ievt_gift02);
			TEC_DO3.PrintText(50*X, 500*Y, "HY견고딕", 30*F, 1, 0, ievt_gift03);
			
			TEC_DO3.PrinterClose();


    }else window.status = "TEC B-452 드라이버를 설치해 주세요"
}

DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-452");


</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  action="" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="fstchk" value="on">
	<input type="hidden" name="iC">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			이벤트종류:<%sbGetOptEventCodeValue "eventkind", sKind, True,"onChange='javascript:document.frmEvt.submit();'"%>
			&nbsp;
			카테고리:<% sbGetOptCategoryLarge "selC", sCategory ,"onChange='javascript:document.frmEvt.submit();'" %>
			&nbsp;
			진행상태:<%sbGetOptEventCodeValue "eventstate", sState, True,"onChange='javascript:document.frmEvt.submit();'"%>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			이벤트타입:
			<input type="checkbox" name="iIsSale" <%= setChkStr(iIsSale) %>>할인
			<input type="checkbox" name="iIsGift" <%= setChkStr(iIsGift) %>>사은품
			<input type="checkbox" name="iIsCoupon" <%= setChkStr(iIsCoupon) %>>쿠폰
			&nbsp;
			코드/명:
			<select class="select" name="selEvt">
    			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
    			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
			</select>
			<input type="text" class="text" name="sEtxt" value="<%=strTxt%>">
			&nbsp;
			기간:
    	 	<!--
    	 	<select name="selDate">        	 	 	
    			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
    			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
    		</select>
    		-->        		
    		<input type="text" class="text" size="11" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:pointer;">
    		 ~ <input type="text" class="text" size="11" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:pointer;">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>
<input type="button" value="전체보기" onClick="jsSearch(document.frmEvt, 'A')">

<p>
<!-- 표 중간바 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">       	
			추가검색조건 : 현재진행중인 이벤트(현재일이 시작일과 종료일 사이에 있어야 하며, 강제종료 안된이벤트)<br>
			진행상태에 하나 더 추가 --> 반품완료(사은품)
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</table>
<!-- 표 중간바 끝 -->

<!-- 표 중간바 시작
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">       	
       	<img src="/images/icon_new_registration.gif" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>');" style="cursor:hand;">     
    	</td>
    	<td align="right">
       	<input type="button" value="스케쥴" onclick="jsSchedule();">       	
       <!--	<input type="button" value="통계" onclick=" ">  -->
       <!--	정렬: <select name="selSort">
       	<option value="1">이벤트코드내림차순</option>
       	
       	</select>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</table>
표 중간바 끝 -->
 
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">이벤트<br>코드</td>
      	<td width="90">진행상태</td>
      	<td width="100">종류</td>
      	<!--<td width="60">타입</td>-->
      	<td>이벤트명</td>      	
      	<td width="100">배너이미지</td>   
      	<td width="120">카테고리</td>
      	<td width="60">시작일</td>
      	<td width="60">종료일</td>
      	<td width="80">관리</td>
      	<td >사은품 </td>
      	<td width="60">인덱스출력</td>
    </tr>
    <%IF isArray(arrList) THEN 
    	For intLoop = 0 To UBound(arrList,2)
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="<%=vwwwUrl%>/event/eventmain.asp?eventid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventstate",arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventkind",arrList(1,intLoop))%></a></td>
      	<!--<td></td>-->
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=db2html(arrList(4,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(10,intLoop) <> "" THEN%> <img src="<%=arrList(10,intLoop)%>" width="100" border="0"><%END IF%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(12,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td>
      	<td><input type="button" value="화면" class="button" onClick="javascript:jsGoUrl('event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<input type="button" value="상품" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="당첨" class="input_b" onClick="jsGoUrl('eventprize_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')"><%END IF%>
      	</td>
    	<td ><%= getGiftItems(arrList(0,intLoop)) %> </td>
    	<td><input type="button" value="출력" class="button" onClick="eventindexprint('<%=arrList(0,intLoop)%>', '<%= left(db2html(arrList(4,intLoop)),24) %>','<%= mid(db2html(arrList(4,intLoop)),25) %>','<%=arrList(5,intLoop)%>','<%=arrList(6,intLoop)%>','<%= left(getGiftItems(arrList(0,intLoop)),24) %>','<%= mid(getGiftItems(arrList(0,intLoop)),25,24) %>','<%= mid(getGiftItems(arrList(0,intLoop)),49) %>')"></td>
    </tr>   
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="11">등록된 내용이 없습니다.</td>
   	</tr>	
   <%END IF%>
   
    <!-- 페이징처리 -->
    <%		
	iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
	
	If (iCurrpage mod iPerCnt) = 0 Then																
		iEndPage = iCurrpage
	Else								
		iEndPage = iStartPage + (iPerCnt-1)
	End If	
	%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
<!-- #include virtual="/lib/db/dbclose.asp" --> -->