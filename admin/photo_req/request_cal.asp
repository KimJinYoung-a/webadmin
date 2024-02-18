<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/sheduleCls.asp"-->
<%
'################################# 달력시작 #################################
Dim ThisDate,ThisYear,ThisMonth,ThisDay,ThisToday
Dim PrevYear,PrevMonth,NextYear,NextMonth
Dim FirstDay,FirstWeekDay
Dim PrintDay,LastDay,LoopWeek,LoopDay,Stop_Flag
Dim 	PrevThisYearDate,NextThisYearDate,PrevThisMonthDate,NextThisMonthDate
Dim NowToday,NowThisDay
Dim r_photo
	r_photo = request("req_photo")
	ThisDate = Request("ThisDate")		' 넘어온 값 받기 년 - 월 - 일
	ThisYear = Request("ThisYear")		' 넘어온 값의 년도
	ThisMonth = Request("ThisMonth")		' 넘어온 값의 월
	ThisDay = Request("ThisDay")		' 넘어온 값의 일

	menupos = request("menupos")		'메뉴고유번호

	IF ThisDate = "" THEN			' 넘어온 값이 없다면 (처음 페이지를 시작했을 경우)
		ThisToday = DATE()		' 현재는 오늘이 됨 (오늘의 년 - 월 - 일)
		ThisYear = YEAR(ThisToday)		' 오늘의 년도
		ThisMonth = MONTH(ThisToday)	' 오늘의 월
		ThisDay = DAY(ThisToday)		' 오늘의 일
	ELSE					' 넘어온 값이 있다면 (해당 년도나 월을 찾을경우)
		ThisToday = CDATE(ThisDate)	' 넘어온 값을 저장
		ThisYear = YEAR(ThisToday)		' 넘어온 값의 년도
		ThisMonth = MONTH(ThisToday)	' 넘어온 값의 월
		ThisDay = DAY(ThisToday)		' 넘어온 값의 일
	END IF

	IF ThisMonth = 1 THEN			' 작년, 내년, 저번달, 다음달 구하기 (1월, 12월, 나머지 달들)
		PrevYear = ThisYear -1
		PrevMonth = 12
		NextYear = ThisYear +1
		NextMonth = 2
	ELSEIF ThisMonth = 12 THEN
		PrevYear = ThisYear -1
		PrevMonth = 11
		NextYear = ThisYear +1
		NextMonth= 1
	ELSE
		PrevYear = ThisYear -1
		PrevMonth = ThisMonth -1
		NextYear = ThisYear +1
		NextMonth = ThisMonth +1
	END IF

	FirstDay = DateSerial(ThisYear,ThisMonth,1)	' 넘겨받은 날의 월 초기값 (년-월-1)
	FirstWeekDay = WeekDay(FirstDay,vbSunday)	' 첫날의 요일을 구함 (월요일기준 : 값 2)

	PrintDay = 1				' 출력 초기값은 1

	IF ThisMonth = 4 OR ThisMonth =6 OR ThisMonth = 9 OR ThisMonth = 11 THEN		' 현재 달의 월말 값 계산
		LastDay = 30
	ELSEIF ThisMonth = 2 AND NOT (ThisYear MOD 4) = 0 THEN
		LastDay = 28
	ELSEIF ThisMonth = 2 AND (ThisYear MOD 4) = 0 THEN
		IF (ThisYear MOD 100) = 0 THEN
			IF (ThisYear MOD 400) = 0 THEN
				LastDay = 29
			ELSE
				LastDay = 30
			END IF
		ELSE
			LastDay = 29
		END IF
	ELSE
		LastDay = 31
	END IF

	ThisDate  = CDATE(ThisYear&"-"&ThisMonth&"-"&ThisDay)			' 페이지 이동시 넘겨질 값들
	PrevThisYearDate = PrevYear&"-"&ThisMonth&"-"&ThisDay
	NextThisYearDate =  NextYear&"-"&ThisMonth&"-"&ThisDay
	PrevThisMonthDate = ThisYear&"-"&PrevMonth&"-01"
	NextThisMonthDate = ThisYear&"-"&NextMonth&"-01"

	NowToday = DATE()						' 오늘 날짜(일) 값 구하기
	NowThisDay = DAY(NowToday)

'################################# 달력끝 #################################
Dim iz, osemi

Dim Thisyyyymm
Thisyyyymm = left(CDATE(ThisYear&"-"&ThisMonth&"-"&ThisDay),7)

Set osemi = new CSeminarRoomCalendar
osemi.FRectYYYYMM = Thisyyyymm
osemi.FReq_photo = r_photo
osemi.list
%>
<!-- 아래는 스크립트보다 위에 있어야 한다. -->
<div ID="viewDIV" STYLE="position:absolute; visibility:hide;"></div>
<!-- 위 한줄은 스크립트보다 위에 있어야 한다. -->
<script language="JavaScript">
<!--
var fcolor = "#ffffff";        // Main background color
var textcolor = "#000000";        // Text color
var border_size = "1";                // border size, 1-3
var border_color = "#000000";        // Border color
var width = "300";                // 팝업 박스의 넓이, 100 - 300
var palign = 0;                // 팝업 박스의 위치, 0:center/1:right/2:left

ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false
ie5 = false;

// Microsoft Stupidity Check.
if (ie4) {
        if (navigator.userAgent.indexOf('MSIE 5')>0) {
                ie5 = true;
        } else {
                ie5 = false; }
} else {
        ie5 = false;
}

var x = 0;
var y = 0;
var offsetx = 10;
var offsety = 10;
var popup_on = 0;
var over;

if ( (ns4) || (ie4) ) {
    if (ns4) over = document.viewDIV;
    if (ie4) over = viewDIV.style;
    document.onmousemove = mouseMove;
    if (ns4) document.captureEvents(Event.MOUSEMOVE);
}
// Clears popups if appropriate
function viewoff() {
    if ( (ns4) || (ie4) ) {
            popup_on = 0;
            hideObject(over);
    }
}
function myColor(num) {
	if (document.all) {
		eval('document.all.cell'+num+'.style.background = "#E6E6F2"');
	}
}

function myColorOut(num) {
	if (document.all) {
		eval('document.all.cell'+num+'.style.background = "#FFFFFF"');
	}
}

function GoPhotoreq(thisday) {
    location.href = "request_cal_day.asp?getday=" + thisday;
}

// Simple popup
function viewon(text) {
    var txt = "<TABLE WIDTH="+width+" STYLE='filter:alpha(opacity=100); border:0 ' BORDER=0 CELLPADDING="+border_size+" CELLSPACING=0 BGCOLOR="+border_color+"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR="+fcolor+"><TR><TD><FONT FACE='Arial,Helvetica' COLOR="+textcolor+" SIZE='2'>"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>"
    layerWrite(txt);
    disp();
}

// Common calls
function disp() {
    if ( (ns4) || (ie4) ) {
        if (popup_on == 0)         {
            if (palign == 0) { // Center
                    moveTo(over,x+offsetx-(width/2),y+offsety);
            }
            if (palign == 1) { // Right
                    moveTo(over,x+offsetx,y+offsety);
            }
            if (palign == 2) { // Left
                    moveTo(over,x-offsetx-width,y+offsety);
            }
            showObject(over);
            popup_on = 1;
        }
    }
// Here you can make the text goto the statusbar.
}
//-->

// Moves the layer
function mouseMove(e) {
    if (ns4) {x=e.pageX; y=e.pageY;}
    if (ie4) {x=event.x; y=event.y;}
    if (ie5) {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
    if (popup_on) {
        if (palign == 0) { // Center
                moveTo(over,x+offsetx-(width/2),y+offsety);
        }
        if (palign == 1) { // Right
                moveTo(over,x+offsetx,y+offsety);
        }
        if (palign == 2) { // Left
                moveTo(over,x-offsetx-width,y+offsety);
        }
    }
}

// Writes to a layer
function layerWrite(txt) {
    if (ns4) {
        var lyr = document.viewDIV.document
        lyr.write(txt)
        lyr.close()
    }
    else if (ie4) document.all["viewDIV"].innerHTML = txt
}

// Make an object visible
function showObject(obj) {
    if (ns4) obj.visibility = "show"
    else if (ie4) obj.visibility = "visible"
}

// Hides an object
function hideObject(obj) {
    if (ns4) obj.visibility = "hide"
    else if (ie4) obj.visibility = "hidden"
}

// Move a layer
function moveTo(obj,xL,yL) {
    obj.left = xL
    obj.top = yL
}
function fnSearch(){
	var frm = document.sFrm;
	frm.submit();
}
</script>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td>
		<form name="sFrm" method="get" action="?">
		<input type="hidden" name="ThisDate" value="<%=ThisDate%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="r_photo" value="<%=r_photo%>">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="font" height="25">
		<tr>
			<td align="right"><% call SelectUser("req_photo", ""&r_photo&"") %></td>
		</tr>
		<tr align="center" valign="middle">
			<td height="25">
				<A HREF="request_cal.asp?ThisDate=<%=PrevThisYearDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">◀</A>&nbsp;&nbsp;<%=ThisYear%>&nbsp;&nbsp;
				<A HREF="request_cal.asp?ThisDate=<%=NextThisYearDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">▶</A>&nbsp;
				<A HREF="request_cal.asp?ThisDate=<%=PrevThisMonthDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">◀</A>&nbsp;&nbsp;<%=ThisMonth%>&nbsp;&nbsp;
				<A HREF="request_cal.asp?ThisDate=<%=NextThisMonthDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">▶</A>
			</td>
		</tr>
		</table>
		</form>
		<br>
		<table border="1" cellspacing="0" cellpadding="0"  bordercolordark="White" bordercolorlight="black">
		<tr>
			<td width="120" height="80" align="center" class="verdana-mid"><font color="#FF0000">일</font></td>
			<td width="120" height="80" align="center" class="verdana-mid">월</td>
			<td width="120" height="80" align="center" class="verdana-mid">화</td>
			<td width="120" height="80" align="center" class="verdana-mid">수</td>
			<td width="120" height="80" align="center" class="verdana-mid">목</td>
			<td width="120" height="80" align="center" class="verdana-mid">금</td>
			<td width="120" height="80" align="center" class="verdana-mid"><font color=#8FB5DA>토</font></td>
		</tr>
<%
	FOR LoopWeek = 1 TO 6						' 주 단위 LOOP  최대 6주
		Response.Write "<tr>"&vbCR			' vbCR = ch<13>  : 엔터
		FOR LoopDay = 1 TO 7					' 요일 LOOP 시작 : 월요일
			IF FirstWeekDay > 1 THEN					' 현재 요일을 나타내는값이 1(일요일)보다 클경우 공백을 만들어주기 위한 코드 ; 테이블 칸수를 맞추기 위해
				Response.Write "<td height=80 align=left valign=top>&nbsp;"
				FirstWeekDay = FirstWeekDay - 1
			ELSE
				IF PrintDay > LastDay THEN				' 현재 일이 월말보다 클경우 (공백처리) ; 테이블 칸수를 맞추기 위해
  					Response.Write "<td height=80 align=left valign=top>"
  				ELSE
  					IF PrintDay = NowThisDay THEN		' 현재 일이 오늘일 경우 (색깔과 크기를 다르게 지정)
  						Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " style='cursor:hand' onclick=GoPhotoreq('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#990000><b>"&PrintDay&"</b></font></td>"
						For iz = 0 to osemi.FResultCount - 1

							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
								</tr><tr><td class=verdana-mid onMouseOver = "viewon('상태 : <%=osemi.FItemList(iz).FReqStatus%><br> 상품명 : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> 부서 :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"파트","") %><br> 요청자 : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End If
						Next
						response.write "</table>"
					ELSEIF LoopDay = 1 THEN			' 일요일은 색깔 다르게
%>
  						<td height=80 align=left valign=top id=cell<%=Cstr(PrintDay)%> style='cursor:hand' onclick=GoPhotoreq(<%=Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay)))%>)><table class=verdana-small><tr><td><font color=#FF0000><%=PrintDay%></font></td></tr>
<%
						For iz = 0 to osemi.FResultCount - 1
							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
							</tr><tr><td class=verdana-mid onMouseOver = "viewon('상태 : <%=osemi.FItemList(iz).FReqStatus%><br> 상품명 : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> 부서 :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"파트","") %><br> 요청자 : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End If
						Next
						response.write "</table>"
					ELSEIF LoopDay = 7 THEN			' 토요일은 색깔 다르게
  						Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " style='cursor:hand' onclick=GoPhotoreq('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#3366CC>"&PrintDay&"</font></td></tr>"
						For iz = 0 to osemi.FResultCount - 1
							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
							</tr><tr><td class=verdana-mid onMouseOver = "viewon('상태 : <%=osemi.FItemList(iz).FReqStatus%><br> 상품명 : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> 부서 :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"파트","") %><br> 요청자 : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End If
						Next
						response.write "</table>"
	  				ELSE					' 나머지 일들을 테이블 안에 넣어줌
  						Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " style='cursor:hand' onclick=GoPhotoreq('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td>"&PrintDay & "</td></tr>"
						For iz = 0 to osemi.FResultCount - 1
							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
								</tr><tr><td class=verdana-mid onMouseOver = "viewon('상태 : <%=osemi.FItemList(iz).FReqStatus%><br> 상품명 : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> 부서 :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"파트","") %><br> 요청자 : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End if
						Next
						response.write "</table>"
	  				END IF
  				END IF
	  			PrintDay = PrintDay + 1				' 일수를 1씩 증가
				IF PrintDay>LastDay THEN  				' 월말보다 Stop_Flag = 1 이라고 지정
					Stop_Flag=1
				END IF
  			END IF
  			Response.Write "</td>" &vbCR
  		NEXT							' 일단위 (7칸) 루프
  		Response.Write "</tr>" &vbCR
		IF Stop_Flag=1 THEN					' 6주까지 안가고 Stop_Flag = 1 이면 주단위 루프 끝냄 / 이부분 생략하면 달력 박스가 일정하게 유지됨
			EXIT FOR
		END IF
  	NEXT								' 주단위 (최대 6줄) 루프
%>
		</table>
	</td>
</tr>
</table>
<%set osemi = Nothing%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->