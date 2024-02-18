<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<%
	Dim masteridx
	dim i, j
	Dim page, SearchKey, SearchString, part_sn, research, SearchYear, SearchMonth
	dim userid, strMsg

	page = Request("page")
	research = Request("research")
	masteridx = Request("masteridx")

	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")

	SearchYear = Request("SearchYear")
	SearchMonth = Request("SearchMonth")

	part_sn = Request("part_sn")

	if page="" then page=1
	if masteridx="" then masteridx=8

	if SearchYear="" then SearchYear=Year(now)
	if SearchMonth="" then SearchMonth=Month(now)

	userid = session("ssBctId")

	'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
	'if Not (session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
	'	part_sn = session("ssAdminPsn")
	'end if

	'// 달력 접수
	dim oCalData, lp, weekno
	Set oCalData = new CAgitCalendar

	oCalData.FRectYear = SearchYear
	oCalData.FRectMonth = Num2Str(SearchMonth,2,"0","R")
	oCalData.CalendarList


	'// 휴가일정 접수
	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	oVacation.FRectpart_sn = part_sn
	oVacation.FRectsearchKey = searchKey
	oVacation.FRectsearchString = searchString
	
	oVacation.FPageSize = 1000

	oVacation.FRectYYYY = SearchYear
	oVacation.FRectMM = SearchMonth
	oVacation.GetVacationList



dim year_from, year_to

year_from = Year(now) - 5
year_to = Year(now) + 1

%>
<style type="text/css">
<!--
.calTT { font-family:malgun gothic; color:#606060; background-color:#E0E0E0; }
.calNoB { font-family:Arial; color:#000; font-weight:bold; }
.calNoR { font-family:Arial; color:#FF7065; font-weight:bold; }
.calHoly { font-family:malgun gothic; color:#FFFFFF; background-color:#AABBFF; font-size:11px; padding-left:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calHolyB { font-family:malgun gothic; color:#FFFFFF; background-color:#BABFCF; font-size:11px; padding-left:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calToday { font-family:malgun gothic; color:#000; background-color:#F2F6FF;}
.calFill1 { font-family:malgun gothic; color:#000; background-color:#F6FFE8; font-size:11px; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calFill2 { font-family:malgun gothic; color:#000; background-color:#E8FFF6; font-size:11px; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calFill3 { font-family:malgun gothic; color:#000; background-color:#FFF6E8; font-size:11px; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calFill4 { font-family:malgun gothic; color:#000; background-color:#FFE8F6; font-size:11px; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calNull { background-color:#F0F0F0 }
.calBtn { height:26px; border-radius:6px; border: 1px solid #c0c0c0; font-size:18px; font-family:tahoma; }
-->
</style>

<script type="text/javascript">
<!--
	function OpenDetailView(masteridx, part_sn) {
		var w = window.open("/admin/member/tenbyten/pop_tenbyten_vacation_detail_list_admin.asp?masteridx=" + masteridx + "&part_sn=" + part_sn,"OpenDetailView","width=800,height=600,scrollbars=yes");
		w.focus();
	}
	
	// 페이지 이동
	function goPage(yyyy,mm) {
		document.frm.SearchYear.value=yyyy;
		document.frm.SearchMonth.value=mm;
		document.frm.submit();
	}
//-->
</script>

<!-- 아래는 스크립트보다 위에 있어야 한다. -->
<div ID="viewDIV" STYLE="position:absolute; visibility:hide;"></div>
<!-- 위 한줄은 스크립트보다 위에 있어야 한다. -->

<script type="text/javascript">

////////////////////////////////////////////////////////////////////////////////////
// CONFIGURATION
////////////////////////////////////////////////////////////////////////////////////

var fcolor = "#ffffff";        // Main background color
var textcolor = "#000000";        // Text color
var border_size = "1";                // border size, 1-3
var border_color = "#000000";        // Border color
var width = "300";                // 팝업 박스의 넓이, 100 - 300
var palign = 0;                // 팝업 박스의 위치, 0:center/1:right/2:left

////////////////////////////////////////////////////////////////////////////////////
// END CONFIGURATION
////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////
// END CONFIGURATION
////////////////////////////////////////////////////////////////////////////////////

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

// Non public functions. These are called by other functions etc.

// Simple popup
function viewon(text) {
        var txt = "<TABLE WIDTH="+width+" STYLE='filter:alpha(opacity=100); border:0 ' BORDER=0 CELLPADDING="+border_size+" CELLSPACING=0 BGCOLOR="+border_color+"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR="+fcolor+"><TR><TD><FONT FACE='Arial,Helvetica' COLOR="+textcolor+" style='font-size:12px;'>"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>"
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

// Moves the layer
function mouseMove(e) {
        if (ns4) {x=e.pageX; y=e.pageY;}
        if (ie4) {x=event.x; y=event.y+document.body.scrollTop;}
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

</script>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		부서: <%=printPartOption("part_sn", part_sn)%><br />
		검색:
		<select name="SearchKey" class="select">
			<option value="">::구분::</option>
			<option value="t.userid">아이디</option>
			<option value="t.username">사용자명</option>
		</select>
		<script type="text/javascript">document.frm.SearchKey.value="<%= SearchKey %>";</script>
		<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">

		&nbsp; /&nbsp; 년월:
		<select name="SearchYear" class="select">
		<% for i = year_from to year_to %>
			<option value="<%= i %>" <% if (CInt(SearchYear) = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		</select>
		/
		<select name="SearchMonth" class="select">
		<% for i = 1 to 12 %>
			<option value="<%= i %>" <% if (CInt(SearchMonth) = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
	<td align="center" style="font-size:26px; font-family:malgun gothic;">
	    <input type="button" value="◀" onclick="goPage('<%=SearchYear-1%>','<%=SearchMonth%>')" class="calBtn">
	    <b><%=SearchYear%></b>
	    <input type="button" value="▶" onclick="goPage('<%=SearchYear+1%>','<%=SearchMonth%>')" class="calBtn">
	    &nbsp;/&nbsp;
	    <input type="button" value="◀" onclick="goPage('<%=chkIIF(SearchMonth-1<1,SearchYear-1,SearchYear)%>','<%=chkIIF(SearchMonth-1<1,"12",SearchMonth-1)%>')" class="calBtn">
	    <b><%=SearchMonth%></b>
	    <input type="button" value="▶" onclick="goPage('<%=chkIIF(SearchMonth+1>12,SearchYear+1,SearchYear)%>','<%=chkIIF(SearchMonth+1>12,"1",SearchMonth+1)%>')" class="calBtn">
	</td>
</tr>
<tr>
	<td align="left">※ am : 오전반차, pm : 오후반차</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 예약달력 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" align="center" bgcolor="#FFFFFF" >
	<td width="14%" class="calTT">일</td>
	<td width="14%" class="calTT">월</td>
	<td width="14%" class="calTT">화</td>
	<td width="14%" class="calTT">수</td>
	<td width="14%" class="calTT">목</td>
	<td width="14%" class="calTT">금</td>
	<td width="14%" class="calTT">토</td>
</tr>
<% if oCalData.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#D0D0D0">
<tr height="120" align="center" valign="top" bgcolor="#FFFFFF">
<%
	'// 해당월 1일의 요일
	weekno = DatePart("w", DateSerial(SearchYear,SearchMonth,"01"))

	'// 달력시작 빈칸 표시
	if weekno>1 then
		for lp=1 to (weekno-1)
			Response.Write "<td width='14%' class='calNull'>&nbsp;</td>"
		next
	end if

	for lp=0 to (oCalData.FResultCount-1)
		weekno = DatePart("w", oCalData.FItemList(lp).FDate)
%>
	<td width="14%" <%=chkIIF(oCalData.FItemList(lp).FDate=cstr(date),"class='calToday'","")%>>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr><td align="right" class="<%=chkIIF(weekno=1 or oCalData.FItemList(lp).Fholiday>1,"calNoR","calNoB")%>"><%=lp+1%></td></tr>
		<tr>
			<td>
				<% if Not(oCalData.FItemList(lp).Fholiday_name="" or isNull(oCalData.FItemList(lp).Fholiday_name)) then %><div class="calHoly<%=chkIIF(oCalData.FItemList(lp).Fholiday>1,"","B")%>"><%=oCalData.FItemList(lp).Fholiday_name%></div><% end if %>
				<%
					for j = 0 to oVacation.FResultCount - 1
						'#팝업레이어 표시내용 작성
						strMsg = ""
						strMsg = strMsg & "· 상태 : " & oVacation.FitemList(j).GetStateDivCDStr
						strMsg = strMsg & "<br>· 부서 : " & oVacation.FitemList(j).Fpart_name
						strMsg = strMsg & "<br>· 기간 : " & Left(oVacation.FitemList(j).Fstartday,10) & " - " & Left(oVacation.FitemList(j).Fendday,10)
						if Not(oVacation.FitemList(j).FworkAgent="" or isNull(oVacation.FitemList(j).FworkAgent)) then strMsg = strMsg & "<br>· 업무대행자 : " & oVacation.FitemList(j).FworkAgent
						if Not(oVacation.FitemList(j).FcallNum="" or isNull(oVacation.FitemList(j).FcallNum)) then strMsg = strMsg & "<br>· 비상연락처 : " & oVacation.FitemList(j).FcallNum
	
						if ((oVacation.FitemList(j).GetDay = (lp+1)) and (Not IsNull(oVacation.FitemList(j).Fpart_sn))) then
							if oVacation.FitemList(j).Ftotalday = "0.5" then
								'@반차 (style:1/3)
								Response.Write "<div class='calFill" & chkIIF(oVacation.FitemList(j).Fstatedivcd="R","3","1") & "'>"
							else
								'@종일 (style:2/4)
								Response.Write "<div class='calFill" & chkIIF(oVacation.FitemList(j).Fstatedivcd="R","4","2") & "'>"
							end if

							'// 팀장이상 시스템팀을 제외하고는 소속팀만 상세보기 클릭가능
							if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or session("ssAdminPsn")=7 or cStr(session("ssAdminPsn"))=cStr(part_sn)) then
				%>
						<a href="javascript:OpenDetailView(<%= oVacation.FitemList(j).Fmasteridx %>, <%= oVacation.FitemList(j).Fpart_sn %>)">
						<font style="cursor:hand" onMouseOver="viewon('<%=strMsg%>'); return true;" onMouseOut="viewoff(); return true;"><%= oVacation.FitemList(j).Fusername %></font>
						</a>
				<%			else %>
						<font onMouseOver="viewon('<%=strMsg%>'); return true;" onMouseOut="viewoff(); return true;"><%= oVacation.FitemList(j).Fusername %></font>
				<%
							end if
							
							If oVacation.FitemList(j).Ftotalday = "0.5" Then
								Response.Write "<font color='silver'>[" & oVacation.FitemList(j).Fhalfgubun & "]</font>"
							End IF

							Response.Write "</div>"
						end if
					next
				%>
			</td>
		</tr>
		</table>
	</td>
<%
		'행구분
		if weekno=7 and day(dateAdd("d",1,oCalData.FItemList(lp).FDate))>1 then Response.Write "</tr><tr height='120' align='center' valign='top' bgcolor='#FFFFFF'>"
	next

	'// 달력끝 여백 표시
	if weekno<7 then
		for lp=(weekno+1) to 7
			Response.Write "<td width='14%' class='calNull'>&nbsp;</td>"
		next
	end if
%>
</tr>
</table>
<% end if %>
<!-- 예약달력 끝 -->
<%
	Set oVacation = Nothing
	Set oCalData = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->