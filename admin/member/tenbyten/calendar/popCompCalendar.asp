<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/tenbyten/companyCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

dim i, j, k
dim research, SearchYear, SearchMonth
dim act_fromDate, act_toDate, currDate
dim department_id, myCalOnly

research   = requestcheckvar(request("research"),4)
SearchYear   = requestcheckvar(request("SearchYear"),4)
SearchMonth   = requestcheckvar(request("SearchMonth"),4)
department_id   = requestcheckvar(request("department_id"),32)
myCalOnly = requestcheckvar(request("myCalOnly"),1)

if SearchYear="" then SearchYear=Year(now)
if SearchMonth="" then SearchMonth=Month(now)

act_fromDate = DateSerial(SearchYear, SearchMonth, 1)
act_toDate = DateSerial(SearchYear, SearchMonth+1, 1)


dim oCalData
dim weekno
Set oCalData = new CAgitCalendar

oCalData.FRectYear = SearchYear
oCalData.FRectMonth = Num2Str(SearchMonth,2,"0","R")
oCalData.CalendarList


dim oCompanyCalendar
set oCompanyCalendar = new CCompanyCalendar
	oCompanyCalendar.FPageSize = 1000
	oCompanyCalendar.FCurrPage = 1

	if (myCalOnly = "Y") then
		'// 사번
		oCompanyCalendar.FRectEmpNO = session("ssBctSn")
	end if

	oCompanyCalendar.FRectUseYN = "Y"
'	oCompanyCalendar.FRectStartDate = act_fromDate
'	oCompanyCalendar.FRectEndDate = act_toDate

	oCompanyCalendar.FRectYear = SearchYear
	oCompanyCalendar.FRectMonth = Num2Str(SearchMonth,2,"0","R")

	oCompanyCalendar.FRectDepartmentID = department_id
	oCompanyCalendar.getCompanyCalendarList()

dim year_from, year_to

year_from = Year(now) - 5
year_to = Year(now) + 1

dim strMsg

%>

<script type="text/javascript">

function goPage(yyyy,mm) {
	var frm = document.frm;

	frm.SearchYear.value=yyyy;
	frm.SearchMonth.value=mm;
	frm.submit();
}

</script>

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
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left" height="30">
		연월 :
		<select name="SearchYear" class="select">
		<% for i = year_from to year_to %>
			<option value="<%= i %>" <% if (CInt(SearchYear) = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		</select>
		<select name="SearchMonth" class="select">
		<% for i = 1 to 12 %>
			<option value="<%= i %>" <% if (CInt(SearchMonth) = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		</select>
		&nbsp;
		부서 : <%= drawSelectBoxDepartment("department_id", department_id) %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="30">
		<input type="checkbox" name="myCalOnly" value="Y" <% if (myCalOnly = "Y") then %>checked<% end if %> > 내 일정만 표시
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
</table>
<!-- 액션 끝 -->

<!-- 달력 시작 -->
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
		for i = 1 to (weekno-1)
			Response.Write "<td width='14%' class='calNull'>&nbsp;</td>"
		next
	end if

	for i = 0 to (oCalData.FResultCount-1)
		weekno = DatePart("w", oCalData.FItemList(i).FDate)
		currDate = Left(DateSerial(SearchYear, SearchMonth, (i+1)),10)
%>
	<td width="14%" <%=chkIIF(oCalData.FItemList(i).FDate=cstr(date),"class='calToday'","")%>>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr><td align="right" class="<%=chkIIF(weekno=1 or oCalData.FItemList(i).Fholiday>1,"calNoR","calNoB")%>"><%= (i + 1) %></td></tr>
		<tr>
			<td>
				<% if Not(oCalData.FItemList(i).Fholiday_name="" or isNull(oCalData.FItemList(i).Fholiday_name)) then %><div class="calHoly<%=chkIIF(oCalData.FItemList(i).Fholiday>1,"","B")%>"><%=oCalData.FItemList(i).Fholiday_name%></div><% end if %>
				<%
				for j = 0 to oCompanyCalendar.FResultCount - 1
					if (Left(oCompanyCalendar.FItemList(j).FstartDate,10) <= currDate) and (Left(oCompanyCalendar.FItemList(j).FendDate,10) >= currDate) then
						strMsg = nl2br(oCompanyCalendar.FItemList(j).Fcontents)
						strMsg = Replace(strMsg, Chr(13), "")

						%><font style="cursor:hand;font-family:malgun gothic; color:#000; font-size:10px; " onMouseOver="viewon('<%= strMsg %>'); return true;" onMouseOut="viewoff(); return true;"><%= oCompanyCalendar.FItemList(j).FpName %> : <%= oCompanyCalendar.FItemList(j).Ftitle %></font><br><%

					end if
				next
				%>
			</td>
		</tr>
		</table>
	</td>
<%
		'행구분
		if weekno=7 and day(dateAdd("d",1,oCalData.FItemList(i).FDate))>1 then Response.Write "</tr><tr height='120' align='center' valign='top' bgcolor='#FFFFFF'>"
	next

	'// 달력끝 여백 표시
	if weekno<7 then
		for i = (weekno+1) to 7
			Response.Write "<td width='14%' class='calNull'>&nbsp;</td>"
		next
	end if
%>
</tr>
</table>
<% end if %>
<!-- 달력 끝 -->

<%
Set oCalData = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
