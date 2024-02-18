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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim masteridx
	dim i, j
	Dim page, SearchKey, SearchString, part_sn, research, SearchYear, SearchMonth
	dim userid, strMsg, department_id, eventTitle, className, popDescript, isAnnv, eStDate, eEdDate

	page = Request("page")
	research = Request("research")
	masteridx = Request("masteridx")

	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")

	SearchYear = Request("SearchYear")
	SearchMonth = Request("SearchMonth")

	part_sn = Request("part_sn")
	department_id = Request("department_id")

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
	dim lp, weekno

	'// 휴가일정 접수
	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	oVacation.FRectpart_sn = part_sn
	oVacation.Fdepartment_id = department_id
	oVacation.FRectsearchKey = searchKey
	oVacation.FRectsearchString = searchString
	oVacation.FRectYYYY = SearchYear
	oVacation.FRectMM = SearchMonth
	oVacation.GetVacationListSimple

dim year_from, year_to

year_from = Year(now) - 5
year_to = Year(now) + 1

%>
<script src='/js/jquery-1.11.0.min.js'></script>
<link href='/js/fullcalendar/packages/core/main.min.css' rel='stylesheet' />
<link href='/js/fullcalendar/packages/daygrid/main.min.css' rel='stylesheet' />
<link href='/js/fullcalendar/packages/timegrid/main.min.css' rel='stylesheet' />
<link href='/js/fullcalendar/packages/list/main.min.css' rel='stylesheet' />
<script src='/js/fullcalendar/packages/core/main.js'></script>
<script src='/js/fullcalendar/packages/core/ko_euckr.js'></script>
<script src='/js/fullcalendar/packages/daygrid/main.js'></script>
<script src='https://unpkg.com/popper.js/dist/umd/popper.min.js'></script>
<script src='https://unpkg.com/tooltip.js/dist/umd/tooltip.min.js'></script>
<style type="text/css">
.fc-center h2 {font-size:28px;}
.fc-other-month { background-color:#EAEAEA; }
.fc-view-container { background-color:#FDFDFD; }
.fc-day-grid-event {border:0; padding:4px 0 4px 5px; box-shadow: 1px 1px 2px #aaa; margin:1px 3px 2px 2px;}
.fc-title {color:#000; font-size:12px;}

.calHoly { background-color:#AABBFF !important;}
.calHolyB { background-color:#BABFCF !important;}
.calFill1 { background-color:#F6FFE8 !important;}
.calFill2 { background-color:#E8FFF6 !important;}
.calFill3 { background-color:#FFF6E8 !important;}
.calFill4 { background-color:#FFE8F6 !important;}
.calFill5 { background-color:#F8F8F6 !important;}

.tooltip {
  position: absolute;
  z-index: 9999;
  background: #FFFFC1;
  color: black;
  width: 200px;
  border-radius: 3px;
  box-shadow: 0 0 3px rgba(0,0,0,0.5);
  padding: 10px;
  text-align: left;
}
.tooltip .tooltip-arrow {width: 0; height: 0; border-style: solid; position: absolute; margin: 5px;}
.tooltip .tooltip-arrow{border-color: #FFFFC1;}
.tooltip[x-placement^="top"] {margin-bottom: 5px;}
.tooltip[x-placement^="top"] .tooltip-arrow {border-width: 5px 5px 0 5px; border-left-color: transparent; border-right-color: transparent; border-bottom-color: transparent; bottom: -5px; left: calc(50% - 5px); margin-top: 0; margin-bottom: 0;}
.tooltip[x-placement^="bottom"] {margin-top: 5px;}
.tooltip[x-placement^="bottom"] .tooltip-arrow {border-width: 0 5px 5px 5px; border-left-color: transparent; border-right-color: transparent; border-top-color: transparent; top: -5px; left: calc(50% - 5px); margin-top: 0; margin-bottom: 0;}
.tooltip[x-placement^="right"],
.tooltip[x-placement^="right"] .tooltip-arrow {border-width: 5px 5px 5px 0; border-left-color: transparent; border-top-color: transparent; border-bottom-color: transparent; left: -5px; top: calc(50% - 5px); margin-left: 0; margin-right: 0;}
.tooltip[x-placement^="left"] {margin-right: 5px;}
.tooltip[x-placement^="left"] .tooltip-arrow {border-width: 5px 0 5px 5px; border-top-color: transparent; border-right-color: transparent; border-bottom-color: transparent; right: -5px; top: calc(50% - 5px); margin-left: 0; margin-right: 0;}
.tooltip-inner {white-space: pre-wrap;}
</style>
<script type="text/javascript">
	function OpenDetailView(masteridx, part_sn) {
		var w = window.open("/admin/member/tenbyten/pop_tenbyten_vacation_detail_list_admin.asp?masteridx=" + masteridx + "&part_sn=" + part_sn,"OpenDetailView","width=800,height=600,scrollbars=yes");
		w.focus();
	}
	
	// 페이지 이동
	function goPage(dt) {
		var yyyy = parseInt(dt.substr(0,4));
		var mm = parseInt(dt.substr(5,2));
		document.frm.SearchYear.value=yyyy;
		document.frm.SearchMonth.value=mm;
		document.frm.submit();
	}

	// 예약달력 작성
	document.addEventListener('DOMContentLoaded', function() {
		var initialLocaleCode = 'ko';
		var calendarEl = document.getElementById('calendar');

		var calendar = new FullCalendar.Calendar(calendarEl, {
		plugins: [ 'dayGrid'],
		defaultView: 'dayGridMonth',
		locale: initialLocaleCode,
		fixedWeekCount: false,
		businessHours: false,
		header: {
			left: 'today',
			center: 'title',
			right: 'prev,next'
		},
		buttonIcons: false, // show the prev/next text
		contentHeight: "auto",

		defaultDate: '<%=dateSerial(SearchYear,SearchMonth,1)%>',

		eventClick: function(info) {
			var eventObj = info.event;
			if(eventObj.url) {
				var arUrl = eventObj.url.split('|');
				OpenDetailView(arUrl[0],arUrl[1]);
				info.jsEvent.preventDefault();
			}
		},

		eventRender: function(info) {
			var tooltip = new Tooltip(info.el, {
				title: info.event.extendedProps.description,
				placement: 'top',
				trigger: 'hover',
				container: 'body'
			});
		},

		events: [
		<%
			for j = 0 to oVacation.FResultCount - 1
				response.Write "{"

				if Not(oVacation.FItemList(j).Fholiday_name="" or isNull(oVacation.FItemList(j).Fholiday_name)) then
					isAnnv = true
					eventTitle = oVacation.FitemList(j).Fholiday_name
				else
					isAnnv = false
					eventTitle = oVacation.FitemList(j).Fusername
				end if
				If oVacation.FitemList(j).Ftotalday = "0.5" Then
					eventTitle = eventTitle & " [" & oVacation.FitemList(j).Fhalfgubun & "]"
				ElseIf oVacation.FitemList(j).Ftotalday = "0.25" Then
					eventTitle = eventTitle & " [2]"
				End IF

				response.Write "title:'" & eventTitle & "',"

				eStDate = left(oVacation.FitemList(j).Fstartday,10)
				eEdDate = left(oVacation.FitemList(j).Fendday,10)
				if eStDate<eEdDate then eEdDate = left(dateadd("d",1,eEdDate),10)
				response.Write "start:'" & eStDate &  "',"
				response.Write "end:'" & eEdDate &  "',"
				response.Write "allDay:true,"
				if oVacation.FitemList(j).Ftotalday = "0.5" then
					'@반차 (style:1/3)
					className = "calFill" & chkIIF(oVacation.FitemList(j).Fstatedivcd="R","3","1")
				elseIf oVacation.FitemList(j).Ftotalday = "0.25" then
					'@반반차 (style:5)
					className = "calFill5"
				else
					'@종일
					if isAnnv then
						'기념일
						className = "calHoly" & chkIIF(oVacation.FItemList(j).Fholiday>1,"","B")
					else
						'평일 (style:2/4)
						className = "calFill" & chkIIF(oVacation.FitemList(j).Fstatedivcd="R","4","2")
					end if
				end if
				response.Write "className:'" & className & "',"

				if not isAnnv then
					'팝업레이어 표시내용 작성
					popDescript = "· 상태 : " & oVacation.FitemList(j).GetStateDivCDStr & "\n"
					popDescript = popDescript & "· 부서 : " & oVacation.FitemList(j).Fpart_name & "\n"
					popDescript = popDescript & "· 기간 : " & Left(oVacation.FitemList(j).Fstartday,10) & " - " & Left(oVacation.FitemList(j).Fendday,10)
					if Not(oVacation.FitemList(j).FworkAgent="" or isNull(oVacation.FitemList(j).FworkAgent)) then popDescript = popDescript & "\n· 업무대행자 : " & oVacation.FitemList(j).FworkAgent
					if Not(oVacation.FitemList(j).FcallNum="" or isNull(oVacation.FitemList(j).FcallNum)) then popDescript = popDescript & "\n· 비상연락처 : " & oVacation.FitemList(j).FcallNum

					response.Write "description: '" & popDescript & "'"

					'// 팀장이상 시스템팀을 제외하고는 소속팀만 상세보기 클릭가능
					if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or session("ssAdminPsn")=7 or cStr(session("ssAdminPsn"))=cStr(oVacation.FitemList(j).Fpart_sn)) then
						response.Write ",url: '" & oVacation.FitemList(j).Fmasteridx &"|"& oVacation.FitemList(j).Fpart_sn & "'"
					end if

				end if

				response.Write "}"

				if j < oVacation.FResultCount-1 then
				response.Write ","
				end if
			next
		%>
		]
		});

		calendar.render();

		$(".fc-today-button").click(function(){
			goPage('<%=left(date,7)%>');
		});
		$(".fc-prev-button").click(function(){
			goPage('<%=left(DateAdd("m",-1,dateSerial(SearchYear,SearchMonth,1)),7)%>')
		});
		$(".fc-next-button").click(function(){
			goPage('<%=left(DateAdd("m",1,dateSerial(SearchYear,SearchMonth,1)),7)%>')
		});
	});
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
		부서: <%= drawSelectBoxDepartmentALL("department_id", department_id) %><br />
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
	<td align="left">※ am : 오전반차, pm : 오후반차</td>
</tr>
</table>
<!-- 액션 끝 -->

<div class="container">
	<!-- // 예약 달력 //-->
	<div id='calendar'></div>
</div>
<%
	Set oVacation = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->