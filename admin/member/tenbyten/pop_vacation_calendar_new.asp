<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
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

	'// 휴가일정 접수
	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	oVacation.FRectpart_sn = part_sn
	oVacation.FRectsearchKey = searchKey
	oVacation.FRectsearchString = searchString
	oVacation.FRectYYYY = SearchYear
	oVacation.FRectMM = SearchMonth
	oVacation.GetVacationListSimple

dim year_from, year_to

year_from = Year(now) - 5
year_to = Year(now) + 1

%>
<style type="text/css">
.fc-center h2 {font-size:28px;}
.calTT { color:#606060; background-color:#E0E0E0; }
.calNoB { color:#000; font-weight:bold; }
.calNoR { color:#FF7065; font-weight:bold; }
.calHoly { color:#FFFFFF; background-color:#AABBFF !important; font-size:11px; padding-left:5px; margin:2px;}
.calHolyB { color:#FFFFFF; background-color:#BABFCF !important; font-size:11px; padding-left:5px; margin:2px;}
.calFill1 { background-color:#B6BBA8 !important; border-color:#B6BBA8 !important;}
.calFill2 { background-color:#A8BBB6 !important;}
.calFill3 { background-color:#BBB6A8 !important;}
.calFill4 { background-color:#BBA8B6 !important;}
.calNull { background-color:#B0B0B0 !important; }
.calBtn { height:26px; border-radius:6px; border: 1px solid #c0c0c0; font-size:18px; font-family:tahoma; }
.fc-other-month { background-color:#F8F8F8; }
.fc-view-container { background-color:#FDFDFD; }
</style>
<link rel='stylesheet' href='/js/fullcalendar/fullcalendar.min.css' />
<script src='/js/jquery-1.7.1.min.js'></script>
<script src='/js/fullcalendar/moment.min.js'></script>
<script src='/js/fullcalendar/fullcalendar.min.js'></script>
<script src='/js/fullcalendar/locale/ko_euckr.js'></script>
<script type="text/javascript">
function OpenDetailView(masteridx, part_sn) {
	var w = window.open("/admin/member/tenbyten/pop_tenbyten_vacation_detail_list_admin.asp?masteridx=" + masteridx + "&part_sn=" + part_sn,"OpenDetailView","width=800,height=600,scrollbars=yes");
	w.focus();
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
	<td align="left">※ am : 오전반차, pm : 오후반차</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 예약달력 시작 -->
<div id='calendar' style="padding-right:10px;"></div>
<script type="text/javascript">
$(function() {
	$('#calendar').fullCalendar({
		defaultDate: '<%=DateSerial(SearchYear,SearchMonth,1)%>',
		locale: 'ko',
		header: {
			left: 'today',
			center: 'title',
			right: 'prev,next'
		},
		buttonIcons: false,
		events:[
		<%
			for j = 0 to oVacation.FResultCount - 1
				response.Write "{"
				response.Write "title:'" & oVacation.FitemList(j).Fusername & "',"
				response.Write "start:'" & left(oVacation.FitemList(j).Fstartday,10) &  "',"
				response.Write "end:'" & left(oVacation.FitemList(j).Fendday,10) &  "',"
				if oVacation.FitemList(j).Fholiday>0 then 
					response.Write "className:'calHoly" & chkIIF(oVacation.FitemList(j).Fholiday>1,"","B") & "'"
				else
					response.Write "className:'calFill" & chkIIF(oVacation.FitemList(j).Fstatedivcd="R","3","1") & "'"
				end if
				response.Write "}"

				if j < oVacation.FResultCount-1 then
				response.Write ","
				end if
			next
		%>
		],
		eventClick: function() {
			alert("aa");
		}
	});
});
</script>
<!-- 예약달력 끝 -->
<%
	Set oVacation = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->