<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.03.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
Dim page, i, reload, isusing, state, reservationdate, title, targetkey, repeatlmsyn, ridx, olms, sendmethod
	sendmethod = requestcheckvar(request("sendmethod"),16)
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
	reload = requestcheckvar(request("reload"),2)
	isusing = requestcheckvar(request("isusing"),1)
	state = requestcheckvar(request("state"),1)
	reservationdate = requestcheckvar(request("reservationdate"),10)
	title = requestcheckvar(request("title"),300)
	targetkey = requestcheckvar(request("targetkey"),10)
	repeatlmsyn = requestcheckvar(request("repeatlmsyn"),1)
	ridx = requestcheckvar(getNumeric(request("ridx")),10)

if page = "" then page = 1
'if repeatlmsyn="" then repeatlmsyn="N"
if reload="" and isusing="" then isusing="Y" ''사용중 기본
'if sendmethod="" then sendmethod="LMS"

set olms = new clms_msg_list
	olms.FPageSize = 50
	olms.FCurrPage = page
	olms.Frectreservedate = reservationdate
	olms.Frectstate = state
	olms.Frectisusing = isusing
	olms.Frecttitle = title
	olms.Frecttargetkey = targetkey
	olms.Frectrepeatlmsyn = repeatlmsyn
	olms.frectridx = ridx
	olms.frectsendmethod=sendmethod
	olms.flmsmsglist()

dim Svr_Info : Svr_Info=""
IF application("Svr_Info")="Dev" THEN
	Svr_Info = "Dev"
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

//예약등록
function AddNewContents(ridx){
	var poplmsin;
	poplmsin = window.open('/admin/appmanage/lms/poplmsmsg_edit.asp?ridx='+ ridx,'poplmsin','width=1600,height=800,scrollbars=yes,resizable=yes');
	poplmsin.focus();
}

//메시지테스트발송
function example_msg(ridx, repeatlmsyn){
	var poplmsexam;
	poplmsexam = window.open('/admin/appmanage/lms/poplmsmsg_example.asp?ridx='+ ridx + '&repeatlmsyn=' + repeatlmsyn,'poplmsexam','width=1600,height=800,scrollbars=yes,resizable=yes');
	poplmsexam.focus();
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function poplmsReport(ridx){
	var poplmsReport;
	poplmsReport = window.open('/admin/appmanage/lms/poplmsmsg_report_realtime.asp?ridx='+ ridx,'lmsReport','width=1600,height=800,scrollbars=yes,resizable=yes');
	poplmsReport.focus();
}

function poplmsReportStatisticsDB(ridx){
	var poplmsReportStatisticsDB;
	poplmsReportStatisticsDB = window.open('/admin/appmanage/lms/poplmsmsgReportStatisticsDB.asp?ridx='+ ridx,'lmsReportStatisticsDB','width=1600,height=800,scrollbars=yes,resizable=yes');
	poplmsReportStatisticsDB.focus();
}

// 타켓쿼리관리
function targetqueryreg(){
	var poplmstarget<%= Svr_Info %>;
	poplmstarget<%= Svr_Info %> = window.open('/admin/appmanage/lms/lmstarget.asp','poplmstarget<%= Svr_Info %>','width=1600,height=800,scrollbars=yes,resizable=yes');
	poplmstarget<%= Svr_Info %>.focus();
}

function templatereg(){
	var poplmstemplatereg<%= Svr_Info %>;
	poplmstemplatereg<%= Svr_Info %> = window.open('/admin/appmanage/lms/lmstemplatereg.asp','poplmstemplatereg<%= Svr_Info %>','width=1600,height=800,scrollbars=yes,resizable=yes');
	poplmstemplatereg<%= Svr_Info %>.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 발송방법 : <% Drawsendmethod "sendmethod",sendmethod, " onchange='frmsubmit("""");'","Y" %>
		&nbsp;&nbsp;
		* 발송타켓 : <% drawSelectBoxlmsTarget "targetkey", targetkey, " onchange='frmsubmit("""");'", repeatlmsyn %>
		&nbsp;&nbsp;
		* 제목 : <input type="text" name="title" value="<%= title %>" size=25 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		&nbsp;&nbsp;
		* 번호 : <input type="text" name="ridx" value="<%= ridx %>" size=8 maxlength=10 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		<br><br>
		* 발송일 : 
		<input type="text" id="termSdt" name="reservationdate" size="7" maxlength=10 value="<%= reservationdate %>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;" />
		<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "termSdt", trigger    : "ChkStart_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					//CAL_End.args.min = date;
					//CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(reservationdate<>"",", max: " & replace(reservationdate,"-",""),"")%>
			});
		</script>
		&nbsp;&nbsp;
		* 사용여부 : 
		<% drawSelectBoxisusingYN "isusing",isusing, " onchange='frmsubmit("""");'" %>
		&nbsp;&nbsp;
		* 상태
		<% Drawlmsstatename "state" , state , " onchange='frmsubmit("""");'", "" %>
        &nbsp;&nbsp;
        * 반복발송 : <% Drawrepeatgubun "repeatlmsyn", repeatlmsyn, " onchange='frmsubmit("""");'", "Y" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	※ 알림톡의 경우 문자 수신여부와 관계없이 전부(Y,N) 발송됩니다. 알림톡 수신거부는 [ON]사이트관리>>알림톡 수신동의 관리 에서 가능합니다.
	<br>LMS 와 친구톡의 경우 문자 수신여부 Y 만 발송 됩니다.
	<br>CSV타켓(휴대폰번호) 타켓일경우 수신여부 체크가 불가하니 전부 발송 됩니다.
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">
		<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="템플릿관리(관리자)" onclick="templatereg();">
			<input type="button" class="button" value="타켓쿼리관리(관리자)" onclick="targetqueryreg();">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= olms.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olms.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=70>번호</td>
	<td width=50>발송방법</td>
	<td width=70>발송일</td>
	<td>메세지</td>
	<td width=50>상태</td>	
	<td width=30>사용<br>여부</td>
	<td width=50>타겟상태</td>
	<td>타겟</td>
	<td width=70>타겟수량</td>
    <td width=50>반복발송</td>
	<td width=70>최초등록</td>
	<td width=70>마지막수정</td>
	<td width=120>비고</td>
</tr>
<% if olms.FresultCount>0 then %>
    <% for i=0 to olms.FresultCount-1 %>

    <% if (olms.FItemList(i).fisusing="N") then %>
		<tr align="center" bgcolor="cccccc" >    
    <% else %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background="#FFFFFF";>
    <% end if %>

    	<td>
    		<%= olms.FItemList(i).Fridx %>
    	</td>
    	<td>
    		<%= Selectsendmethodname(olms.FItemList(i).fsendmethod) %>
    	</td>
    	<td>
			<%= left(olms.FItemList(i).Freservedate,10) %>
			<br><%= mid(olms.FItemList(i).Freservedate,12,11) %>
    	</td>
    	<td align="left">
			<% if olms.FItemList(i).fsendmethod="KAKAOFRIEND" or olms.FItemList(i).fsendmethod="KAKAOALRIM" then %>
				<%= chrbyte(olms.FItemList(i).fcontents,50,"Y") %>
			<% else %>
				<%= chrbyte(olms.FItemList(i).ftitle,50,"Y") %>
			<% end if %>	
    	</td>
    	<td>
    		<%= lmsmsgstate(olms.FItemList(i).fstate)%>
    	</td>
    	<td>
    		<%= olms.FItemList(i).fisusing %>
    	</td>
    	<td><%=olms.FItemList(i).getTargetStateName%></td>
    	<td>
    	    <%= olms.FItemList(i).ftargetName %>
    	</td>
    	<td><%=FormatNumber(olms.FItemList(i).fTargetCnt,0)%></td>
    	<td>
    		<%= Selectlmsgubunname(olms.FItemList(i).frepeatlmsyn) %>
    	</td>
		<td>
			<%= left(olms.FItemList(i).fregdate,10) %>
			<br><%= mid(olms.FItemList(i).fregdate,12,11) %>
			<br><%=olms.FItemList(i).fregadminid%>
		</td>
		<td>
			<%= left(olms.FItemList(i).flastupdate,10) %>
			<br><%= mid(olms.FItemList(i).flastupdate,12,11) %>
			<br><%=olms.FItemList(i).flastadminid%>
		</td>
    	<td>
			<input type="button" value="수정" onclick="AddNewContents('<%= olms.FItemList(i).Fridx %>');" class="button" />
    		<input type="button" value="테스트(<%= olms.FItemList(i).ftestsend %>건)" onclick="example_msg(<%= olms.FItemList(i).Fridx %>,'N');" class="button" />
			
			<%
			' CSV타켓(휴대폰번호):1
			'if olms.FItemList(i).ftargetkey<>"1" then
			%>
				<%
				' 발송완료
				if olms.FItemList(i).fstate="9" then
				%>
					<% ' 통계(실시간)도 사용은 가능하게 코딩은 해놓음. 부하 문제로 숨겨놓음. %>
					<!--<br><input type="button" value="통계(실시간)" onClick="poplmsReport('<%'= olms.FItemList(i).Fridx %>');" class="button" />-->
					<input type="button" value="통계" onClick="poplmsReportStatisticsDB('<%= olms.FItemList(i).Fridx %>');" class="button" />
				<% end if %>
			<% 'end if %>
    	</td>
    </tr>
    <% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if olms.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= olms.StartScrollPage-1 %>')">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + olms.StartScrollPage to olms.StartScrollPage + olms.FScrollCount - 1 %>
				<% if (i > olms.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(olms.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if olms.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
session.codePage = 949
set olms = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->