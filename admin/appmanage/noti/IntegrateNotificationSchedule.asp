<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 통합알림스케줄
' Hieditor : 2022.12.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/noti/IntegrateNotificationCls.asp" -->
<%
Dim page, i, oSchedule, reservationdate, linkCode, pushIsusing, kakaoAlrimIsusing, isusing, reload, notiType
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	page = requestcheckvar(getNumeric(trim(request("page"))),10)
	reservationdate = requestcheckvar(trim(request("reservationdate")),10)
	linkCode = requestcheckvar(getNumeric(trim(request("linkCode"))),10)
	isusing = requestcheckvar(trim(request("isusing")),1)
	pushIsusing = requestcheckvar(trim(request("pushIsusing")),1)
	kakaoAlrimIsusing = requestcheckvar(trim(request("kakaoAlrimIsusing")),1)
	reload = requestcheckvar(request("reload"),2)
    notiType=requestcheckvar(trim(request("notiType")),32)

if page = "" then page = 1
if reload="" and isusing="" then isusing="Y"

set oSchedule = new cNotiList
	oSchedule.FPageSize = 50
	oSchedule.FCurrPage = page
	oSchedule.frectreservationdate = reservationdate
	oSchedule.frectlinkCode = linkCode
	oSchedule.frectisusing = isusing
	oSchedule.frectpushIsusing = pushIsusing
	oSchedule.frectkakaoAlrimIsusing = kakaoAlrimIsusing
	oSchedule.frectnotiType = notiType
	oSchedule.fIntegrateNotificationScheduleList()
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

function AddNewContents(sIdx){
	var popIntegrateNotificationSchedule;
	popIntegrateNotificationSchedule = window.open('/admin/appmanage/noti/popIntegrateNotificationScheduleEdit.asp?sIdx='+ sIdx,'popIntegrateNotificationSchedule','width=1600,height=800,scrollbars=yes,resizable=yes');
	popIntegrateNotificationSchedule.focus();
}

// 푸시 메시지테스트발송
function pushTest(sIdx){
	var poppush;
	poppush = window.open('/admin/appmanage/noti/popIntegrateNotificationSchedulePushTestSend.asp?sIdx='+ sIdx + '&menupos=<%= menupos %>','poppush','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppush.focus();
}

//메시지테스트발송
function kakaoAlrimTest(sIdx){
	var popkakaoAlrim;
	popkakaoAlrim = window.open('/admin/appmanage/noti/popIntegrateNotificationSchedulekakaoAlrimTestSend.asp?sIdx='+ sIdx + '&menupos=<%= menupos %>','popkakaoAlrim','width=1600,height=800,scrollbars=yes,resizable=yes');
	popkakaoAlrim.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=2>검색<br>조건</td>
	<td align="left">
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
		* 관련코드 : <input type="text" name="linkCode" value="<%= linkCode %>" size=8 maxlength=10 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 구분 : 
		<% DrawNotiType "notiType",notiType,"" %>
		&nbsp;&nbsp;
		* 알림사용여부 : 
		<% drawSelectBoxisusingYN "isusing",isusing, "" %>
		&nbsp;&nbsp;
		* 푸시사용 : 
		<% drawSelectBoxisusingYN "pushIsusing",pushIsusing, "" %>
		&nbsp;&nbsp;
		* 카카오알림톡사용 : 
		<% drawSelectBoxisusingYN "kakaoAlrimIsusing",kakaoAlrimIsusing, "" %>
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
		<font color="red">※발송시간 30분전 발송대상자가 푸시나 알림톡 발송 매뉴에 생성 됩니다.</font> 등록/수정시 중복발송되지 않게 조심해 주세요.
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oSchedule.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oSchedule.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=60>스케줄<br>번호</td>
	<td width=60>구분</td>
	<td width=60>관련코드</td>
	<td width=140>기간</td>
	<td width=50>발송시간</td>
	<td width=40>푸시<br>사용</td>
	<td>푸시제목</td>
	<td width=40>카카오<br>알림톡<br>사용</td>
	<td width=70>카카오<br>알림톡<br>템플릿코드</td>
	<td>알림톡제목</td>
	<td width=40>알림<br>사용<br>여부</td>
	<td width=150>테스트</td>
	<td width=40>비고</td>
</tr>
<% if oSchedule.FresultCount>0 then %>
    <% for i=0 to oSchedule.FresultCount-1 %>

    <% if oSchedule.FItemList(i).fisusing="N" then %>
		<tr align="center" bgcolor="cccccc" >    
    <% else %>
		<tr align="center" bgcolor="#FFFFFF">
    <% end if %>

    	<td>
    		<%= oSchedule.FItemList(i).fsIdx %>
    	</td>
    	<td>
			<%= getNotiType(oSchedule.FItemList(i).fnotiType) %>
    	</td>
    	<td>
			<%= oSchedule.FItemList(i).flinkCode %>
    	</td>
    	<td>
			<%= left(oSchedule.FItemList(i).fstartDate,10) %>~<%= left(oSchedule.FItemList(i).fendDate,10) %>
    	</td>
    	<td>
			<%= oSchedule.FItemList(i).freserveTime %>
    	</td>
    	<td>
			<%= oSchedule.FItemList(i).fpushIsusing %>
    	</td>
    	<td align="left">
    		<%= chrbyte(oSchedule.FItemList(i).fpushtitle,20,"Y") %>
    	</td>
    	<td>
			<%= oSchedule.FItemList(i).fkakaoAlrimIsusing %>
    	</td>
    	<td>
    		<%= oSchedule.FItemList(i).ftemplateCode %>
    	</td>
    	<td align="left">
    		<%= chrbyte(oSchedule.FItemList(i).fcontents,20,"Y") %>
    	</td>
    	<td>
    		<%= oSchedule.FItemList(i).fisusing %>
    	</td>
    	<td>
			<% if oSchedule.FItemList(i).fpushIsusing="Y" then %>
				<input type="button" value="푸시(<%= oSchedule.FItemList(i).fpushTestCount %>건)" onclick="pushTest(<%= oSchedule.FItemList(i).fsIdx %>);" class="button" />
			<% end if %>
			<% if oSchedule.FItemList(i).fkakaoAlrimIsusing="Y" then %>
				<input type="button" value="알림톡(<%= oSchedule.FItemList(i).fkakaoAlrimTestCount %>건)" onclick="kakaoAlrimTest(<%= oSchedule.FItemList(i).fsIdx %>);" class="button" />
			<% end if %>
    	</td>
    	<td>
			<input type="button" value="수정" onclick="AddNewContents('<%= oSchedule.FItemList(i).fsIdx %>');" class="button" />
    	</td>
    </tr>
    <% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if oSchedule.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= oSchedule.StartScrollPage-1 %>')">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oSchedule.StartScrollPage to oSchedule.StartScrollPage + oSchedule.FScrollCount - 1 %>
				<% if (i > oSchedule.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oSchedule.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oSchedule.HasNextScroll then %>
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
set oSchedule = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->