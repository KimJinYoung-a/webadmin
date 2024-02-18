<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 통합알림신청자
' Hieditor : 2022.12.26 한용민 생성
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
Dim page, i, oNoti, linkCode, isusing, reload, notiType, sendType, userId
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	page = requestcheckvar(getNumeric(trim(request("page"))),10)
	linkCode = requestcheckvar(getNumeric(trim(request("linkCode"))),10)
	isusing = requestcheckvar(trim(request("isusing")),1)
	reload = requestcheckvar(request("reload"),2)
    notiType=requestcheckvar(trim(request("notiType")),32)
	sendType=requestcheckvar(trim(request("sendType")),16)
	userId=requestcheckvar(trim(request("userId")),32)

if page = "" then page = 1
if reload="" and isusing="" then isusing="Y"

set oNoti = new cNotiList
	oNoti.FPageSize = 50
	oNoti.FCurrPage = page
	oNoti.frectlinkCode = linkCode
	oNoti.frectisusing = isusing
	oNoti.frectnotiType = notiType
	oNoti.frectsendType = sendType
	oNoti.frectuserId = userId
	oNoti.fIntegrateNotificationList()
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	document.frm.target = "";
	document.frm.action = "";
	frm.page.value=page;
	frm.submit();
}

function AddNoti(nIdx){
	var popIntegrateNotification;
	popIntegrateNotification = window.open('/admin/appmanage/noti/popIntegrateNotificationEdit.asp?nIdx='+ nIdx,'popIntegrateNotification','width=1600,height=800,scrollbars=yes,resizable=yes');
	popIntegrateNotification.focus();
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/appmanage/noti/IntegrateNotificationExcel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
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
		* 관련코드 : <input type="text" name="linkCode" value="<%= linkCode %>" size=8 maxlength=10 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		&nbsp;&nbsp;
		* 고객아이디 : <input type="text" name="userId" value="<%= userId %>" size=12 maxlength=32 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
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
		* 발송구분 : 
		<% DrawsendType "sendType",sendType,"" %>
		&nbsp;&nbsp;
		* 사용여부 : 
		<% drawSelectBoxisusingYN "isusing",isusing, "" %>
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
		<font color="red">※발송시간 30분전 발송대상자가 푸시나 알림톡 발송 매뉴에 생성 됩니다.</font> 고객 안내시 참고 부탁 드립니다.
		<br>6개월이 지난 데이터는 다른곳에 백업되고 리스트에서 삭제 됩니다.
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNoti('');">
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oNoti.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oNoti.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=70>신청<br>번호</td>
	<td>구분</td>
	<td>관련코드</td>
	<td>발송구분</td>
	<td>고객아이디</td>
	<td width=50>신청채널</td>
	<td width=80>등록일</td>
	<td width=40>비고</td>
</tr>
<% if oNoti.FresultCount>0 then %>
    <% for i=0 to oNoti.FresultCount-1 %>

    <% if oNoti.FItemList(i).fisusing="N" then %>
		<tr align="center" bgcolor="cccccc" >    
    <% else %>
		<tr align="center" bgcolor="#FFFFFF">
    <% end if %>

    	<td>
    		<%= oNoti.FItemList(i).fnIdx %>
    	</td>
    	<td>
			<%= getNotiType(oNoti.FItemList(i).fnotiType) %>
    	</td>
    	<td>
			<%= oNoti.FItemList(i).flinkCode %>
    	</td>
    	<td>
			<%= getSendType(oNoti.FItemList(i).fsendType) %>
    	</td>
    	<td>
			<%= oNoti.FItemList(i).fuserId %>
    	</td>
    	<td>
			<%= getIntegrateNotificationDevice(oNoti.FItemList(i).fdevice) %>
    	</td>
    	<td>
			<%= left(oNoti.FItemList(i).fregDate,10) %>
			<br><%= mid(oNoti.FItemList(i).fregDate,12,20) %>
    	</td>
    	<td>
			<input type="button" value="수정" onclick="AddNoti('<%= oNoti.FItemList(i).fnIdx %>');" class="button" />
    	</td>
    </tr>
    <% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if oNoti.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= oNoti.StartScrollPage-1 %>')">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oNoti.StartScrollPage to oNoti.StartScrollPage + oNoti.FScrollCount - 1 %>
				<% if (i > oNoti.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oNoti.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oNoti.HasNextScroll then %>
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

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
session.codePage = 949
set oNoti = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->