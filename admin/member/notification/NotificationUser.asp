<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 알림관리
' Hieditor : 2023.03.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<%
dim oNoti , i , userId, menupos, notificationType
	userId = requestcheckvar(request("userId"),32)
    menupos = requestCheckvar(getNumeric(Request("menupos")),10)

set oNoti = new CUserNotification
	oNoti.frectuserId = userId
	oNoti.GetNotificationUserList()
%>

<script type="text/javascript">

//삭제
function NotificationUserDel(idx,userId){
	if (confirm('삭제 하시겠습니까?') == true) {
        frmNoti.action='/admin/member/notification/Notificationprocess.asp';
        frmNoti.mode.value="NotificationUserDel";
        frmNoti.idx.value=idx;
        frmNoti.userId.value=userId;
		frmNoti.submit();
	}
}

// 알림추가
function NotificationReg(){
    if (frmNoti.userId.value==''){
        alert('추가할 직원아이디를 입력해주세요');
        frmNoti.userId.focus();
        return;
    }
    
    if (frmNoti.notificationType.value==''){
        alert('추가할 알림을 선택해주세요');
        frmNoti.notificationType.focus();
        return;
    }
    
    frmNoti.action='/admin/member/notification/Notificationprocess.asp';
    frmNoti.mode.value='NotificationReg';
    frmNoti.submit();
}

function jsGoPage(){
	document.frmNoti.submit();
}

</script>

<form name="frmNoti" method="get" action="" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="idx">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        * 직원아이디 : <input type="text" class="text" name="userId" size="17" value="<%=userId%>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="jsGoPage();">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
    </td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
        
	</td>
	<td align="right">
		* 추가할 알림 : <% DrawNotificationType "notificationType" , notificationType, "" %>
		<input type="button" onclick="NotificationReg();" value="추가" class="button">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oNoti.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>직원아이디</td>
    <td>직원명</td>
	<td>알림구분</td>
	<td>알림명</td>	
	<td>비고</td>
</tr>
<% if oNoti.ftotalcount > 0 then %>
	
<% for i=0 to oNoti.ftotalcount - 1 %>

<tr align="center" bgcolor="#FFFFFF" >
    <td><%= oNoti.FItemList(i).fuserid %></td>
    <td><%= oNoti.FItemList(i).fusername %></td>
	<td>
		<%= oNoti.FItemList(i).fnotificationType %>
	</td>
	<td>
		<%= oNoti.FItemList(i).fnotificationTypeName %>
	</td>	
	<td>
		<input type="button" onclick="NotificationUserDel('<%= oNoti.FItemList(i).fidx %>','<%= oNoti.FItemList(i).fuserid %>');" value="삭제" class="button">
	</td>	
</tr>   
<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oNoti = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->