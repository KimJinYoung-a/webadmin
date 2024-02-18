<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 통합알림스케줄
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
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
dim nIdx, notiType, linkCode, sendType, userId, device, regDate, lastUpdate, replaceItemId, replaceMileage, isusing
dim oNoti, mode
	nIdx = requestcheckvar(getNumeric(trim(request("nIdx"))),10)
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)

If nIdx = "0" Or nIdx = "" Or isnull(nIdx) Then 
    mode = "NotiInsert"
Else
    mode = "NotiEdit"
End If 

set oNoti = new cNotiList
    oNoti.FrectnIdx = nIdx

    If mode = "NotiEdit" then
		oNoti.fIntegrateNotificationOne()
        if oNoti.FResultCount > 0 then
            notiType=oNoti.FOneItem.fnotiType
            linkCode=oNoti.FOneItem.flinkCode
            sendType=oNoti.FOneItem.fsendType
            userId=oNoti.FOneItem.fuserId
            device=oNoti.FOneItem.fdevice
            regDate=oNoti.FOneItem.fregDate
            lastUpdate=oNoti.FOneItem.flastUpdate
            replaceItemId=oNoti.FOneItem.freplaceItemId
            replaceMileage=oNoti.FOneItem.freplaceMileage
            isusing=oNoti.FOneItem.fisusing
        end if
    end if
set oNoti = Nothing

if notiType="" then notiType="EVENT"
if isusing="" then isusing="Y"

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

//저장
function checkSubmit(){
	var frm=document.inputfrm;

	if (frm.notiType.value==''){ 
		alert('구분을 선택해 주세요.');
		frm.notiType.focus();
		return;
	}
	if (frm.linkCode.value==''){ 
		alert('관련코드를 등록해 주세요.');
		frm.linkCode.focus();
		return;
	}
	if (frm.sendType.value==''){ 
		alert('발송구분을 선택해 주세요.');
		frm.sendType.focus();
		return;
	}
	if (frm.userId.value==''){ 
		alert('고객아이디를 입력해 주세요.');
		frm.userId.focus();
		return;
	}
	if (frm.device.value==''){ 
		alert('신청채널을 선택해 주세요.');
		frm.device.focus();
		return;
	}
	if (frm.isusing.value==''){ 
		alert('사용여부를 선택해 주세요.');
		frm.isusing.focus();
		return;
	}
	//frm.target="_blank";
	frm.submit();
}

</script>

<form name="inputfrm" id="inputfrm" method="post" action="/admin/appmanage/noti/IntegrateNotificationProcess.asp" style="margin:0px;">
<input type="hidden" name="nIdx" value="<%= nIdx %>">
<input type="hidden" name="mode" value="<%= mode %>">
<table width="100%" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>통합알림신청자 등록/수정</b></font>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<% If mode = "NotiEdit" then %>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				신청번호
			</td>
			<td bgcolor="FFFFFF" align="left">
				<%= nIdx %>
			</td>	
		</tr>
		<% end if %>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				구분
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% DrawNotiType "notiType",notiType,"" %>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				관련코드
			</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="linkCode" value="<%= linkCode %>" size=8 maxlength=10 >
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				발송구분
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% DrawsendType "sendType",sendType,"" %>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				고객아이디
			</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="userId" value="<%= userId %>" size=15 maxlength=32 >
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				신청채널
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% DrawIntegrateNotificationDevice "device",device,"" %>
			</td>	
		</tr>
		<tr>
			<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">
				사용여부
			</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxisusingYN "isusing",isusing, "" %>
			</td>	
		</tr>
		<% If mode = "NotiEdit" then %>
			<tr>
				<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">최초등록</td>
				<td bgcolor="FFFFFF" align="left">
					<%= regDate %>
				</td>
			</tr>
			<tr>
				<td width="120" bgcolor="<%= adminColor("tabletop") %>" align="center">마지막수정</td>
				<td bgcolor="FFFFFF" align="left">
					<%= lastUpdate %>
				</td>
			</tr>
		<% end if %>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">
	    <input type="button" value=" 저장 " class="button" onclick="checkSubmit();"/>
	</td>
</tr>
</table>
</form>

<% if (application("Svr_Info")="Dev") then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="100%" height="500"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->