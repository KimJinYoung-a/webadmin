<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 통합알림스케줄
' Hieditor : 2022.12.22 한용민 생성
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
dim sIdx, notiType, linkCode, startDate, endDate, reserveTime, pushIsusing, kakaoAlrimIsusing, pushtitle, pushcontents
dim pushurl, templateCode, contents, button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type
dim failed_subject, failed_msg, etc_template_code, member_smsok_checkyn, member_kakaoalrimyn_checkyn, regDate
dim lastUpdate, adminUserid, lastUserid, isusing, time1, time2
dim oSchedule, mode
	sIdx = requestcheckvar(getNumeric(trim(request("sIdx"))),10)
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)

If sIdx <> "" and not(isnull(sIdx)) then
	set oSchedule = new cNotiList
        oSchedule.FrectsIdx = sIdx
		
		oSchedule.fIntegrateNotificationScheduleOne()
        if oSchedule.FResultCount > 0 then
            notiType=oSchedule.FOneItem.FnotiType
            linkCode=oSchedule.FOneItem.flinkCode
            startDate = Left(oSchedule.FOneItem.fstartDate,10)
            endDate = Left(oSchedule.FOneItem.FendDate,10)
            reserveTime=oSchedule.FOneItem.FreserveTime
			time1=left(oSchedule.FOneItem.FreserveTime,2)
			time2=right(oSchedule.FOneItem.FreserveTime,2)
            pushIsusing=oSchedule.FOneItem.fpushIsusing
            kakaoAlrimIsusing=oSchedule.FOneItem.fkakaoAlrimIsusing
            pushtitle=oSchedule.FOneItem.fpushtitle
            if oSchedule.FOneItem.Fpushcontents<>"" then
                pushcontents = replace(oSchedule.FOneItem.Fpushcontents,"\n",vbcrlf)
            end if
            pushurl=oSchedule.FOneItem.Fpushurl
            templateCode=oSchedule.FOneItem.FtemplateCode
            contents=oSchedule.FOneItem.Fcontents
            button_name=oSchedule.FOneItem.Fbutton_name
            button_url_mobile=oSchedule.FOneItem.fbutton_url_mobile
            button_name2=oSchedule.FOneItem.fbutton_name2
            button_url_mobile2=oSchedule.FOneItem.fbutton_url_mobile2
            failed_type=oSchedule.FOneItem.Ffailed_type
            failed_subject=oSchedule.FOneItem.Ffailed_subject
            failed_msg=oSchedule.FOneItem.Ffailed_msg
            etc_template_code=oSchedule.FOneItem.fetc_template_code
            ' 수기템플릿
            if not(etc_template_code="" or isnull(etc_template_code)) then
                templateCode = "etc-9999"
            end if
            member_smsok_checkyn=oSchedule.FOneItem.Fmember_smsok_checkyn
            member_kakaoalrimyn_checkyn=oSchedule.FOneItem.fmember_kakaoalrimyn_checkyn
            regDate=oSchedule.FOneItem.fregDate
            lastUpdate=oSchedule.FOneItem.flastupdate
            adminUserid=oSchedule.FOneItem.fadminUserid
            lastUserid=oSchedule.FOneItem.flastUserid
            isusing=oSchedule.FOneItem.fisusing
        end if
	set oSchedule = Nothing
Else
	Response.write "<script type='text/javascript'>alert('잘못된 접근입니다.');</script>"
	session.codePage = 949
	Response.write "<script type='text/javascript'>self.close();</script>"
	Response.End 
End If 

dim testPkeyVal : testPkeyVal=""
dim testikeyVal : testikeyVal=sIdx+1000000

''해당기능사용안함. 사용할경우 아래 주석처리
testikeyVal = ""
testPkeyVal = ""
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

// 등록된전체기기
function allSubmit(){
    if (frmAct.useridarr.value.length<1){
        alert('일괄발송할 아이디를 입력하세요.');
        frmAct.useridarr.focus();
        return;
    }
    if (frmAct.message.value==''){
        alert('제목을 등록해주세요');
        frmAct.message.focus();
        return;
    }

    if (frmAct.pushcontents.value==''){
        alert('내용을 등록해주세요');
        frmAct.pushcontents.focus();
        return;
    }

    if (confirm('발송하시겠습니까?')){
        frmAct.mode.value="test_allinsert";
        frmAct.submit();
    }
}

// 등록된전체기기
function allSubmit(){
    if (frmAct.useridarr.value.length<1){
        alert('일괄발송할 아이디를 입력하세요.');
        frmAct.useridarr.focus();
        return;
    }

    if (frmAct.pushtitle.value==''){
        alert('제목을 등록해주세요');
        frmAct.pushtitle.focus();
        return;
    }

    if (frmAct.pushcontents.value==''){
        alert('내용을 등록해주세요');
        frmAct.pushcontents.focus();
        return;
    }

    if (confirm('발송하시겠습니까?')){
        frmAct.mode.value="pushTestSend";
        frmAct.submit();
    }
}

</script>

<form name="frmAct" method="post" action="/admin/appmanage/noti/IntegrateNotificationScheduleProcess.asp" style="margin:0px;">
<input type="hidden" name="sIdx" value="<%= sIdx %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시테스트<br>일괄발송 아이디</td>
	<td>
		<input type="text" name="useridarr" value="" size="100" maxlength="150" bgcolor="#FFFFFF">
		<br>예) tozzinet,kobula
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="발송(등록된전체기기)" onClick="allSubmit();" class="button">
	</td>
</tr>
</table>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top: 10px;">
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >스케줄번호</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<b><%=sidx%>번</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시제목</td>
	<td>
		<input type="text" name="param0" value="pushtitle" readonly>
	</td>
	<td>
		<input type="text" name="pushtitle" value="<%= pushtitle %>" size="80" />
	</td>
	<td>필수,alert메세지</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시내용</td>
	<td>
		<input type="text" name="param1" value="pushcontents" readonly>
	</td>
	<td>
		<textarea name="pushcontents" cols=80 rows=8><%= pushcontents %></textarea>
	</td>
	<td>필수,alert메세지</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시경로</td>
	<td>
		<input type="text" name="params" value="url" >
	</td>
	<td><input type="text" name="paramvalue" value="<%=pushurl%>" size="80"></td>
	<td>필수,alert메세지</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">중복방지키</td>
	<td>
		<input type="text" name="params" value="pkey" > <% ''안드로이드 중복 방지용 key 상위  %>
	</td>
	<td><input type="text" name="paramvalue" value="" size="80"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시알림음</td>
	<td>
		<input type="text" name="params" value="sound" >
	</td>
	<td><input type="text" name="paramvalue" value="default" size="80"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시타입</td>
	<td>
		<input type="text" name="params" value="type" >
	</td>
	<td><input type="text" name="paramvalue" value="event" size="80"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">배찌</td>
	<td>
		<input type="text" name="params" value="badge" >
	</td>
	<td><input type="text" name="paramvalue" value="1" size="80"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">클릭체크키</td>
	<td>
		<input type="text" name="params" value="ikey" >
	</td>
	<td><input type="text" name="paramvalue" value="<%=testikeyVal%>" size="80"></td>
	<td></td>
</tr>
</table>
</form>

<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->