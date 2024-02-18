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
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
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

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function lmsSubmit(){
		if (frmAct.useridarr.value.length<1){
			alert('일괄발송할 아이디를 입력하세요.');
			frmAct.useridarr.focus();
			return;
		}

		if (frmAct.contents.value==''){
			alert('카카오 알림톡 내용을 등록해주세요');
			frmAct.contents.focus();
			return;
		}

        if (frmAct.template_code.value==''){
            alert('카카오톡 알림톡 템플릿코드가 지정되어 있지 않습니다.');
            frmAct.template_code.focus();
            return;
        }
        // 알림톡 수기템플릿
        if ($('#frmAct select[name="template_code"] option:selected').val()=='etc-9999'){
            if(frmAct.etc_template_code.value==''){ 
                alert("수기템플릿코드를 입력해 주세요.");
                frmAct.etc_template_code.focus();
                return;
            }
        }

		if (confirm('테스트 메세지를 발송하시겠습니까?')){
			frmAct.mode.value="kakaoAlrimTestSend";
			frmAct.submit();
		}
	}

</script>

<form name="frmAct" id="frmAct" method="post" action="/admin/appmanage/noti/IntegrateNotificationScheduleProcess.asp" style="margin:0px;">
<input type="hidden" name="sIdx" value="<%= sIdx %>">
<input type="hidden" name="mode" value="">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top: 20px;" >
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>카카오알림톡 테스트 발송</b></font><br/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >스케줄번호</td>
	<td>
		<%= sIdx %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >발송방법</td>
	<td>
        카카오알림톡
        <input type="hidden" name="sendmethod" value="KAKAOALRIM">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">내용</td>
	<td>
        템플릿 : <% drawSelectBoxtemplate "template_code", templateCode, "", "KAKAOALRIM" %><br><br>
        <% if templateCode="etc-9999" then %>
            템플릿코드 : <input type="text" class="text" name="etc_template_code" id="etc_template_code" value="<%= etc_template_code %>" maxlength="32" size="10" /><br><br>
        <% end if %>
		<textarea name="contents" cols=100 rows=8><%= contents %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>" align="center">KAKAO</td>
    <td>
        버튼이름1 : 
        <Br><input type="text" class="text" name="button_name" value="<%= button_name %>" size="64" maxlength=64 />
        <Br>
        버튼모바일주소1 : 
        <Br><input type="text" class="text" name="button_url_mobile" value="<%= button_url_mobile %>" size="120" maxlength=256 />
        <Br><Br>
        버튼이름2 : 
        <Br><input type="text" class="text" name="button_name2" value="<%= button_name2 %>" size="64" maxlength=64 />
        <Br>
        버튼모바일주소2 : 
        <Br><input type="text" class="text" name="button_url_mobile2" value="<%= button_url_mobile2 %>" size="120" maxlength=256 />
        <Br><Br>
        실패시문자발송여부 : <% Drawfailed_type "failed_type", failed_type, "" %>
        <% if failed_type<>"" then %>
            <br><br>실패시문자제목:
            <br><input type="text" class="text" name="failed_subject" value="<%= failed_subject %>" size="55" maxlength=50 />
            <br><br>실패시문자내용:
            <br><textarea name="failed_msg" cols=100 rows=8><%= failed_msg %></textarea>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">발송아이디</td>
	<td>
		<input type="text" name="useridarr" value="" size="100" maxlength="150" bgcolor="#FFFFFF">
		<br>예) tozzinet,kobula
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
		<input type="button" value="테스트메세지발송" onClick="lmsSubmit();" class="button">
	</td>
</tr>
</table>
</form>

<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->