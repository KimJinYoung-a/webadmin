<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.04.03 한용민 생성
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
Dim ridx, sendmethod, title, contents, state , testsend, repeatlmsyn, olms , adminid, i, targetName, targetkey
dim button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type, failed_subject, failed_msg, template_code, etc_template_code
dim reservedate, date1, time1, time2
	ridx = requestcheckvar(getNumeric(request("ridx")),10)
	repeatlmsyn = requestCheckVar(request("repeatlmsyn"),1)
	menupos = requestcheckvar(getNumeric(request("menupos")),10)

if ridx="" or isnull(ridx) then
    response.write "<script type='text/javascript'>"
    response.write "	alert('잘못된 접근입니다.');"
    response.write "	self.close();"
    response.write "</script>"
    session.codePage = 949
    dbget.close()	:	response.End
end if

set olms = new clms_msg_list
    olms.FRectridx = ridx
    olms.lmsmsg_getrow()

    if olms.ftotalcount > 0 then
        sendmethod = olms.FOneItem.fsendmethod
        title	= olms.FOneItem.ftitle
        'if olms.FOneItem.fcontents<>"" then
        '    contents     = replace(olms.FOneItem.fcontents,"\n",vbcrlf)
        'end if
		contents     = olms.FOneItem.fcontents
        state		= olms.FOneItem.fstate
        targetkey			= olms.FOneItem.ftargetkey
        testsend	= olms.FOneItem.ftestsend
		reservedate			= olms.FOneItem.freservedate
		targetName = olms.FOneItem.ftargetName
		button_name = olms.FOneItem.fbutton_name
		button_url_mobile = olms.FOneItem.fbutton_url_mobile
		button_name2 = olms.FOneItem.fbutton_name2
		button_url_mobile2 = olms.FOneItem.fbutton_url_mobile2
		failed_type = olms.FOneItem.ffailed_type
		failed_subject = olms.FOneItem.ffailed_subject
		failed_msg = olms.FOneItem.ffailed_msg
		template_code = olms.FOneItem.ftemplate_code
		etc_template_code = olms.FOneItem.fetc_template_code
		' 수기템플릿
		if sendmethod="KAKAOALRIM" then
			if not(etc_template_code="" or isnull(etc_template_code)) then
				template_code = "etc-9999"
			end if
		end if
		date1 = Left(reservedate,10)
		time1 = Mid(FormatDateTime(reservedate,4),1,2)
		time2 = Mid(FormatDateTime(reservedate,4),4,2)
    end if
set olms = Nothing

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function lmsSubmit(){
		if (frmAct.useridarr.value.length<1){
			alert('일괄발송할 아이디를 입력하세요.');
			frmAct.useridarr.focus();
			return;
		}

		if (frmAct.sendmethod.value=='LMS'){
			if (frmAct.title.value==''){ 
				alert('제목을 등록해 주세요.');
				frmAct.title.focus();
				return;
			}
		}

		if (frmAct.contents.value==''){
			alert('내용을 등록해주세요');
			frmAct.contents.focus();
			return;
		}

        if (frmAct.sendmethod.value=='KAKAOFRIEND' || frmAct.sendmethod.value=='KAKAOALRIM'){
			//if (frmAct.button_name.value.length<1){
			//	alert('카카오톡 버튼 이름을 입력해 주세요.');
			//	frmAct.button_name.focus();
			//	return;
			//}
			//if (frmAct.button_url_mobile.value.length<1){
			//	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');
			//	frmAct.button_url_mobile.focus();
			//	return;
			//}
			if (frmAct.failed_type.value=='LMS'){
				if (frmAct.failed_subject.value==''){
					alert('카카오톡 실패시 문자제목를 입력해 주세요.');
					frmAct.failed_subject.focus();
					return;
				}
				if (GetByteLength(frmAct.failed_subject.value) > 50){
					alert("카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.");
					frmAct.failed_subject.focus();
					return;
				}
				if (frmAct.failed_msg.value==''){
					alert('카카오톡 실패시 문자내용을 입력해 주세요.');
					frmAct.failed_msg.focus();
					return;
				}
			}
			if (frmAct.sendmethod.value=='KAKAOALRIM'){
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
			}
        }

		if (confirm('테스트 메세지를 발송하시겠습니까?')){
			frmAct.mode.value="test_lmsinsert";
			frmAct.submit();
		}
	}

</script>

<form name="frmAct" id="frmAct" method="post" action="/admin/appmanage/lms/dolmsmsg_proc.asp" style="margin:0px;">
<input type="hidden" name="ridx" value="<%= ridx %>">
<input type="hidden" name="mode" value="test_lmsinsert">
<input type="hidden" name="repeatlmsyn" value="<%= repeatlmsyn %>">
<input type="hidden" name="targetkey" value="<%= targetkey %>">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top: 20px;" >
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b><%= Selectsendmethodname(sendmethod) %> 테스트 발송</b></font><br/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >번호</td>
	<td>
		<%=ridx%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >발송방법</td>
	<td>
		<%= Selectsendmethodname(sendmethod) %>
		<input type="hidden" name="sendmethod" value="<%= sendmethod %>" />
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >발송일</td>
	<td>
		<%= reservedate %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">타게팅여부</td>
	<td>
		<%= targetName %>
	</td>
</tr>

<% if sendmethod="LMS" then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">제목</td>
		<td>
			<input type="text" name="title" value="<%= title %>" size="160" />
		</td>
	</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">내용</td>
	<td>
		<% if sendmethod="KAKAOALRIM" then %>
			템플릿 : <% drawSelectBoxtemplate "template_code", template_code, "", sendmethod %><br><br>
			<% if template_code="etc-9999" then %>
				템플릿코드 : <input type="text" class="text" name="etc_template_code" id="etc_template_code" value="<%= etc_template_code %>" maxlength="32" size="10" /><br><br>
			<% end if %>
		<% end if %>
		<textarea name="contents" cols=100 rows=8><%= contents %></textarea>
	</td>
</tr>

<% if sendmethod<>"LMS" then %>
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
			<% if sendmethod<>"LMS" and failed_type<>"" then %>
				<br><br>실패시문자제목:
				<br><input type="text" class="text" name="failed_subject" value="<%= failed_subject %>" size="55" maxlength=50 />
				<br><br>실패시문자내용:
				<br><textarea name="failed_msg" cols=100 rows=8><%= failed_msg %></textarea>
			<% end if %>
		</td>
	</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">발송아이디</td>
	<td>
		<input type="text" name="useridarr" value="" size="180" maxlength="150" bgcolor="#FFFFFF">
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