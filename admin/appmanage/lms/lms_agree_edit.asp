<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS/친구톡/알림톡 수신동의 관리
' Hieditor : 2021.08.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->

<%
dim adminuserid, i, menupos, olms
dim userid, regdate, lastupdate, reguserid, lastuserid, kakaoalrimyn
	userid = requestcheckvar(request("userid"),32)
	menupos = requestcheckvar(getNumeric(request("menupos")),10)

adminuserid=session("ssBctId")

set olms = new clms_msg_list
	olms.frectuserid = userid
	
	if userid <> "" then
		olms.flms_agree_one()
		
		if olms.ftotalcount > 0 then
			userid = olms.foneitem.fuserid
			regdate = olms.foneitem.fregdate
			lastupdate = olms.foneitem.flastupdate
			reguserid = olms.foneitem.freguserid
			lastuserid = olms.foneitem.flastuserid
            kakaoalrimyn = olms.foneitem.fkakaoalrimyn
		end if
	end if

if kakaoalrimyn = "" then kakaoalrimyn = "Y"
%>

<script type="text/javascript">

function reg_lms_agree(){
	if (frm.userid.value==""){
		alert('아이디를 입력해 주세요.');
		frm.userid.focus();
		return;
	}
	if (frm.kakaoalrimyn.value==""){
		alert('알림톡 수신여부를 선택해주세요.');
		frm.kakaoalrimyn.focus();
		return;
	}
	
	frm.action="/admin/appmanage/lms/lms_agree_process.asp";
	frm.mode.value="lms_agree_edit";
	frm.submit();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<input type="hidden" name="mode">
<tr bgcolor="#FFFFFF">
	<td align="center"><b>아이디</b><br></td>
	<td>
        <% if userid<>"" then %>
		    <%= userid %>
		    <input type="hidden" name="userid" value="<%=userid%>">
        <% else %>
            <input type="text" name="userid" value="<%=userid%>" size=12 maxlength=32>
        <% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>알림톡수신여부</b><br></td>
	<td>
        <% drawSelectBoxisusingYN "kakaoalrimyn", kakaoalrimyn, "" %>
	</td>
</tr>
<% if userid<>"" then %>
    <tr bgcolor="#FFFFFF">
        <td align="center"><b>최초등록</b><br></td>
        <td>
            <% if regdate<>"" then %>
                <%= regdate %>
            <% end if %>

            <% if reguserid<>"" then %>
                <Br>(<%= reguserid %>)
            <% end if %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center"><b>최종수정</b><br></td>
        <td>
            <% if lastupdate<>"" then %>
                <%= lastupdate %>
            <% end if %>

            <% if lastuserid<>"" then %>
                <Br>(<%= lastuserid %>)
            <% end if %>
        </td>
    </tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
		<input type="button" value="저장" onclick="reg_lms_agree();" class="button">
	</td>
</tr>
</table>
</form>

<%
set olms = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
