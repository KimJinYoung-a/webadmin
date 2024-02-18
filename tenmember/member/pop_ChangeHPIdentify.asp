<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  본인확인을 사용한 휴대폰번호 변경 팝업
' History : 2011.05.30 허진원 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
dim cMember
dim userid
dim empno, susername, sjuminno, susercell, hp1, hp2, hp3

empno = session("ssBctSn")

'// 직원 기본정보 접수
Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData

	empno   		= cMember.Fempno
	susername      	= cMember.Fusername
	susercell      	= cMember.Fusercell

Set cMember = Nothing

if empno="" or isNull(empno) then
	Call Alert_close("직원 정보가 없습니다.\n관리자에게 문의요망")
	response.End
end if

'//휴대폰 번호 분리
if Not(trim(susercell)="" or isNull(susercell)) then
	susercell = split(susercell,"-")
	if ubound(susercell)>1 then
		hp1 = susercell(0)
		hp2 = susercell(1)
		hp3 = susercell(2)
	end if
end if
%>
<script language='javascript'>
var chkSendAuth = false;

// 인증번호 전송
function chkHPIdentify(){
    var frm = document.frmChkId;

	if(frm.hpNum2.value.length<3){
		alert('핸드폰번호를 입력해주세요');
		frm.hpNum2.focus();
		return ;
	}

	if(frm.hpNum3.value.length<4){
		alert('핸드폰번호를 입력해주세요');
		frm.hpNum3.focus();
		return ;
	}

	if(!chkSendAuth) {
		alert('[인증번호 받기]로 인증번호를 받아주세요.');
		return ;
	}

	frm.target ="hidFrm";
	frm.submit();
}

// SMS 인증번호 발송
function popSMSAuthNo(frm) {
	if(frm.hpNum2.value.length<3){
		alert('핸드폰번호를 입력해주세요');
		frm.hpNum2.focus();
		return ;
	}

	if(frm.hpNum3.value.length<4){
		alert('핸드폰번호를 입력해주세요');
		frm.hpNum3.focus();
		return ;
	}
	frm.chgHp.value = frm.hpNum1.value+"-"+frm.hpNum2.value+"-"+frm.hpNum3.value;

	hidFrm.location.href="/tenmember/member/iframe_adminChgHP_SendSMS.asp?eno="+frm.empNo.value+"&chp="+frm.chgHp.value;
	chkSendAuth = true;
}

// SMS입력 카운터 작동(3분간:180초)
var iSecond=180;
var timerchecker = null;

function startLimitCounter(cflg) {
	if(cflg=="new") {
		if(timerchecker != null) {
			alert("이미 인증번호를 발송하였습니다.\n휴대폰의 SMS를 확인해주세요.");
			return;
		} else if(timerchecker == null) {
			document.getElementById("lySMSTime").style.display="";
		}
		iSecond=180;
	}
    rMinute = parseInt(iSecond / 60);
    rSecond = iSecond % 60;
    if(rSecond<10) {rSecond="0"+rSecond};

    if(iSecond > 0) {
        document.forms[0].sLimitTime.value = rMinute+":"+rSecond;
        iSecond--;
        timerchecker = setTimeout("startLimitCounter()", 1000); // 1초 간격으로 체크
    } else {
        clearTimeout(timerchecker);
        document.forms[0].sLimitTime.value = "0:00";
        timerchecker = null;
        chkSendAuth = false;
        alert("인증번호 입력 시간이 종료되었습니다.\n\nSMS를 받지 못했다면 다시 번호를 받아주세요.");
        document.getElementById("lySMSTime").style.display="none";
    }
}
</script>
<form name="frmChkId" method="post" action="doChangeHPIdentify.asp" onsubmit="return false;">
<input type="hidden" name="empNo" value="<%=empno%>">
<input type="hidden" name="chgHp" value="">
<table width="100%" cellpadding="2" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td colspan="2" bgcolor="#E8F0FF"><b>휴대폰번호 적용 / 본인 확인</b></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사원번호</td>
	<td bgcolor="white"><%=empno%></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">이름</td>
	<td bgcolor="white"><input type="text" name="username" value="<%=susername%>" readonly class="text_ro" size="10"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">휴대폰번호</td>
	<td bgcolor="white">
		<select name="hpNum1" class="select">
			<option value="010" <%=chkIIF(hp1="010","checked","")%>>010</option>
			<option value="011" <%=chkIIF(hp1="011","checked","")%>>011</option>
			<option value="016" <%=chkIIF(hp1="016","checked","")%>>016</option>
			<option value="017" <%=chkIIF(hp1="017","checked","")%>>017</option>
			<option value="018" <%=chkIIF(hp1="018","checked","")%>>018</option>
			<option value="019" <%=chkIIF(hp1="019","checked","")%>>019</option>
		</select>-
		<input name="hpNum2" type="text" class="text" size="4" maxlength="4" value="<%=hp2%>">-
		<input name="hpNum3" type="text" class="text" size="4" maxlength="4" value="<%=hp3%>">
		<input type="button" value='인증번호 받기' onclick="popSMSAuthNo(this.form)" style="padding-top:1px; width:80px; height:20px; border:1px solid #E0E0E0; background-color:#E8F0FF;font-size:11px;">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">인증번호</td>
	<td bgcolor="white">
		<input name="authNo" type="text" class="text" size="6" maxlength="6">
	</td>
</tr>
<!-- // SMS인증번호 입력 유효시간 // -->
<tr id="lySMSTime" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td>
    <td align="left" bgcolor="white">
      	입력 유효시간 : <input type=text name="sLimitTime" value="-:--" readolny style="width:40px; border:1px dotted #E0E0E0; text-align:center;background-color:#F8F8F8;">
    </td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FEF8F8">
		※ 입력한 휴대폰으로 인증문자가 발송되며, 본인확인이 완료되면 [내 정보]의 [휴대폰번호]가 수정됩니다.<br>
		&nbsp;&nbsp;(본인확인은 휴대폰번호 수정 시 최초 1회만 확인)
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="white" align="center">
		<input type="button" class="button" value="본인확인" onclick="chkHPIdentify()">
		&nbsp;&nbsp;<input type="button" class="button" value=" 창닫기 " onclick="self.close();">
	</td>
</tr>
</table>
<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->