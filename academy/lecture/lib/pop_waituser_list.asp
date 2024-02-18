<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<!-- 강좌 대기자 -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/common.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
<%
dim lec_idx
lec_idx=RequestCheckvar(request("lec_idx"),10)
dim wlec,w_i,tbcolor
set wlec = new CWaitLecture
wlec.GetWaitList lec_idx
%>

<script language='javascript'>

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,5)=="wfrm_") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function subcheck(){

	for (var i=0;i<document.forms.length;i++){
		sfrm = document.forms[i];
		if (sfrm.name.substr(0,5)=="wfrm_") {
			if (sfrm.cksel.checked){
				realfrm.arridx.value = realfrm.arridx.value + sfrm.widx.value + "," ;
			}
		}
	}
}

function saveopen(){

	var ret = confirm('선택한 사용자의 강좌 등록을 허락합니다.');

	if (ret){
		subcheck();
		realfrm.mode.value="open";
		realfrm.submit();
	}
}


function deluser(){

	var ret = confirm('선택한 사용자를 대기리스트에서 삭제합니다.');

	if (ret){
		subcheck();
		realfrm.mode.value="del";
		realfrm.submit();
	}
}
</script>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr>
		<!--<td bgcolor="#DDDDFF"><input type="checkbox"></td>-->
		<td colspan="9"align="center" bgcolor="#DDDDFF">대기자리스트</td>
	</tr>

	<% for w_i = 1 to wlec.FResultCount %>
	<% if wlec.Flec_idx(w_i) = wlec.Flec_idx(w_i-1) then %>
	<% else %>
	<tr>
		<td colspan="9" bgcolor="#EEEEEE">
			<img src="<%= wlec.FLec_smallimg(w_i) %>" border="0"><%= wlec.FLec_title(w_i) %>(강좌	코드 : <%= wlec.Flec_idx(w_i) %>)</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="20"></td>
		<td width="30" align="center">순위</td>
		<td width="80" align="center">Userid</td>
		<td width="70" align="center">신청인수</td>
		<td width="60" align="center">이름</td>
		<td width="90" align="center">연락처</td>
		<td align="center">이메일</td>
		<td width="120" align="center">신청일</td>
		<td width="120" align="center">접수가능일</td>
	</tr>

	<% end if %>
	<form name="wfrm_<%= w_i %>" method="get" action="">
	<input type="hidden" name="widx" value="<%= wlec.FWaitidx(w_i) %>">
	<%
	 if wlec.FIsusing(w_i)="N" then
		tbcolor="#CCCCCC"
		else
		tbcolor="#FFFFFF"
	 end if
	  %>
	<tr>
		<td bgcolor="<%= tbcolor %>"><input type="checkbox" name="cksel" <% if wlec.Flec_isopen(w_i)="Y" then response.write "checked" %> onClick="AnCheckClick(this);"></td>
		<td bgcolor="<%= tbcolor %>" align="center"><% =wlec.FRegrank(w_i) %></td>
		<td bgcolor="<%= tbcolor %>" align="center"><% =wlec.FUserid(w_i) %>(<%= wlec.FWaitidx(w_i) %>)</td>
		<td bgcolor="<%= tbcolor %>" align="center"><% =wlec.FRegcount(w_i) %></td>
		<td bgcolor="<%= tbcolor %>" align="center"><% =wlec.FUserName(w_i) %></td>
		<td bgcolor="<%= tbcolor %>" align="left"><% =wlec.FPhone(w_i) %></td>
		<td bgcolor="<%= tbcolor %>" align="left"><% =wlec.FEmail(w_i) %></td>
		<td bgcolor="<%= tbcolor %>" align="left"><% =wlec.FRegdate(w_i) %></td>
		<td bgcolor="<%= tbcolor %>" align="left"><% =wlec.FRegEndDay(w_i) %></td>
	</tr>
	</form>
	<% next %>
	<tr>
		<td bgcolor="#FFFFFF" colspan="9" align="center">
			<input type="button" value="적용" onclick="javascript:saveopen();">
			<input type="button" value="삭제" onclick="javascript:deluser();">
			<input type="button" value="취소" onclick=";">
		</td>
	</tr>
</table>
<form name="realfrm" method="post" action="/academy/lecture/lib/doPopWaitUser.asp">
<input type="hidden" name="arridx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="lec_idx" value="<%= lec_idx %>">
</form>

<% set wlec= nothing %>
<br><br>
<script>
function tempsub(frm){

	if (frm.lec_idx.value==""){
		alert('강좌번호는 필수입니다');
		lec_idx.focus();
		return;
	}

	frm.submit();
}
</script>
<div align="left">
<table width="50%" border="1" align="center" class="a" cellpadding="2" cellspacing="1">
	<form name="waittempfrm" method="post" action="doWait.asp">
	<tr>
		<td width="100" align="center"><font color="red">강좌 번호</font<</td>
		<td align="left"><input type="text" name="lec_idx" size="4" maxlength="4" value="<%= lec_idx %>"></td>
	</tr>
	<tr>
		<td width="100" align="center">User Id</td>
		<td align="left"><input type="text" name="userid" size="12" maxlength="32" value=""></td>
	</tr>
	<tr>
		<td width="100" align="center"><font color="red">신청인수</font> </td>
		<td align="left"><input type="text" name="regcount" size="1" maxlength="2" value="1">명</td>
	</tr>
	<tr>
		<td width="100" align="center"><font color="red">이름</font></td>
		<td align="left"><input type="text" name="username" size="6" maxlength="12" value=""></td>
	</tr>
	<tr>
		<td width="100" align="center">연락처</td>
		<td align="left"><input type="text" name="tel01" size="4" maxlength="4" value="">-<input type="text" name="tel02" size="4" maxlength="4" value="">-<input type="text" name="tel03" size="4" maxlength="4" value=""></td>
	</tr>
	<tr>
		<td width="100" align="center">이메일</td>
		<td align="left"><input type="text" name="useremail" size="32" maxlength="64" value=""></td>
	</tr>
	<tr>
		<td  align="center" colspan="2" align="center">
			<input type="button" onclick="javascript:tempsub(this.form);" value="저장">
		</td>
	</tr>
	</form>
</table>
</div>
</body>
</html>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->