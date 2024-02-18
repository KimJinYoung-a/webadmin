<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description :  사원등록
' History : 2010.12.15 정윤정 수정
'           2013.06.24 허진원 수정; SCM 로그인 비밀번호 확인 추가
'			2016.07.07 한용민 수정
'			2019-01-15 정윤정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim i, menupos, idepartment_id
	menupos = requestCheckVar(Request("menupos"),10)

IF menupos ="" THEN menupos = 1176
%>
<html>
<head>
<title>직원정보 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">

function chk_form(form){
	if(typeof(document.all.sfImg)=="undefined"){
		form.sUImg.value = "";
	}else{
	form.sUImg.value = document.all.sfImg.value;
}

	if(form.sEP.value == "")
	{
		alert("사번 로그인용 비밀번호를 입력해주세요.");
		form.sEP.focus();
		return false;
	}

	if(form.sUN.value == "")
	{
		alert("이름을 입력해주세요.");
		form.sUN.focus();
		return false;
	}
	if(form.department_id.value == "")
	{
		alert("부서를 입력해주십시요.");
		form.department_id.focus();
		return false;
	}

  if(form.hidNm.value==0){
  	alert("이름의 중복확인을 해주세요");
  	return false;
  }
	if(form.hiduserNameEN.value==0){
		alert("영문 이름의 중복확인을 해주세요");
		return false;
	}

	if(form.sJN1.value == "")
	{
		alert("주민등록번호을 입력해주세요.");
		form.sJN1.focus();
		return false;
	}
    /*
	if(form.sJN2.value == "")
	{
		alert("주민등록번호을 입력해주세요.");
		form.sJN2.focus();
		return false;
	}
	*/


	if(form.selJD_y.value == "" )
	{
		alert("입사일을 입력해주세요.");
		form.selJD_y.focus();
		return false;
	}

	if(form.selJD_m.value == "" )
	{
		alert("입사일을 입력해주세요.");
		form.selJD_m.focus();
		return false;
	}

	if(form.selJD_d.value == "" )
	{
		alert("입사일을 입력해주세요.");
		form.selJD_d.focus();
		return false;
	}

	if(form.selPN.value == "")
	{
		alert("부서를 입력해주세요.");
		return false;
	}

	if(form.selPoN.value == "")
	{
		alert("직급을 입력해주세요.");
		return false;
	}

	if(confirm("입사일 "+form.selJD_y.value +"-"+form.selJD_m.value +"-"+form.selJD_d.value +" 로 등록하시겠습니까?(입사일은 수정불가능합니다)"))	{
	return true;
	}
	return false;
}


function jumin_format() {
	var tmp;

	tmp = document.frm_member.sJN1.value + document.frm_member.sJN2.value ;
	if(tmp==""){return;}
	tmp = tmp.replace(/\-/g, "");
/*
	if(!jsChkSocialNum(tmp)){
		alert("유효하지 않은 주민등록번호입니다. 확인 후 다시 등록해주세요");
		document.frm_member.sJN1.value = "";
		document.frm_member.sJN2.value = "";
		document.frm_member.sJN1.focus();
		return;
	}
*/
	if(tmp.substring(6,7)=="1"|| tmp.substring(6,7)=="2"){
	document.frm_member.selBD_y.value = "19"+tmp.substring(0,2);
	}else{
	document.frm_member.selBD_y.value = "20"+tmp.substring(0,2);
	}

	document.frm_member.selBD_m.value = parseInt(tmp.substring(2,4),10);
	document.frm_member.selBD_d.value = parseInt(tmp.substring(4,6),10);

	if(tmp.substring(6,7)=="1"|| tmp.substring(6,7)=="3"){
	document.frm_member.rdoSf[0].checked = true;
	document.frm_member.rdoSf[1].checked = false;
	}else{
	document.frm_member.rdoSf[0].checked = false;
	document.frm_member.rdoSf[1].checked = true;
	}
}

function jsChkSocialNum(sno){
	  var IDAdd = "234567892345";
	  var iDot=0;

	  //숫자확인
	  if(!IsDigit(sno)){
		return false;
	}
	  //숫자가 13자리 인지 확인
	  if(sno.length != 13){
		return false;
	 }

	  if (sno.substring(2,3) > 1) return false;
	  if (sno.substring(4,5) > 3) return false;
	  if (sno.substring(0,2) == '00' && (sno.substring(6,7) != 0 || sno.substring(6,7) != 9 || sno.substring(6,7) != 3 || sno.substring(6,7) !=4)) return false;
	  if (sno.substring(0,2) != '00' && (sno.substring(6,7) > 4 || sno.substring(6,7) == 0)) return false;

	  for(var i=0; i < 13; i ++)
		iDot = iDot + sno.substr(i, 1) * IDAdd.substr(i,1);

	  iDot = 11 - (iDot % 11);

	  if(iDot == 10){
		iDot = 0;
	  } else if (iDot == 11){
		iDot = 1;
	  }

	  if(sno.substr(12,1) == iDot){
		return true;
	  } else {
		return false;
	  }
}

function jsChkfgSocialNum(sno){
	 var sum = 0;
     var odd = 0;

    //숫자확인
	  if(!IsDigit(sno)){
		return false;
	}
	  //숫자가 13자리 인지 확인
	  if(sno.length != 13){
		return false;
	 }

   buf = new Array(13);
   for (i = 0; i < 13; i++) buf[i] = parseInt(sno.charAt(i));

   odd = buf[7]*10 + buf[8];

   if (odd%2 != 0) {
     return false;
   }

   if ((buf[11] != 6)&&(buf[11] != 7)&&(buf[11] != 8)&&(buf[11] != 9)) {
     return false;
   }

   multipliers = [2,3,4,5,6,7,8,9,2,3,4,5];
   for (i = 0, sum = 0; i < 12; i++) sum += (buf[i] *= multipliers[i]);


   sum=11-(sum%11);

   if (sum>=10) sum-=10;

   sum += 2;

   if (sum>=10) sum-=10;

   if ( sum != buf[12]) {
       return false;
   }
   else {
       return true;
   }
}

function fgjumin_format() {
	var tmp;
	var frm = document.frm_member;

	tmp = frm.sJN1.value + frm.sJN2.value ;
	if(tmp=="") {
		return;
	}

	tmp = tmp.replace(/\-/g, "");
/*
	if(!jsChkfgSocialNum(tmp)) {
		alert("유효하지 않은 외국인등록번호입니다. 확인 후 다시 등록해주세요");
		frm.sJN1.value = "";
		frm.sJN2.value = "";
		frm.sJN1.focus();
		return;
	}
*/
	// 871230-6120190
	switch (tmp.substring(6,7) * 1) {
		case 5:
		case 6:
			frm.selBD_y.value = "19"+tmp.substring(0,2);
			break;
		case 7:
		case 8:
			frm.selBD_y.value = "20"+tmp.substring(0,2);
			break;
		case 9:
		case 0:
			frm.selBD_y.value = "18"+tmp.substring(0,2);
			break;
		default:
			// ddd
	}

	frm.selBD_m.value = parseInt(tmp.substring(2,4),10);
	frm.selBD_d.value = parseInt(tmp.substring(4,6),10);

	switch (tmp.substring(6,7) * 1) {
		case 5:
		case 7:
		case 9:
			frm.rdoSf[0].checked = true;
			break;
		case 6:
		case 8:
		case 0:
			frm.rdoSf[1].checked = true;
			break;
		default:
			// ddd
	}
}

function jsChkSN(){

	var iValue;

	for(i=0;i<document.frm_member.rdoSN.length;i++){
		if(document.frm_member.rdoSN[i].checked){
			iValue = document.frm_member.rdoSN[i].value;
		}
	}

	if(iValue == "0"){
		fgjumin_format();
	}else{
		jumin_format();
	}
}

function jsSNReset(){
	document.frm_member.sJN1.value = "";
	document.frm_member.sJN2.value = "";
	document.frm_member.sJN1.focus();
	return;
}

function jsChkName(){

	if (!document.frm_member.sUN.value){
		alert("이름을 등록해주세요");

		document.frm_member.sUN.focus();
		return;
	}

	document.frmNm.target = "ifrmPrc";
	document.frmNm.hidName.value = document.frm_member.sUN.value;
	frmNm.mode.value='C';
	document.frmNm.submit();
}

function jsChkENName(){

	if (!document.frm_member.userNameEN.value){
		alert("이름을 등록해주세요");

		document.frm_member.userNameEN.focus();
		return;
	}

	document.frmNm.target = "ifrmPrc";
	document.frmNm.hiduserNameEN.value = document.frm_member.userNameEN.value;
	frmNm.mode.value='checkNameEN';
	document.frmNm.submit();
}

function jsRegPhoto(){
	var winP= window.open('popAddphoto.asp','imageupload','width=380,height=150');
	winP.focus();
}

function jsFileDel(){
	$("#dFile").html('');
}
</script>
</head>
<body leftmargin="10" topmargin="10">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>사원 정보 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td><font color="red">[기본정보]</font><br>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<form name="frm_member" method="POST" action="/admin/member/tenbyten/member_process.asp" onsubmit="return chk_form(this)" style="margin:0px;" >
		<input type="hidden" name="hidID" value="1">
		<input type="hidden" name="mode" value="A">
		<input type="hidden" name="hidNm" value="0">
		<input type="hidden" name="hiduserNameEN" value="0">
		<input type="hidden" name="sUImg" value="">
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">비밀번호<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<input type="password" name="sEP" class="text" size="20" maxlength="60" autocomplete="off">
				&nbsp;
				(사번 로그인용) <div style="font-size:11px;color:gray;">최소8자 이상, 영문숫자 조합, 같은문자 3번 연속금지</div>
			</td>
			<td rowspan="4" bgcolor="#FFFFFF" width="130"  align="center">
				<%dim vUserImage%>
				<table border="0" cellpadding="0" cellspacing="0" height="132" class="a">
				<tr >
					<td >
						<div id="dFile">
						<img src="<%=vUserImage%>" width="130" alt="원본이미지보기" style="cursor:pointer" onClick="window.open('http://www.10x10.co.kr/common/showimage.asp?img=<%=vUserImage%>', 'imageView', 'width=10,height=10,status=no,resizable=yes,scrollbars=yes');">
						<input type="hidden" name="sfImg" value="<%=vUserImage%>">
						</div>
					</td>
				</tr>
				<tr>
					<td align="center" bgcolor="FFFFFF" valign="bottom"><input type="button" class="button" value="사진등록" onClick="jsRegPhoto();"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">이름<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF"><input type="text" name="sUN" class="text" size="15" maxlength="60" onkeyUp="document.frm_member.hidNm.value =0;" autocomplete="off">
				 <input type="button" name="chkName" value="중복확인" class="button" onClick="jsChkName();"></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">영문이름<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF"><input type="text" name="userNameEN" class="text" size="15" maxlength="32" onkeyUp="document.frm_member.hiduserNameEN.value =0;" autocomplete="off">
				 <input type="button" name="chkENName" value="중복확인" class="button" onClick="jsChkENName();"></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">주민등록번호<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<input type="radio" name="rdoSN" value="1" checked onClick="jsSNReset();"> 내국인
				<input type="radio" name="rdoSN" value="0"  onClick="jsSNReset();"> 외국인
				&nbsp;&nbsp;<input type="text" name="sJN1" class="text" size="6" maxlength="6">-<input type="text" name="sJN2" class="text" size="1" maxlength="1" onFocusOut="jsChkSN();">******
				<br>(앞자리만 입력, 전체번호는 인사기록부로 제출)
				<% if (FALSE) then %><input type="password" name="sJN2" class="text" size="7" maxlength="7" onFocusOut="jsChkSN();"><% end if %>
			</td>
		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">생년월일</td>
    			<td bgcolor="#FFFFFF" colspan="2">
    			    <label><input type="text" name="selBD_y" class="text" value="" size="4" maxlength="4" />년</label>
					<label><input type="text" name="selBD_m" class="text" value="" size="2" maxlength="2" />월</label>
					<label><input type="text" name="selBD_d" class="text" value="" size="2" maxlength="2" />일</label>
    				<label><input type="radio" name="rdoS" value="Y" checked> 양력</label>
					<label><input type="radio" name="rdoS" value="N"> 음력</label>
    			</td>
    		</tr>
    		<tr align="left" height="25">
    		<td bgcolor="<%= adminColor("tabletop") %>">성별</td>
    		<td bgcolor="#FFFFFF" colspan="2"><input type="radio" name="rdoSf" value="M" checked> 남  <input type="radio" name="rdoSf" value="F"> 여</td>
    	</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">핸드폰번호</td>
    			<td bgcolor="#FFFFFF" colspan="2"><input type="text" name="sUC" size="16" class="text" onFocusOut="phone_format(frm_member.sUC)" autocomplete="off"></td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">집전화번호</td>
    			<td bgcolor="#FFFFFF" colspan="2"><input type="text" name="sUP" size="16" class="text"  onFocusOut="phone_format(frm_member.sUP)" autocomplete="off"></td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">우편번호</td>
    			<td bgcolor="#FFFFFF" colspan="2">
    				<input type="text" name="zipcode" size="16" class="text_ro" >
    				<input type="button" class="button" value="검색" onClick="FnFindZipNew('frm_member','B')">
					<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frm_member','B')">
    				<% '<input type="button" class="button" value="검색(구)" onClick="javascript:PopSearchZipcode('frm_member');"> %>
    			</td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">주소</td>
    			<td bgcolor="#FFFFFF" colspan="2">
    				<input type="text" name="zipaddr" size="50" class="text_ro" value="" autocomplete="off">
    				<br><input type="text" name="useraddr" size="60" maxlength="60" class="text" autocomplete="off">
    			</td>
    		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="left" height="25" >
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">입사일(정규직)<font color="red">(*)</font></td><!-- 신규 -->
			<td bgcolor="#FFFFFF">
	    		<select name="selJD_y">
	    			<option value="">-선택-</option>
	<% for i = Year(dateadd("yyyy",1,now()))  to  2001 step -1%>
	    			<option value="<%= i %>" ><%= i %></option>
	<% next %>
	    		</select>년
	    		<select name="selJD_m">
	    		<option value="">-선택-</option>
	<% for i = 1 to 12 %>
	    			<option value="<%= i %>"><%= i %></option>
	<% next %>
	    		</select>월
	    		<select name="selJD_d">
	    		<option value="">-선택-</option>
	<% for i = 1 to 31 %>
	    			<option value="<%= i %>"><%= i %></option>
	<% next %>
	    		</select>일
			</td>
		</tr>
			<%IF menupos = "1176" THEN%>
		<tr align="left" height="25" >
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">실제 입사일 </td><!-- 계약직입사 ==> 항상 사용(실제입사일 입력안한 경우 기본 입사일과 동일) -->
			<td bgcolor="#FFFFFF">
	    		<select name="selRJD_y">
	    			<option value="">-선택-</option>
	<% for i =  Year(dateadd("yyyy",1,now()))  to  2001 step -1%>
	    			<option value="<%= i %>" ><%= i %></option>
	<% next %>
	    		</select>년
	    		<select name="selRJD_m">
	    		<option value="">-선택-</option>
	<% for i = 1 to 12 %>
	    			<option value="<%= i %>"><%= i %></option>
	<% next %>
	    		</select>월
	    		<select name="selRJD_d">
	    		<option value="">-선택-</option>
	<% for i = 1 to 31 %>
	    			<option value="<%= i %>"><%= i %></option>
	<% next %>
	    		</select>일  (계약직->정규직전환시 사용)
			</td>
		</tr>
		<%END IF%>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(사내메일)</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sUM" class="text" size="30" maxlength="80" >
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(개인메일)</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sPM" class="text" size="30" maxlength="80" >
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">전화번호(내선)</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sCUP" class="text" size="16" maxlength="16" >

				&nbsp;&nbsp;
				내선: <input type="text" name="sCE" class="text" size="4" maxlength="16" >
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">070 직통번호</td>
		    	<td bgcolor="#FFFFFF"><input type="text" name="sD070" class="text" size="16" maxlength="16" ></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">텐바이텐사이트 아이디</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sFUI" class="text" size="20" maxlength="32" value=""></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">GSSHOP아이디</td>
			<td bgcolor="#FFFFFF"><input type="text" name="gsshopuserid" class="text" size="20" maxlength="32" value=""></td>
		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">MSN메신저</td>
    			<td bgcolor="#FFFFFF"><input type="text" name="sMM" class="text" size="30" maxlength="80" ></td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">NateOn</td>
    			<td bgcolor="#FFFFFF"><input type="text" name="sNt" class="text" size="30" maxlength="80" ></td>
    		</tr>
    			<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">부서</td>
			<td bgcolor="#FFFFFF">
				<%= drawSelectBoxDepartment("department_id", idepartment_id) %>
			</td>
		</tr>
    		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">어드민권한부서<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<%=printPartOption("selPN", "")%>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">직급<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<select name="selRank"><%=fnRankInfoSelectBox("0")%></select>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">직위<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<%IF menupos = "1176" THEN	'사원관리일떄는 직급전체/ 계약직관리에서는 계약직관련직급%>
				<%=printPositOption("selPoN", "")%>
				<%ELSE%>
				<%=printPositOptionPartTime("selPoN", "")%>
				<%END IF%>
			</td>
		</tr>
		<%IF menupos = 1176 THEN	'사원관리일떄만 직책 보여준다.%>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">직책</td>
			<td bgcolor="#FFFFFF">
				<%=printJobOption("selJN", "")%>
			</td>
		</tr>
		<%END IF%>
	  	<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">담당업무<br>(카테고리)</td>
    			<td bgcolor="#FFFFFF"><% SelectBoxBrandCategory "selC", "" %>
    			</td>
    		</tr>
    		</table>
    	</td>
</tr>
<tr align="center" height="25">
	<td >
		<input type="submit" class="button" value="확인">
		<input type="button" class="button" value="취소" onClick="self.close()">
	</td>
</tr>
</table>
</form>
<form name="frmNm" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
	<input type="hidden" name="hidName" value="">
	<input type="hidden" name="hiduserNameEN" value="">
	<input type="hidden" name="mode" value="">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="ifrmPrc" id="ifrmPrc"  src="about:blank;" style="width:800px; height:300px; frameborder:0;" ></iframe>
<% else %>
	<iframe name="ifrmPrc" id="ifrmPrc"  src="about:blank;" style="width:0px; height:0px; frameborder:0;" ></iframe>
<% end if %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
