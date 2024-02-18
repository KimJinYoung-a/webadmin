<%@ language=vbscript %>
<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lecture_waitingusercls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim Cookieuserid, Cookieusername, Cookieuseremail

Cookieuserid = request.cookies("uinfo")("userid")
Cookieusername = request.cookies("uinfo")("username")
Cookieuseremail = request.cookies("uinfo")("useremail")


dim idx,lec_idx,owaiting
lec_idx=RequestCheckvar(request("lec_idx"),10)
idx=RequestCheckvar(request("idx"),10)

if idx="" then idx="0"

set owaiting = new CLecWaitUser
owaiting.FRectIdx = idx
owaiting.GetOneWaitUser


dim RegMode
if (owaiting.FResultCount>0) then
	RegMode = "edit"
else
	RegMode = "add"
	if lec_idx="" then
		response.write "<script>alert('목록에서 강좌코드를 검색 후 등록이 가능합니다.');</script>"
		response.write "<script>self.close();</script>"
		response.end
	end if
end if


dim olecture
set olecture = new CLecture
olecture.FRectIdx = lec_idx

if (RegMode="add") and (lec_idx<>"") then
	olecture.GetOneLecture
end if


%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<title>강좌 대기신청</title>
<link href="<%=wwwFingers%>/lib/css/2009fingers.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/js/common.js"></script>
<script language="javascript">
<!--
self.resizeTo(637,750);

function popUserInfo(comp){
	var uid= comp.value;

	if (uid.length<1){
		alert('아이디를 입력하세요');
		comp.focus();
		return;
	}
	var popwin=window.open('/common/popuserinfo.asp?uid=' + uid,'popuserinfo','width=100,height=100.scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReActUser(username,userphone,userhp,useremail){
	waitfrm.username.value=username;
	waitfrm.useremail.value=useremail;

	userhp = userhp.split('-');
	waitfrm.tel01.value = userhp[0];
	waitfrm.tel02.value = userhp[1];
	waitfrm.tel03.value = userhp[2];
}

function frmsub(frm){


	if (frm.lec_idx.value.length<1){
		alert('강좌번호를 입력해주십시요..');
		frm.lec_idx.focus();
		return;
	}

	if (!(IsDigit(frm.lec_idx.value))){
		alert('강좌번호는는 숫자만 가능합니다.');
		frm.lec_idx.focus();
		return;
	}


	if (frm.userid.value.length<1){
		alert('고객 아이디를 입력해주십시요. 필수 사항입니다.');
		frm.userid.focus();
		return;
	}

	if (frm.regcount.value.length<1){
		alert('대기신청수를 입력해주십시요..');
		frm.regcount.focus();
		return;
	}

	if (!(IsDigit(frm.regcount.value))){
		alert('대기신청수는 숫자만 가능합니다.');
		frm.regcount.focus();
		return;
	}

	if ((frm.regcount.value>2)){
		alert('대기신청수는 최대 2명 까지 가능합니다.');
		frm.regcount.focus();
		return;
	}

	if (frm.username.value.length<1){
		alert('성함을 입력해주십시요..');
		frm.username.focus();
		return;
	}

	if (frm.tel01.value.length<1){
		alert('연락처를 입력해주십시요.');
		frm.tel01.focus();
		return;
	}

	if (frm.tel02.value.length<1){
		alert('연락처를 입력해주십시요.');
		frm.tel02.focus();
		return;
	}

	if (frm.tel03.value.length<1){
		alert('연락처를 입력해주십시요.');
		frm.tel03.focus();
		return;
	}

	if (!(IsDigit(frm.tel01.value.length))){
		alert('연락처를 입력해주십시요.');
		frm.tel01.focus();
		return;
	}

	if (!(IsDigit(frm.tel02.value.length))){
		alert('연락처를 입력해주십시요.');
		frm.tel02.focus();
		return;
	}

	if (!(IsDigit(frm.tel03.value.length))){
		alert('연락처를 입력해주십시요.');
		frm.tel03.focus();
		return;
	}

	if (frm.useremail.value.length<1){
		alert('이메일을 입력해주십시요.');
		frm.useremail.focus();
		return;
	}

	<% if RegMode="edit" then %>
	if ((frm.currstate.value=="3")&&(frm.regendday.value.length!=19)){
		alert('결제 마감일을 입력하세요.');
		frm.regendday.focus();
		return;
	}

	if (confirm('대기자 내역을 수정 하시겠습니까?')){
			frm.submit();
	}
	<% else %>
	if (!frm.lecOption.value){
		alert('강좌일을 선택해주십시요.');
		frm.lecOption.focus();
		return;
	}

	if (confirm('대기자를 신규 등록 하시겠습니까?')){
			frm.submit();
	}
	<% end if %>
}

function chgOption() {
	var msgSold, remainNo = document.waitfrm.lecOption.options[document.waitfrm.lecOption.selectedIndex].id;
	if(!remainNo) remainNo=0;
	msgSold = "<b>" + remainNo + " 명</b>"
	document.all.msgSold.innerHTML=msgSold;
}
-->
</script>
</head>
<body style="margin-left: 0px;margin-top: 0px;">
<table width="600" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td><img src="<%=imgFingers%>/academy2009/lecture/waitapply_title.gif" width="600" height="60"></td>
</tr>
<tr>
	<td bgcolor="#f7f7f7" style="padding:15px 20px 15px 40px;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="11" height="17"><img src="<%=imgFingers%>/academy2009/lecture/icon_arrow01.gif" width="11" height="8"></td>
			<td>성함과 휴대전화번호를 정확히 입력하시기 바랍니다.</td>
		</tr>
		<tr>
			<td width="11" height="17"><img src="<%=imgFingers%>/academy2009/lecture/icon_arrow01.gif" width="11" height="8"></td>
			<td>공석 발생 시, 대기 순서대로 SMS를 발송해드리며 매일 오후 1시에 일괄발송합니다.</td>
		</tr>
		<tr>
			<td width="11" height="17"><img src="<%=imgFingers%>/academy2009/lecture/icon_arrow01.gif" width="11" height="8"></td>
			<td>SMS 수신 후, 24시간내에 홈페이지의 [마이핑거스&gt;대기신청조회] 에서 결제하셔야 합니다.</td>
		</tr>
		<tr>
			<td width="11" height="17"><img src="<%=imgFingers%>/academy2009/lecture/icon_arrow01.gif" width="11" height="8"></td>
			<td>24시간내에 결제가 완료되지 않을 경우, 다음 순서의 대기자에게 기회가 돌아갑니다.</td>
		</tr>
		</table>
	</td>
</tr>
<% if RegMode="edit" then %>
<%
dim userphone, userphone1, userphone2, userphone3
userphone = owaiting.FOneItem.Fuser_phone
userphone = split(userphone,"-")

if UBound(userphone)>=0 then
	userphone1 = userphone(0)
end if


if UBound(userphone)>=1 then
	userphone2 = userphone(1)
end if

if UBound(userphone)>=2 then
	userphone3 = userphone(2)
end if

%>
<!--- // 수정모드 // ---->
<tr>
	<td align="center" style="padding:20px 50px 20px 50px;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<form name="waitfrm" method="post" action="/academy/lecture/lib/doLecwait.asp">
		<input type="hidden" name="idx" value="<%= idx %>">
		<input type="hidden" name="mode" value="edit">
		<tr>
			<td style="padding:0 30px 20px 30px;">
			<!---- 강좌정보 START ----->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">강좌번호</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><input type="text" name="lec_idx" value="<%= owaiting.FOneItem.Flec_idx %>" size="6"></td>
				</tr>
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">고객아이디</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;">
						<input type="text" name="userid" value="<%= owaiting.FOneItem.Fuserid %>" size="10">
			            <input type="button" value="검색" class="input_02" onclick="popUserInfo(waitfrm.userid)">
					</td>
				</tr>
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">강좌명</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= owaiting.FOneItem.Flec_title %></td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">강좌일시</td>
					<td style="padding-top:2px;border-bottom:1px solid #eaeaea;"><% Response.Write FormatDateTime(owaiting.FOneItem.FoptLecSDate,1) & " " & FormatDateTime(owaiting.FOneItem.FoptLecSDate,4) & "~" & FormatDateTime(owaiting.FOneItem.FoptLecEDate,4) %></td></tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">강좌비</td>
					<td class="sale11px01" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= FormatNumber(owaiting.FOneItem.FLec_Cost,0) %>원  &nbsp;재료비 <%= FormatNumber(owaiting.FOneItem.FMat_cost,0) %>원</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">현재대기자수</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= owaiting.FOneItem.FoptWaitCnt %>명</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">대기순위</td>
					<td class="gray11px02" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= owaiting.FOneItem.Fregrank %> 순위</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #cacaca;">신청인수</td>
					<td class="gray11px02" style="padding-top:2px;border-bottom:1px solid #cacaca;"><input type="text" name="regcount" size="1" maxlength="1" class="input_02" value="<%= owaiting.FOneItem.Fregcount %>">명 (최대 2명)</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">진행상태</td>
					<td class="gray11px02" style="padding-top:2px;border-bottom:1px solid #eaeaea;">
					<% if owaiting.FOneItem.FCurrState=7 then %>
						<font color="<%= owaiting.FOneItem.getStateNameColor %>"><%= owaiting.FOneItem.getStateName %></font>
					<% else %>
						<select name="currstate">
						<option value="0" <% if owaiting.FOneItem.FCurrState=0 then response.write "selected" %> >대기신청중
						<option value="3" <% if owaiting.FOneItem.FCurrState=3 then response.write "selected" %> >결제대기
						</select>
						<% if (owaiting.FOneItem.FCurrState=3) and (owaiting.FOneItem.IsSettleExpired) then %>
						<font color="red">결제기간만료</font>
						<% end if %>
						<br>
						<% if Not IsNULL(owaiting.FOneItem.Fregendday) then %>
						<input type="text" name="regendday" value="<%= owaiting.FOneItem.Fregendday %>" size="19" maxlength="19">
						<% else %>
						<input type="text" name="regendday" value="" size="19" maxlength="19">
						<% end if %>
					<% end if %>
					</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">사용여부</td>
					<td class="gray11px02" style="padding-top:2px;border-bottom:1px solid #eaeaea;">
						<% if owaiting.FOneItem.Fisusing="Y" then %>
						<input type="radio" name="isusing" value="Y" checked >사용함
						<input type="radio" name="isusing" value="N">사용안함
						<% else %>
						<input type="radio" name="isusing" value="Y">사용함
						<input type="radio" name="isusing" value="N" checked ><font color="red">사용안함</font>
						<% end if %>
					</td>
				</tr>
				</table>
			<!---- 강좌정보 END ----->
			</td>
		</tr>
		<tr>
			<td style="padding:15px 30px 20px 30px;border:4px double #98c573;">
			<!---- 신청정보입력 START ----->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td height="27" colspan="2" style="border-bottom:1px solid #eaeaea;"><img src="<%=imgFingers%>/academy2009/lecture/waitapply_text01.gif" width="86" height="15"></td>
				</tr>
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">대기자 이름</td>
					<td class="green11pxb" style="border-bottom:1px solid #eaeaea;">
						<input type="text" name="username"  class="input_02" value="<%= owaiting.FOneItem.Fuser_name %>" maxlength="20" style="width:100px;height:18px;">
					</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">대기자 휴대전화</td>
					<td style="border-bottom:1px solid #eaeaea;">
						<input name="tel01" type="text" class="input_02" style="width:30px;height:18px;" maxlength="3" value="<%= userphone1 %>" /> -
						<input name="tel02" type="text" class="input_02" style="width:40px;height:18px;" maxlength="4" value="<%= userphone2 %>" /> -
						<input name="tel03" type="text" class="input_02" style="width:40px;height:18px;" maxlength="4" value="<%= userphone3 %>" />
					</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">대기자 이메일</td>
					<td class="sale11px01" style="border-bottom:1px solid #eaeaea;">
						<input name="useremail" type="text" maxlength="64" class="input_02" value="<%= owaiting.FOneItem.Fuser_email %>" style="width:180px;height:18px;">
					</td>
				</tr>
				</table>
			<!---- 신청정보입력 END ----->
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<% else %>
<!----- // 신규모드 // ----->
<tr>
	<td align="center" style="padding:20px 50px 20px 50px;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<form name="waitfrm" method="post" action="/academy/lecture/lib/doLecwait.asp">
		<input type="hidden" name="idx" value="<%= idx %>">
		<input type="hidden" name="mode" value="add">
		<tr>
			<td style="padding:0 30px 20px 30px;">
			<!---- 강좌정보 START ----->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">강좌번호</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><input type="text" name="lec_idx" value="<%= lec_idx %>" size="6"></td>
				</tr>
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">고객아이디</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;">
						<input type="text" name="userid" value="" size="10">
			            <input type="button" value="검색" class="input_02" onclick="popUserInfo(waitfrm.userid)">
					</td>
				</tr>
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">강좌명</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= olecture.FOneItem.Flec_title %></td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">강좌일</td>
					<td style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= getLecOptionBoxHTML(lec_idx,"lecOption","AddWait") %></td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">강좌비</td>
					<td class="sale11px01" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><%= FormatNumber(olecture.FOneItem.FLec_Cost,0) %>원  &nbsp;재료비 <%= FormatNumber(olecture.FOneItem.FMat_cost,0) %>원</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">현재대기자수</td>
					<td class="green11pxb" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><span id="msgSold"><%= olecture.FOneItem.FWaitcount %>명</span></td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">신청인수</td>
					<td class="gray11px02" style="padding-top:2px;border-bottom:1px solid #eaeaea;"><input type="text" name="regcount" size="1" maxlength="1" class="input_02" value="1">명 (최대 2명)</td>
				</tr>
				</table>
			<!---- 강좌정보 END ----->
			</td>
		</tr>
		<tr>
			<td style="padding:15px 30px 20px 30px;border:4px double #98c573;">
			<!---- 신청정보입력 START ----->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td height="27" colspan="2" style="border-bottom:1px solid #eaeaea;"><img src="<%=imgFingers%>/academy2009/lecture/waitapply_text01.gif" width="86" height="15"></td>
				</tr>
				<tr>
					<td width="90" height="27" style="border-bottom:1px solid #eaeaea;">대기자 이름</td>
					<td class="green11pxb" style="border-bottom:1px solid #eaeaea;">
						<input type="text" name="username"  class="input_02" value="" maxlength="20" style="width:100px;height:18px;">
					</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">대기자 휴대전화</td>
					<td style="border-bottom:1px solid #eaeaea;">
						<input name="tel01" type="text" class="input_02" style="width:30px;height:18px;" maxlength="3" value="" /> -
						<input name="tel02" type="text" class="input_02" style="width:40px;height:18px;" maxlength="4" value="" /> -
						<input name="tel03" type="text" class="input_02" style="width:40px;height:18px;" maxlength="4" value="" />
					</td>
				</tr>
				<tr>
					<td height="27" style="border-bottom:1px solid #eaeaea;">대기자 이메일</td>
					<td class="sale11px01" style="border-bottom:1px solid #eaeaea;">
						<input name="useremail" type="text" maxlength="64" class="input_02" value="" style="width:180px;height:18px;">
					</td>
				</tr>
				</table>
			<!---- 신청정보입력 END ----->
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" style="padding-bottom:20px;">
		<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td><img src="<%=imgFingers%>/academy2009/lecture/btn_waitapply02.gif" width="131" height="42" border="0" align="absmiddle" onclick="frmsub(waitfrm);" style="cursor:pointer"></td>
			<td style="padding-left:20px;"><img src="<%=imgFingers%>/academy2009/lecture/btn_cancel03.gif" width="131" height="42" onclick="self.close()" style="cursor:pointer"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<% set owaiting = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->