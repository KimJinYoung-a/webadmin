<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  인트라넷 개인정보 수정
' History : 2007.07.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/10x10staffcls.asp"-->
<%
dim txBirthday1, txBirthday2, txBirthday3, txPhone1, txPhone2 , birth_isSolar
dim txPhone3, txPhone4, txCell1, txCell2, txCell3, txZip1, txZip2

dim onepartner
	set onepartner = new CPartnerUser
	onepartner.GetOnePartner session("ssBctId")

dim birthdayarr,telarr,hparr,ziparr
	if onepartner.Fbirthday <> "" then
		birthdayarr = split(left(onepartner.Fbirthday,10),"-")
			if ubound(birthdayarr) >= 2 then
				txBirthday1 = birthdayarr(0)
				txBirthday2 = birthdayarr(1)
				txBirthday3 = birthdayarr(2)
			end if
	end if
	if onepartner.Ftel <> "" then
		telarr = split(onepartner.Ftel,"-")
			if ubound(telarr) >= 3 then
				txPhone1 = telarr(0)
				txPhone2 = telarr(1)
				txPhone3 = telarr(2)
				txPhone4 = telarr(3)
			end if
	end if
	if onepartner.Fmanager_hp <> "" then
		hparr = split(onepartner.Fmanager_hp,"-")
			if ubound(hparr) >= 2 then
				txCell1 = hparr(0)
				txCell2 = hparr(1)
				txCell3 = hparr(2)
			end if
	end if
	if onepartner.Fzipcode <> "" then
		ziparr = split(onepartner.Fzipcode,"-")
			if ubound(ziparr) >= 1 then
				txZip1  = ziparr(0)
				txZip2  = ziparr(1)
			end if
	end if
	if onepartner.Fzipcode <> "" then
		ziparr = split(onepartner.Fzipcode,"-")
			if ubound(ziparr) >= 1 then
				txZip1  = ziparr(0)
				txZip2  = ziparr(1)
			end if
	end if

	if onepartner.fbirth_isSolar = "Y" or onepartner.fbirth_isSolar = "" or isnull(onepartner.fbirth_isSolar) then
		birth_isSolar = "Y"
	else
		birth_isSolar = "N"
	end if
%>

<script language="javascript">

function TnFindZip(frmname){
	window.open('/lib/searchzip2.asp?target=' + frmname, 'findzipcdode', 'width=460,height=250,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
}

function TnJoin10x10(frm){
	if (frm.txBirthday1.value == ''){
		alert("생년월일을 입력하세요");
		frm.txBirthday1.focus();
	return;
	}
	if (frm.txBirthday2.value == ''){
		alert("생년월일을 입력하세요");
		frm.txBirthday2.focus();
	return;
	}
	if (frm.txBirthday3.value == ''){
		alert("생년월일을 입력하세요");
		frm.txBirthday3.focus();
	return;
	}
	if (frm.txPhone1.value == ''){
		alert("전화번호를 입력하세요");
		frm.txPhone1.focus();
	return;
	}
	if (frm.txPhone2.value == ''){
		alert("전화번호를 입력하세요");
		frm.txPhone2.focus();
	return;
	}
	if (frm.txPhone3.value == ''){
		alert("전화번호를 입력하세요");
		frm.txPhone3.focus();
	return;
	}
	if (frm.txPhone4.value == ''){
		alert("전화번호를 입력하세요");
		frm.txPhone4.focus();
	return;
	}
	if (frm.txCell1.value == ''){
		alert("핸드폰번호를 입력하세요");
		frm.txCell1.focus();
	return;
	}
	if (frm.txCell2.value == ''){
		alert("핸드폰번호를 입력하세요");
		frm.txCell2.focus();
	return;
	}
	if (frm.txCell3.value == ''){
		alert("핸드폰번호를 입력하세요");
		frm.txCell3.focus();
	return;
	}
	if (frm.txZip1.value == ''){
		alert("우편번호를 입력하세요");
		frm.txZip1.focus();
	return;
	}
	if (frm.txZip2.value == ''){
		alert("우편번호를 입력하세요");
		frm.txZip2.focus();
	return;
	}
	if (frm.txAddr1.value == ''){
		alert("주소를 입력하세요");
		frm.txAddr1.focus();
	return;
	}
	if (frm.txAddr2.value == ''){
		alert("주소를 입력하세요");
		frm.txAddr2.focus();
	return;
	}
	var ret = confirm('수정 하시겠습니까?');
	if(ret){
	frm.submit();
	}
}

function passwordedit(frm){
	if (frm.txpass2.value != frm.txpass3.value){
		alert("변경하실 비밀번호가 일치하지 않습니다. 두번 정확히 입력해주십시오.");
		frm.txpass2.value=""
		frm.txpass3.value=""
		frm.txpass2.focus();
		return ;
	}

	if (frm.txpass1.value == "" || frm.txpass2.value == "" || frm.txpass3.value == ""){
		alert("기존비밀번호와 변경할비밀번호를 정확히 입력해 주세요.");
		frm.txpass1.value=""
		frm.txpass2.value=""
		frm.txpass3.value=""
		frm.txpass1.focus();
		return ;
	}

	var ret = confirm('비밀번호를 수정 하시겠습니까?');
		if(ret){
			frm.submit();
		}
	}

</script>

<br><br><br><br><font color=red>삭제 예정 페이지</font><br><br><br><br>

<!--기본정보변경 시작-->
<table width="48%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="myinfoForm" method="post" action="domyinfo.asp">
	<input type="hidden" name="userid" value="<%= onepartner.Fid %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>기본정보</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td width="120" bgcolor="<%= adminColor("tabletop") %>">이름</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txName" id="[on,off,2,16][성명]" type="text" size="10" class="text" value="<%= onepartner.Fcompany_name %>">
    		성명 기입시 공백이 없이 입력 하여 주십시요.(예 : 홍길동)
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">아이디</td>
    	<td bgcolor="#FFFFFF">
    		<%= onepartner.Fid %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">생년월일</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txBirthday1" id="[on,on,4,4][태어난해]" size="4" maxlength="4" class="text" value="<%= txBirthday1 %>">년
			<input type="text" name="txBirthday2" id="[on,on,2,2][태어난달]" size="4" maxlength="2" class="text" value="<%= txBirthday2 %>">월
			<input type="text" name="txBirthday3" id="[on,on,2,2][태어난일]" size="4" maxlength="2" class="text" value="<%= txBirthday3 %>">일
			&nbsp; &nbsp; &nbsp; &nbsp; 양력:<input type="radio" name="birth_isSolar" value="Y" <% if birth_isSolar = "Y" then response.write " checked" %>>
			음력:<input type="radio" name="birth_isSolar" value="N" <% if birth_isSolar = "N" then response.write " checked" %>>
    	</td>
    </tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(사내메일)</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txEmail1" id="[on,off,off,off][사내메일]" type="text" size="30" class="text" value="<%= onepartner.Femail %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">MSN메신저</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txEmail2" id="[on,off,off,off][MSN메신저]" type="text" size="30" class="text" value="<%= onepartner.Fmsn %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">전화번호(내선)</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txPhone1" id="[on,on,2,4][전화번호1]" size="5" class="text" value="<% = txPhone1 %>">-
			<input type="text" name="txPhone2" id="[on,on,2,4][전화번호2]" size="5" class="text" value="<% = txPhone2 %>">-
			<input type="text" name="txPhone3" id="[on,on,2,4][전화번호3]" size="5" class="text" value="<% = txPhone3 %>">&nbsp;&nbsp;내선:
			<input type="text" name="txPhone4" id="[on,on,2,4][내선]" size="5" class="text" value="<%= txPhone4 %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">070 직통번호</td><!-- 신규 -->
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="" id="" size="5" class="text" value="070">-
			<input type="text" name="" id="" size="5" class="text" value="0000">-
			<input type="text" name="" id="" size="5" class="text" value="0000">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰번호</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txCell1" id="[on,on,2,4][핸드폰번호1]" size="5" class="text" value="<% = txCell1 %>">-
			<input type="text" name="txCell2" id="[on,on,2,4][핸드폰번호2]" size="5" class="text" value="<% = txCell2 %>">-
			<input type="text" name="txCell3" id="[on,on,2,4][핸드폰번호3]" size="5" class="text" value="<% = txCell3 %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">우편번호</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txZip1"  id="[on,on,3,3][우편번호1]"  size="5" style="background-color:#EEEEEE;" class="text" readonly  value="<%= txZip1 %>">-
			<input type="text" name="txZip2"  id="[on,on,3,3][우편번호2]" size="5" style="background-color:#EEEEEE;" class="text" readonly  value="<% = txZip2 %>">
			<input type="button" class="button_s" value="주소입력" onClick="javascript:TnFindZip('myinfoForm');">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">주소</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txAddr1" id="[on,off,1,64][주소1]"  size="50" style="background-color:#EEEEEE;" class="text" readonly  value="<%= onepartner.Faddress %>">
			<br> <input type="text" name="txAddr2" id="[on,off,1,64][주소2]" size="50" maxlength="128" class="text"  value="<%= onepartner.Fmanager_address %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">부서-파트</td>
    	<td bgcolor="#FFFFFF">
    		<%= onepartner.fpart_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">어드민권한</td>
    	<td bgcolor="#FFFFFF">
    		관리자 / 마스터 / 파트선임권한 / 파트구성원권한 ......
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">직급</td>
    	<td bgcolor="#FFFFFF">
    		<%= onepartner.fposit_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">직책</td><!-- 신규 -->
    	<td bgcolor="#FFFFFF">
    		팀장 / 파트장 / 파트구성원
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">담당업무(카테고리)</td><!-- 신규 -->
    	<td bgcolor="#FFFFFF">
    		MD의 경우 담당 카테고리 선택등록(아니면 디비로 지정)
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td colspan="2" bgcolor="#FFFFFF">
    		<input type="button" class="button_s" value="기본정보 수정" onclick="javascript:TnJoin10x10(myinfoForm)">
    	</td>
    </tr>
		<!--	<a href="javascript:TnFindZip('myinfoForm')" ><img src="/images/page_2_3.gif" width="60" height="20" border="0" align="absmiddle"></a>	-->

	</form>
<!--기본정보변경 끝-->

<!--비밀번호변경 시작-->
	<form name="myinfopassword" method="post" action="domyinfo_password.asp">
	<tr>
		<td valign="bottom" colspan=2 bgcolor="FFFFFF">
			<font color="red"><strong>비밀번호 수정</strong></font> *비밀번호 수정시에만 입력하세요.
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">기존비밀번호</td>
    	<td bgcolor="#FFFFFF">
    		<input  type="password" name="txpass1" id="[off,off,off,off][기존비밀번호]"  size="16" class="input_01">
    		*기존 비밀번호를 입력하세요.
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">신규비밀번호</td>
    	<td bgcolor="#FFFFFF">
    		<input  type="password" name="txpass2" id="[off,off,off,off][비밀번호변경]"  size="16" class="input_01">
			<input  type="password" name="txpass3" id="[off,off,off,off][비밀번호변경]"  size="16" class="input_01">
    		*사용하실 비밀번호를 두번 입력해 주세요.
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td colspan="2" bgcolor="#FFFFFF">
			<input type="button" class="button" value="비밀번호 수정" onclick="javascript:passwordedit(myinfopassword)">
    	</td>
    </tr>
	</form>
<!--비밀번호변경 끝-->

</table>

<table width="50%" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>추가정보</strong></font>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">입사일</td>
    	<td width="100">근속연수</td>
    	<td width="100"></td>
      	<td width="100"></td>
      	<td width="100"></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>2001-08-23</td>
    	<td>9</td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>연차(휴가)정보</strong></font>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>전년이월</td>
    	<td>금년</td>
    	<td>2010년 총연차</td>
      	<td>사용일수</td>
      	<td>잔여일수</td>
      	<td>비고</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>5</td>
    	<td>12</td>
      	<td>17</td>
      	<td>4</td>
      	<td>13</td>
      	<td><input type="button" class="button" value="휴가신청 및 내역보기" onclick=""></td>
    </tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			*전년이월연차는 3월까지만 유효하며, 휴가신청시 전년이월연차부터 차감됩니다.
		</td>
	</tr>
</table>



<%
set onepartner = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->