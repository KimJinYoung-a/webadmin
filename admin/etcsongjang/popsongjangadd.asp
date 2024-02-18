<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 이벤트 당첨자
' History : 2009.04.17 최초생성자 모름
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;

    eval(frmname + ".addr1").value = addr;
    eval(frmname + ".addr2").value = dong;
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function delThis(){
    var frm = document.infoform;

    if (confirm('삭제 하시겠습니까?')){
        if (confirm('정말로 삭제 하시겠습니까?')){
            frm.mode.value="del";
    		frm.submit();
		}
	}
}

function gotowrite(){
    var frm = document.infoform;
	if(frm.gubuncd.value == ""){
		alert("구분을 선택해주세요.");
	    frm.gubuncd.focus();
	    return;
	}

	if(frm.gubunname.value == ""){
		alert("이벤트명(구분명)을 입력해주세요.");
	    frm.gubunname.focus();
	    return;
	}

	if(frm.prizetitle.value == ""){
		alert("상품명을 입력해주세요.");
	    frm.prizetitle.focus();
	    return;
	}

	if ((frm.useDefaultAddr[1].checked == true) || (frm.useDefaultAddr[2].checked == true)) {
		if (frm.userid.value == '') {
			alert('아이디가 입력된 경우에만 선택가능합니다.');
			return;
		}
	} else {
		if(frm.username.value == ""){
			alert("당첨자성함을 입력해주세요.");
			frm.username.focus();
			return;
		}

		if(frm.reqname.value == ""){
			alert("받으시는 분의 이름을 입력해주세요.");
			frm.reqname.focus();
			return;
		}

		if(frm.reqphone1.value == "" || frm.reqphone2.value == "" || frm.reqphone3.value == ""){
			alert("받으시는 분의 전화번호를 입력해주세요.");
			frm.reqphone1.focus();
			return;
		}

		if(frm.reqhp1.value == "" || frm.reqhp2.value == "" || frm.reqhp3.value == ""){
			alert("받으시는 분의 핸드폰 번호를 입력해주세요.");
			frm.reqphone1.focus();
			return;
		}

		if(frm.zipcode.value == ""){
			alert("받으시는 분의 주소를 입력해주세요.");
			frm.zipcode.focus();
			return;
		}

		if(frm.addr2.value == ""){
			alert("받으시는 분의 나머지주소를 입력해주세요.");
			frm.addr2.focus();
			return;
		}
	}

	if (frm.reqdeliverdate.value.length!=10){
	    alert('출고 요청일을 입력하세요.');
	    frm.reqdeliverdate.focus();
	    return;
	}

	if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
	    alert('배송 구분을 선택 하세요.');
	    frm.isupchebeasong[0].focus();
	    return;
	}
	if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&(frm.jungsanValue.value=="")){
	    alert('정산액(매입가)를 입력하세요');
	    frm.jungsanValue.focus();
	    return;
	}
	if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
	    alert('업체 배송인 경우 브랜드 아이디를  선택 하세요.');
	    frm.makerid.focus();
	    return;
	}
	if (confirm('입력 내용이 정확합니까?')){
		frm.submit();
	}
}

function disabledBox(comp){
    var frm = comp.form;
    if (comp.value=="Y"){
        frm.makerid.disabled = false;
        frm.jungsan.disabled = false;

		frm.jungsanValue.disabled = false;
        frm.jungsan.checked = true;
    }else{
        frm.makerid.selectedIndex = 0;
        frm.makerid.value = '';
		frm.makerid.disabled = true;
		frm.jungsan.disabled = true;

        frm.jungsanValue.value = '';
        frm.jungsanValue.disabled = true;
        frm.jungsan.checked = false;
    }
}

function disableAddressBox(comp) {
    var frm = comp.form;
    if (comp.value != "C"){
        frm.username.disabled = true;
        frm.reqname.disabled = true;
		frm.reqphone1.disabled = true;
		frm.reqphone2.disabled = true;
		frm.reqphone3.disabled = true;
		frm.reqhp1.disabled = true;
		frm.reqhp2.disabled = true;
		frm.reqhp3.disabled = true;
		frm.zipcode.disabled = true;
		frm.addr2.disabled = true;
    }else{
        frm.username.disabled = false;
        frm.reqname.disabled = false;
		frm.reqphone1.disabled = false;
		frm.reqphone2.disabled = false;
		frm.reqphone3.disabled = false;
		frm.reqhp1.disabled = false;
		frm.reqhp2.disabled = false;
		frm.reqhp3.disabled = false;
		frm.zipcode.disabled = false;
		frm.addr2.disabled = false;
    }

	var evtprize_enddate = $('.evtprize_enddate');
	if (comp.value == "N") {
		evtprize_enddate.show();
	} else {
		evtprize_enddate.hide();
	}
}

function jungsanYN(){
	var frm = document.infoform;
	if(frm.jungsan.checked==true){
		frm.jungsanValue.disabled = false;
	}else{
		frm.jungsanValue.value = '';
		frm.jungsanValue.disabled = true;
	}
}
function checkover1(obj) {
	var val = obj.value;
	if (val) {
		if (val.match(/^\d+$/gi) == null) {
			alert("숫자만 넣으세요!");
			document.infoform.jungsanValue.value = '';
			obj.select();
			return;
		}
	}
}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing=0 class="a">
<form name="infoform" method="post" action="/admin/etcsongjang/lib/doeventbeasonginfo.asp">
<input type="hidden" name="mode" value="I">
<tr>
	<td align="center">
		<table width="90%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr height="30">
			<td height="2" colspan="2" >* 기타출고 정보 입력</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">구분</td>
			<td style="padding-left:7">
				<select name="gubuncd" class="select">
					<option value="">전체
<!--
					<option value="96">고객
					<option value="97">29cm용
-->
					<option value="98">판촉
					<option value="99">기타
					<option value="80">CS출고
				</select>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">이벤트명(구분명) </td>
			<td style="padding-left:7">
				<input type="text" class="text" name="gubunname" size="40" maxlength="64" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">당첨상품</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="prizetitle" size="40" maxlength="64" value="" > * <font color="red">프런트노출</font>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">아이디</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="userid" size="20" maxlength="32" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">배송지 등록구분</td>
			<td style="padding-left:7">
				<input type=radio name="useDefaultAddr" value="C" checked onClick="disableAddressBox(this)">직접입력
				<input type=radio name="useDefaultAddr" value="N" onClick="disableAddressBox(this)">User 가 배송지 입력
				<input type=radio name="useDefaultAddr" value="Y" onClick="disableAddressBox(this)">User 기본 주소 사용
			</td>
		</tr>
		<tr height="1" class="evtprize_enddate" style="display: none;">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr class="evtprize_enddate" style="display: none;">
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">배송지입력 마감일</td>
			<td style="padding-left:7">
				<input type="text" class="text_ro" name="evtprize_enddate" size="10" maxlength="10"  value="<%= Left(DateAdd("m", 3, Now()), 10) %>">
				<a href="javascript:jsPopCal('evtprize_enddate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">당첨자성함</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="username" size="20" maxlength="20" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">수령인성함</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="reqname" size="20" maxlength="20" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">연락처</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqphone1" size="3" class="verdana_s" maxlength="3" value="">
				-
				<input type="text" class="text" name="reqphone2" size="4" class="verdana_s" maxlength="4" value="">
				-
				<input type="text" class="text" name="reqphone3" size="4" class="verdana_s" maxlength="4" value="">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">핸드폰</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqhp1" size="3" class="verdana_s"  maxlength="3" value="">
				-
				<input type="text" class="text" name="reqhp2" size="4" class="verdana_s"  maxlength="4" value="">
				-
				<input type="text" class="text" name="reqhp3" size="4" class="verdana_s"  maxlength="4" value="">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">수령인 주소</td>
			<td class="verdana_s" style="padding:5 0 5 7">
				<input type="text" class="text_ro" name="zipcode" size="7" class="verdana_s" readOnly value="">
				<input type="button" class="button" value="검색" onClick="FnFindZipNew('infoform','E')">
				<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('infoform','E')">
				<% '<input type="button" value="검색(구)" class="button" onclick="PopSearchZipcode('infoform');" onFocus="this.blur();"> %>
				<br>
				<input type="text" class="text_ro" name="addr1" size="16" maxlength="64"  readOnly value="" ><br>
				<input type="text" class="text" name="addr2" size="40" maxlength="64" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">기타요청사항</td>
			<td class="verdana_s" style="padding:5 0 5 7"><textarea class="text" name="reqetc" class="textarea" style="width:350px;height:40px;"></textarea></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">출고요청일</td>
			<td class="verdana_s" style="padding:5 0 5 7">
			<input type="text" class="text_ro" name="reqdeliverdate" size="10" maxlength="10"  value="" >
			<a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">배송구분</td>
			<td style="padding-left:7">
				<input type=radio name=isupchebeasong value="N" checked onClick="disabledBox(this);">텐바이텐배송
				<input type=radio name=isupchebeasong value="Y" onClick="disabledBox(this);">업체직접배송
			<br>
			<% drawSelectBoxDesignerwithName "makerid", "" %>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<script>
			document.infoform.makerid.disabled = true;
		</script>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">정산여부</td>
			<td style="padding-left:7">
				<input type="checkbox" class="checkbox" name="jungsan" id="jungsan" onclick="javascript:jungsanYN();" disabled >정산함&nbsp;&nbsp;
				정산액(매입가) : <input type="text" class="text" id="jungsanValue" name="jungsanValue" value="" onkeyup="checkover1(this)">원
			</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
		<tr height="30">
			<td colspan="2" align="center">
			<input type="button" class="button" value=" 저 장 " onClick="gotowrite();" onfocus="this.blur();">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
