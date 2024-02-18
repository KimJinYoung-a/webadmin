<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->

<script language='javascript'>

function submitForm(upfrm){
	if (upfrm.useridarr.value == ""){
		alert("아이디를 입력해주세요!");
		upfrm.useridarr.focus();
		return;
	}

	if(upfrm.couponvalue.value == ""){
		alert("쿠폰금액 또는 할인율을 입력해주세요!");
		upfrm.couponvalue.focus();
		return;
	}

	if(upfrm.couponname.value == ""){
		alert("쿠폰명을 입력해주세요!");
		upfrm.couponname.focus();
		return;
	}

	if(upfrm.minbuyprice.value == ""){
		alert("최소금액을 입력해주세요!");
		upfrm.minbuyprice.focus();
		return;
	}

	if(upfrm.startdate.value == "" || upfrm.expiredate.value == ""){
		alert("사용기간을 입력해주세요!");
		return;
	}

//태훈 2006-05-09 주석처리
//	if (upfrm.targetitemusing.checked){
//		if (!IsDigit(upfrm.targetitemlist.value)) {
//			alert('상품번호는 숫자만 가능합니다.');
//			upfrm.targetitemlist.focus();
//			return;
//		}

//		if ((upfrm.couponmeaipprice.value!='')&&(!IsDigit(upfrm.couponmeaipprice.value))) {
//			alert('매입가는 숫자만 가능합니다.');
//			upfrm.couponmeaipprice.focus();
//			return;
//		}
//	}

	if (confirm('쿠폰을 발행하시겠습니까?')){
		upfrm.submit();
	}
}

function EnableBox(comp){
	if (comp.checked){
		frmarr.targetitemlist.disabled = false;
		frmarr.couponmeaipprice.disabled = false;

		frmarr.targetitemlist.style.backgroundColor = "#FFFFFF";
		frmarr.couponmeaipprice.style.backgroundColor = "#FFFFFF";
	}else{
		frmarr.targetitemlist.disabled = true;
		frmarr.couponmeaipprice.disabled = true;

		frmarr.targetitemlist.style.backgroundColor = "#E6E6E6";
		frmarr.couponmeaipprice.style.backgroundColor = "#E6E6E6";
	}

}
</script>
<font color="#FF6699">*** 콤마로 구분(ex : corpse2,icommang)</font>
<table width="760" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="#B2B2B2">
<form name="frmarr" method="post" action="eventcouponedit_Process.asp">
<input type="hidden" name="mode" value="">
<tr>
	<td bgcolor="#E6E6E6" width="130" align="center">아이디추가</td>
	<td width="200" align="right"><input type="text" name="useridarr" value="" size="80"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">쿠폰타입</td>
	<td bgcolor="#FFFFFF">
	<input type=text name=couponvalue maxlength=7 size=10>
	<input type=radio name=coupontype value="1" onclick="alert('% 할인 쿠폰입니다.');">%할인
	<input type=radio name=coupontype value="2" checked >원할인
	(금액 또는 % 할인)
	</td>
</tr>
<!-- 사용안함 : 상품쿠폰으로 일반화
<tr>
	<td bgcolor="#E6E6E6" align="center">특정상품쿠폰</td>
	<td bgcolor="#FFFFFF">
		특정상품 쿠폰 사용함: <input type=checkbox name=targetitemusing onclick="EnableBox(this)"><br>
		상품번호: <input type=text name=targetitemlist size=9 maxlength=9 disabled style='background-color:#E6E6E6;'>(특정 상품만 할인됨)
		&nbsp;&nbsp;
		쿠폰적용시 매입가: <input type=text name=couponmeaipprice size=7 maxlength=9 disabled style='background-color:#E6E6E6;'>(업체부담할 경우 매입가 지정)
	</td>
</tr>
-->
<tr>
	<td bgcolor="#E6E6E6" align="center">쿠폰명</td>
	<td bgcolor="#FFFFFF"><input type=text name=couponname maxlength="100" size=80>
	<br>(2004년 1월 부터 5월까지 10만원 이상 구매고객에게 드리는 상품권입니다.)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">최소구매금액</td>
	<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>원 이상 구매시 사용가능(숫자)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">유효기간</td>
	<td bgcolor="#FFFFFF"><input type=text name=startdate value="<%= left(now(),10) %> 00:00:00" maxlength=19 size=19>~<input type=text name=expiredate maxlength=19 size=19>(<%= left(now(),10) %> 00:00:00 ~ <%= left(now(),10) %> 23:59:59)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">발급자 ID </td>
	<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
</tr>
<tr>
	<td colspan="2" align=center bgcolor="#FFFFFF"><input type=button value="저장" onClick="submitForm(this.form);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->