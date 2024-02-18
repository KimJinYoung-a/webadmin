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

	if ((upfrm.coupontype[0].checked == true) && (upfrm.couponvalue.value*1 > 15)) {
		// 사고방지
		alert('15% 를 넘는 할인쿠폰은 생성할 수 없습니다.');
		upfrm.couponvalue.focus();
		return;
	}

	if ((upfrm.coupontype[1].checked == true) && (upfrm.couponvalue.value*1 > upfrm.minbuyprice.value*0.2)) {
		// 사고방지
		alert('정액할인액이 최소구매금액의 20% 를 넘을 수 없습니다.');
		upfrm.couponvalue.focus();
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

		frmarr.targetitemlist.style.backgroundColor = "<%= adminColor("tabletop") %>";
		frmarr.couponmeaipprice.style.backgroundColor = "<%= adminColor("tabletop") %>";
	}

}
</script>

<table width="760" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="#B2B2B2">
<form name="frmarr" method="post" action="eventcouponedit_Process_off.asp">
<input type="hidden" name="mode" value="">
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" width="130" align="center">아이디추가</td>
	<td align="left">
		<input type="text" class="text" name="useridarr" value="" size="40">
		<font color="#FF6699">*** 콤마로 구분(ex : corpse2,icommang)</font>
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">쿠폰타입</td>
	<td bgcolor="#FFFFFF">
	<input type="text" class="text" name=couponvalue maxlength=7 size=10>
	<input type=radio name=coupontype value="1" onclick="alert('% 할인 쿠폰입니다.');">%할인
	<input type=radio name=coupontype value="2" checked >원할인
	(금액 또는 % 할인)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">쿠폰명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name=couponname maxlength="100" size=80>
	<br>(2004년 1월 부터 5월까지 10만원 이상 구매고객에게 드리는 상품권입니다.)</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">최소구매금액</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name=minbuyprice maxlength=7 size=10>원 이상 구매시 사용가능(숫자)</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">유효기간</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name=startdate value="<%= left(now(),10) %> 00:00:00" maxlength=19 size=20> ~ <input type="text" class="text" name=expiredate maxlength=19 size=20>(<%= left(now(),10) %> 00:00:00 ~ <%= left(now(),10) %> 23:59:59)</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">발급자 ID </td>
	<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
</tr>
<tr height=30>
	<td colspan="2" align=center bgcolor="#FFFFFF"><input type=button class=button value="저장" onClick="submitForm(this.form);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->