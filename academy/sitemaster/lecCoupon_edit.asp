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

	if (confirm('강좌 보너스 쿠폰을 발행하시겠습니까?\n\n※발행된 쿠폰은 취소할 수 없으므로 신중히 확인하세요.')){
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
<form name="frmarr" method="post" action="lecCouponedit_Process.asp">
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
<tr>
	<td bgcolor="#E6E6E6" align="center">쿠폰명</td>
	<td bgcolor="#FFFFFF"><input type=text name=couponname maxlength="100" size=80>
	<br>(Ex. 우수회원을 위한 특별한 강좌 10%할인 쿠폰)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">최소구매금액</td>
	<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>원 이상 강좌 수강시 사용가능(숫자)</td>
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