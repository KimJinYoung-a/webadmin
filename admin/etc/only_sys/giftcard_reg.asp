<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script>
function jsGiftCardReg(){
	if(frm1.iid.value == ""){
		alert("텐바이텐 Gift 카드를 선택하세요.");
		frm1.iid.focus();
		return;
	}
	if(frm1.opt.value == ""){
		alert("금액권을 선택하세요.");
		frm1.opt.focus();
		return;
	}
	if(frm1.mmstitle.value == ""){
		alert("MMS 제목을 입력하세요.");
		frm1.mmstitle.focus();
		return;
	}
	if(frm1.mmsmessage.value == ""){
		alert("MMS 메세지를 입력하세요.");
		frm1.mmsmessage.focus();
		return;
	}
	if(frm1.userid.value == ""){
		alert("아이디를 입력하세요.");
		frm1.userid.focus();
		return;
	}
	frm1.submit();
}
</script>

<form name="frm1" action="giftcard_reg_proc.asp" method="post" style="margin:0px;">
<table class="a">
<tr>
	<td>
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td height="50">
				기프트카드번호 : <select name="iid"><option value="">-선택-</option><option value="101" selected>[101]텐바이텐 Gift 카드</option></select>
				&nbsp;&nbsp;&nbsp;
				<select name="opt">
					<option value="">-금액권선택-</option>
					<option value="0001">1만원권</option>
					<option value="0002">2만원권</option>
					<option value="0003">3만원권</option>
					<option value="0004">5만원권</option>
					<option value="0005">8만원권</option>
					<option value="0006">10만원권</option>
					<option value="0007">15만원권</option>
					<option value="0008">20만원권</option>
					<option value="0009">30만원권</option>
				</select>
			</td>
		</tr>
		<tr>
			<td height="50">
				MMS 제목 : <font color="red">※ ' " 제거 필수.</font><br>
				<input type="text" name="mmstitle" id="mmstitle" value="" size="70"><br><br>
			</td>
		</tr>
		<tr>
			<td height="50">
				MMS 메세지 : <font color="red">※ ' " 제거 필수.</font><br>
				<textarea name="mmsmessage" id="mmsmessage" rows="4" cols="100">[텐바이텐] 고급진 더블마일리지를 누려 상품후기 이벤트에 당첨되셨습니다.
당첨되신 분께는 텐바이텐 기프트 카드 1만원권을 드립니다.
기프트 카드는 마이텐바이텐에서 확인 가능합니다.</textarea><br><br>
			</td>
		</tr>
		<tr>
			<td height="50">
				발급 아이디 : <br>
				<textarea name="userid" id="userid" rows="10" cols="100"></textarea><br><br>
				<input type="button" class="button" value="실 행" style="width:100px;height:60px;" onClick="jsGiftCardReg()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->