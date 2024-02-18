<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!DOCTYPE html>
<html lang="ko">
<html>
	<head>
		<meta http-equiv="X-UA-Compatible" content="IE=edge" />
		<script type="text/javascript">

		function jsSubmitLogin() {
			if (confirm("로그인 하시겠습니까?") != true) {
				return;
			}

			var popwin = window.open("","jsSubmitLogin","width=300 height=100 scrollbars=yes resizable=yes");
			var frm = document.frm;
			popwin.focus();

			frm.mode.value = "login";
			frm.target = "jsSubmitLogin";
			frm.submit();
		}

		function jsSubmitOrder() {
			if (confirm("주문접수 하시겠습니까?") != true) {
				return;
			}

			var popwin = window.open("","jsSubmitOrder","width=300 height=100 scrollbars=yes resizable=yes");
			var frm = document.frm;
			popwin.focus();

			frm.mode.value = "sendorder";
			frm.target = "jsSubmitOrder";
			frm.submit();
		}

		function jsSubmitCheckPrice() {
			var popwin = window.open("","jsSubmitCheckPrice","width=300 height=100 scrollbars=yes resizable=yes");
			var frm = document.frm;
			popwin.focus();

			frm.mode.value = "checkprice";
			frm.target = "jsSubmitCheckPrice";
			frm.submit();
		}

		function jsSubmitOrderList() {
			var popwin = window.open("","jsSubmitOrderList","width=300 height=100 scrollbars=yes resizable=yes");
			var frm = document.frm;
			popwin.focus();

			frm.mode.value = "orderlist";
			frm.target = "jsSubmitOrderList";
			frm.submit();
		}

		function jsSubmitOrderView() {
			var popwin = window.open("","jsSubmitOrderView","width=300 height=100 scrollbars=yes resizable=yes");
			var frm = document.frm;
			popwin.focus();

			frm.mode.value = "vieworder";
			frm.target = "jsSubmitOrderView";
			frm.submit();
		}

		function jsSubmitOrderCancel() {
			var popwin = window.open("","jsSubmitOrderCancel","width=300 height=100 scrollbars=yes resizable=yes");
			var frm = document.frm;
			popwin.focus();

			frm.mode.value = "cancelorder";
			frm.target = "jsSubmitOrderCancel";
			frm.submit();
		}

		</script>
	</head>
	<body>
		<input type="button" class="button" value=" 로그인 " onClick="jsSubmitLogin();">
		<input type="button" class="button" value="주문전송" onClick="jsSubmitOrder();">
		<input type="button" class="button" value="가격조회" onClick="jsSubmitCheckPrice();">
		<input type="button" class="button" value="주문목록" onClick="jsSubmitOrderList();">
		&nbsp;
		<input type="button" class="button" value="주문조회" onClick="jsSubmitOrderView();">
		<input type="button" class="button" value="주문취소" onClick="jsSubmitOrderCancel();">

		<br /><br />

		<form name="frm" method="post" action="naldo_act.asp">
			<input type="hidden" name="mode" value="">

			<input type="text" class="text" name="orderNumber" value="5021743566" size="10"><br /><br />

			<input type="text" class="text" name="orderPhoneNumber" value="0101111111" size="10">
			<input type="text" class="text" name="senderPhoneNumber" value="0101111111" size="10">
			<input type="text" class="text" name="receiverPhoneNumber" value="0101111111" size="10"><br />

			<input type="text" class="text" name="receiverName" value="받는사람" size="10">
			<input type="text" class="text" name="senderName" value="보낸사람" size="10"><br />

			<input type="text" class="text" name="etc" value="배송시 유의사항" size="10"><br />

			<input type="text" class="text" name="company" value="(주)텐바이텐" size="10"><br />

			<input type="text" class="text" name="smsForward" value="true" size="10">
			<input type="text" class="text" name="smsTarget" value="0101114444" size="10"><br />

			<input type="text" class="text" name="fromSido" value="서울시" size="10">
			<input type="text" class="text" name="fromGugun" value="종로구" size="10">
			<input type="text" class="text" name="fromDong" value="대학로12길" size="10">
			<input type="text" class="text" name="fromDetail" value="31 자유빌딩 5층" size="10">
			<input type="text" class="text" name="fromAddressType" value="NEW" size="10"><br />

			<input type="text" class="text" name="toSido" value="서울시" size="10">
			<input type="text" class="text" name="toGugun" value="동작구" size="10">
			<input type="text" class="text" name="toDong" value="상도3동" size="10">
			<input type="text" class="text" name="toDetail" value="279-508 대원빌라 201호" size="10">
			<input type="text" class="text" name="toAddressType" value="OLD" size="10"><br />

			<input type="text" class="text" name="smallCount" value="1" size="10">
			<input type="text" class="text" name="mediumCount" value="0" size="10">
			<input type="text" class="text" name="bigCount" value="0" size="10">
			<input type="text" class="text" name="complexCount" value="0" size="10"><br />

			<input type="text" class="text" name="billType" value="CREDIT" size="10"><br />

			<input type="text" class="text" name="reservation" value="true" size="10">
			<input type="text" class="text" name="reservationTime" value="2015-06-11 11:30" size="10"><br />

			<input type="text" class="text" name="options" value="" size="10">
		</form>
	</body>
</html>
