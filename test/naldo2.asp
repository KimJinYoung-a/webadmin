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
		<script type="text/javascript" src="/test/swagger-ui/shred.bundle_XXX.js"></script>
		<script type="text/javascript" src="/cscenter/js/jquery-1.8.3.js"></script>
		<script type="text/javascript" src="/cscenter/js/jquery-ui-1.9.2.min.js"></script>
		<script type="text/javascript" src="/test/swagger-ui/handlebars-1.0.0_XXX.js"></script>
		<script type="text/javascript" src="/test/swagger-ui/underscore-min_XXX.js"></script>
		<script type="text/javascript" src="/test/swagger-ui/backbone-min_XXX.js"></script>
		<script type="text/javascript" src="/test/swagger-ui/swagger.js"></script>
		<script type="text/javascript" src="/test/swagger-ui/swagger-ui_XXX.js"></script>
		<script type="text/javascript">

		$(function () {
			/*
			var apiUrl = "http://1.234.83.82:8080/open/api/authenticate?username=user&password=user";
			var authToken = null;

			if(typeof(Storage) !== "undefined") {
				if (localStorage.getItem("ls.token")) {
					alert(localStorage.getItem("ls.token"));
				} else {
					// aaaaaaaaaa
					// alert("먼저 로그인하세요.");
				}
			} else {
				alert("API 사용불가!!\n\n다른 브라우저를 이용하세요.");
			}
			*/

			// alert(window.authorizations);

			// var authToken = JSON.parse(localStorage.getItem("ls.token")).token;

			/*
			window.swaggerUi = new SwaggerUi({
				url: apiUrl,
				dom_id: "swagger-ui-container",
				supportedSubmitMethods: ['post'],
				onComplete: function (swaggerApi, swaggerUi) {
					alert("11");
				},
				onFailure: function (data) {
					alert("22" + data);
				},
				docExpansion: "none"
			});

			window.swaggerUi.load();
			*/



			/*
			xhr.onloadend = function () {
				// alert("kk");

				if (xhr.readyState == 4 && xhr.status == 200) {
					alert(xhr.responseText);
				} else {
					alert("111 " + this.status);
				}
			}
			 */
		});

		function jsCheckStorageAvailable() {
			if(typeof(Storage) !== "undefined") {
				return true;
			}
			return false;
		}

		function jsCheckLogin() {
			if(typeof(Storage) !== "undefined") {
				if (localStorage.getItem("ls.token") != "") {
					return true;
				}
			}
			return false;
		}

		function jsSubmitLogin() {
			if (jsCheckStorageAvailable() != true) {
				alert("API 사용불가!!\n\n다른 브라우저를 이용하세요.");
				return;
			}

			// 메뉴 > 도구 > 인터넷옵션 > 보안탭 > 사용자 지정 수준 > 기타-도메인간의 데이터 원본 엑세스 : 사용
			var apiUrl = "http://1.234.83.82:8080/open/api/authenticate?username=user&password=user";

			var xhr = new XMLHttpRequest();
			xhr.onreadystatechange = doLogin;
			xhr.open("POST", apiUrl, true);
			xhr.setRequestHeader('Accept', 'application/json');
			xhr.send("");

			function doLogin() {
				if (xhr.readyState == 4 && xhr.status == 200) {
					alert("로그인 되었습니다.");
					localStorage.setItem("ls.token", xhr.responseText);
					// alert(xhr.responseText);
				}
			}
		}

		function jsSubmitOrder() {
			var frm = document.frm;

			if (jsCheckStorageAvailable() != true) {
				alert("API 사용불가!!\n\n다른 브라우저를 이용하세요.");
				return;
			}

			if (jsCheckLogin() != true) {
				alert("먼저 로그인하세요.");
				return;
			}

			var apiUrl = "http://1.234.83.82:8080/open/rest/order";

			var authToken = JSON.parse(localStorage.getItem("ls.token")).token;
            // window.authorizations.add("key", new ApiKeyAuthorization("X-Auth-Token", authToken, "header"));

			var data = {};
			for (var i = 0, len = frm.length; i < len; ++i) {
				var input = frm[i];
				if (input.name) {
					data[input.name] = input.value;
				}
			}

			var xhr = new XMLHttpRequest();
			xhr.onreadystatechange = doOrder;
			xhr.open("POST", apiUrl, true);
			xhr.setRequestHeader('Content-Type', 'application/json');
			xhr.setRequestHeader('Accept', 'application/json');
			xhr.setRequestHeader('X-Auth-Token', authToken);
			xhr.send(JSON.stringify(data));

			function doOrder() {
				if (xhr.readyState == 4) {
					if (xhr.status == 200) {
						alert(xhr.responseText);
					} else {
						alert("222 " + xhr.status);
						alert(xhr.responseText);
					}
				}
			}
		}


		</script>
	</head>
	<body>
		<input type="button" class="button" value=" 로그인 " onClick="jsSubmitLogin();">
		<input type="button" class="button" value="주문전송" onClick="jsSubmitOrder();">
		<form name="frm" method="get">
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
		<div id="swagger-ui-container" class="swagger-ui-wrap"></div>

		<script type="text/javascript">

		</script>

	</body>
</html>
