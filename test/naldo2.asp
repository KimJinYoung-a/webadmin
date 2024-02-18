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
					// alert("���� �α����ϼ���.");
				}
			} else {
				alert("API ���Ұ�!!\n\n�ٸ� �������� �̿��ϼ���.");
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
				alert("API ���Ұ�!!\n\n�ٸ� �������� �̿��ϼ���.");
				return;
			}

			// �޴� > ���� > ���ͳݿɼ� > ������ > ����� ���� ���� > ��Ÿ-�����ΰ��� ������ ���� ������ : ���
			var apiUrl = "http://1.234.83.82:8080/open/api/authenticate?username=user&password=user";

			var xhr = new XMLHttpRequest();
			xhr.onreadystatechange = doLogin;
			xhr.open("POST", apiUrl, true);
			xhr.setRequestHeader('Accept', 'application/json');
			xhr.send("");

			function doLogin() {
				if (xhr.readyState == 4 && xhr.status == 200) {
					alert("�α��� �Ǿ����ϴ�.");
					localStorage.setItem("ls.token", xhr.responseText);
					// alert(xhr.responseText);
				}
			}
		}

		function jsSubmitOrder() {
			var frm = document.frm;

			if (jsCheckStorageAvailable() != true) {
				alert("API ���Ұ�!!\n\n�ٸ� �������� �̿��ϼ���.");
				return;
			}

			if (jsCheckLogin() != true) {
				alert("���� �α����ϼ���.");
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
		<input type="button" class="button" value=" �α��� " onClick="jsSubmitLogin();">
		<input type="button" class="button" value="�ֹ�����" onClick="jsSubmitOrder();">
		<form name="frm" method="get">
			<input type="text" class="text" name="orderPhoneNumber" value="0101111111" size="10">
			<input type="text" class="text" name="senderPhoneNumber" value="0101111111" size="10">
			<input type="text" class="text" name="receiverPhoneNumber" value="0101111111" size="10"><br />

			<input type="text" class="text" name="receiverName" value="�޴»��" size="10">
			<input type="text" class="text" name="senderName" value="�������" size="10"><br />

			<input type="text" class="text" name="etc" value="��۽� ���ǻ���" size="10"><br />

			<input type="text" class="text" name="company" value="(��)�ٹ�����" size="10"><br />

			<input type="text" class="text" name="smsForward" value="true" size="10">
			<input type="text" class="text" name="smsTarget" value="0101114444" size="10"><br />

			<input type="text" class="text" name="fromSido" value="�����" size="10">
			<input type="text" class="text" name="fromGugun" value="���α�" size="10">
			<input type="text" class="text" name="fromDong" value="���з�12��" size="10">
			<input type="text" class="text" name="fromDetail" value="31 �������� 5��" size="10">
			<input type="text" class="text" name="fromAddressType" value="NEW" size="10"><br />

			<input type="text" class="text" name="toSido" value="�����" size="10">
			<input type="text" class="text" name="toGugun" value="���۱�" size="10">
			<input type="text" class="text" name="toDong" value="��3��" size="10">
			<input type="text" class="text" name="toDetail" value="279-508 ������� 201ȣ" size="10">
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
