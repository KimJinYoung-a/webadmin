<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">안전인증 대상</h1>
			<div class="artSafeSet">
				<div class="selectBtn">
					<div class="grid2"><button type="button" class="btnM1 btnGry">대상</button></div>
					<div class="grid2"><button type="button" class="btnM1 btnGry">대상 아님</button></div>
				</div>
				<div class="safeCertify"><!-- for dev msg : 안전인증 대상 선택시 노출됩니다.-->
					<ul class="list">
						<li class="selectBtn">
							<select name="safetyDiv" class="select">
								<option value="">안전인증구분을 선택해주세요</option>
								<option value="10">국가통합인증(KC마크)</option>
								<option value="20">전기용품 안전인증</option>
								<option value="30">KPS 안전인증 표시</option>
								<option value="40">KPS 자율안전 확인 표시</option>
								<option value="50">KPS 어린이 보호포장 표시</option>
							</select>
						</li>
						<li>
							<dfn><b>인증번호</b></dfn>
							<div><input type="number" placeholder="인증번호를 입력해주세요" /></div>
						</li>
					</ul>
					<div class="optionUnit fs1-2r cGy1 rt">※ 유아용품이나 전기용품일 경우 필수 입력</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>