<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	// select box control
	$('#ctgy1').on('change', function () {
		$('#ctgy2').removeAttr('disabled');
		$('.addBtn button').removeAttr('disabled');
		$('.addBtn button').addClass('active');
	});
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">카테고리 설정</h1>
			<div class="ctgySetting">
				<div class="selectBtn">
					<div class="grid2">
						<select id="ctgy1">
							<option>대분류</option>
							<option>홈/데코</option>
						</select>
					</div>
					<div class="grid2">
						<select id="ctgy2" disabled="disabled">
							<option>중분류</option>
							<option>인테리어 소품</option>
						</select>
					</div>
				</div>
				<div class="addBtn">
					<button type="button" class="btnB1 btnDkGry" disabled="disabled"><span class="itemAdd">추가</span></button>
				</div>
				<div class="addCtgyList">
					<ul>
						<li><div><span>홈/데코</span> <button type="button" class="btnListDel">삭제</button></div></li>
						<li><div><span>홈/데코</span> <span>인테리어 소품</span> <button type="button" class="btnListDel">삭제</button></div></li>
					</ul>
				</div>
				<!-- for dev msg : 카테고리 리스트 쌓일때는 아래 내용은 안보여집니다. -->
				<div class="linkNotice">
					<p class="fs1-5r">카테고리 추가 후, 확인 버튼을 눌러주세요</p>
					<p class="tMar1-5r">추가된 카테고리는 더핑거스 웹사이트의 <br />해당 카테고리 리스트에 노출됩니다.</p>
				</div>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>