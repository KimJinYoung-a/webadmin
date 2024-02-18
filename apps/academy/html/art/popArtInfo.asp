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
			<h1 class="hidden">상품정보 제공 고시</h1>
			<div class="artInfoSet">
				<div class="selectBtn">
					<div>
						<select>
							<option value="">품목을 선택해주세요</option>
							<option value="01">의류</option>
							<option value="02">구두/신발</option>
							<option value="03">가방</option>
							<option value="04">패션잡화(모자/벨트/액세서리)</option>
							<option value="05">침구류/커튼</option>
							<option value="06">가구(침대/소파/싱크대/DIY제품)</option>
							<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
							<option value="17">주방용품</option>
							<option value="18">화장품</option>
							<option value="19">귀금속/보석/시계류</option>
							<option value="20">식품(농수산물)</option>
							<option value="21">가공식품</option>
							<option value="22">건강기능식품/체중조절식품</option>
							<option value="23">영유아용품</option>
							<option value="24">악기</option>
							<option value="25">스포츠용품</option>
							<option value="26">서적</option>
							<option value="35">기타</option>
						</select>
					</div>
				</div>
				<ul class="infoList"><!-- for dev msg : 품목 선택에 따라 각 입력항목 노출됩니다.-->
					<li>
						<dl class="infoUnit">
							<dt><strong>1. 제품소재</strong></dt>
							<dd><input type="text" placeholder="~을 입력해주세요 또는 부가 설명" /></dd><!-- for dev msg : 텍스트 인풋박스 일때 -->
						</dl>
					</li>
					<li>
						<dl class="infoUnit">
							<dt><strong>6. 제조자(수입여부)</strong></dt>
							<dd class="sltYN selectBtn"><!-- for dev msg : yes/no 버튼 있는 경우 -->
								<div class="grid2"><button type="button" class="btnM1 btnGry">Yes</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry">No</button></div>
							</dd>
							<dd><input type="text" placeholder="~을 입력해주세요 또는 부가 설명" /></dd>
							<dd class="addition">수입품 아님: N/제조자 표기, 수입품: Y/수입자를 함께 표기(병행수입의 경우 병행수입 여부로 대체 가능)</dd><!-- for dev msg : 추가적 설명 있는 경우 -->
						</dl>
					</li>
					<li>
						<dl class="infoUnit">
							<dt><strong>8. 식품위생법에 따른 수입 기구,용기 여부</strong></dt>
							<dd class="sltYN selectBtn">
								<div class="grid2"><button type="button" class="btnM1 btnGry">Yes</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry">No</button></div>
							</dd>
							<dd><input type="text" placeholder="~을 입력해주세요 또는 부가 설명" /></dd>
						</dl>
					</li>
					<li>
						<dl class="infoUnit">
							<dt><strong>9. 품질보증기준</strong></dt>
							<dd><textarea rows="3" placeholder="~을 입력해주세요 또는 부가 설명"></textarea></dd><!-- for dev msg : textarea 일때 -->
							<dd class="addition">품질보증기간 및 수리/교환/반품 등의 보상방법 정보</dd>
						</dl>
					</li>
					<li>
						<dl class="infoUnit">
							<dt><strong>10. A/S 책임자/전화번호</strong></dt>
							<dd><input type="text" placeholder="~을 입력해주세요 또는 부가 설명" /></dd>
							<dd class="addition">더핑거스 고객행복센터 1644-1557</dd>
						</dl>
					</li>
				</ul>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>