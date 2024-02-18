<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry"><!-- for dev msg : bgGry 클래스 추가 (11/24)--->
			<h1 class="hidden">가격 변경 요청</h1>
			<div class="priceChange">
				<ul class="list">
					<li class="">
						<dfn><b>공급 마진</b></dfn>
						<div class="aftPrice">100 %</div>
					</li>
					<li class="critical">
						<dfn><b>판매가</b></dfn>
						<div class="prePrice">100,000 원</div>
						<div class="chgArw"></div>
						<div class="aftPrice"><input type="number" value="" placeholder="100,000" /></div>
						<div style="width:1.6rem">원</div>
					</li>
					<li class="">
						<dfn><b>공급가 <span class="fs1-1r">(부가세 포함)</span></b></dfn>
						<div class="prePrice">10,000 원</div>
						<div class="chgArw"></div>
						<div class="aftPrice"><input type="number" value="" placeholder="10,000" /></div>
						<div style="width:1.6rem">원</div>
					</li>
				</ul>
				<!-- for dev msg : 변경사유 입력폼 추가 (11/24)--->
				<div class="linkInsert tMar2r">
					<textarea rows="8" placeholder="변경사유를 입력해주세요."></textarea>
				</div>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>