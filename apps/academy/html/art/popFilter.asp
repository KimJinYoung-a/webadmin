<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	// select box control
	$('#ctgy1').on('change', function () {
		$('#ctgy2').removeAttr('disabled');
	});

	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});
});
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">필터</h1>
			<!-- for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
			<div class="filterWrap">
				<dl class="dfCompos">
					<dt>카테고리</dt>
					<dd class="selectBtn">
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
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>판매상태</dt>
					<dd class="selectBtn">
						<div class="grid3"><button type="button" class="btnM1 btnGry selected">전체</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry">판매중</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry">일시품절</button></div>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>대기상태</dt>
					<dd class="selectBtn">
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry selected">전체</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry">임시저장</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry">대기</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry">보류</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry">반려</button></li>
						</ul>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>한정구분</dt>
					<dd class="selectBtn">
						<div class="grid3"><button type="button" class="btnM1 btnGry selected">전체</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry">한정</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry">비한정</button></div>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>정렬기준</dt>
					<dd class="selectBtn">
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry selected"><span class="sort srtUp">등록순</span></button></li><!-- for dev msg : 한번 클릭시 button 에 selected 붙여주시고 한번더 클릭하면 span 태그의 srtUp/srtDown 이 토글되면 됩니다 -->
							<li class="grid3"><button type="button" class="btnM1 btnGry"><span class="sort srtDown">매출순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry"><span class="sort srtDown">판매량순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry"><span class="sort srtDown">관심등록순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry"><span class="sort srtDown">가격순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry"><span class="sort srtDown">할인순</span></button></li>
						</ul>
					</dd>
					<dd class="selectBtn" style="display:none">
						<div class="grid2"><button type="button" class="btnM1 btnGry selected"><span class="sort srtDown">등록순</span></button></div>
						<div class="grid2"><button type="button" class="btnM1 btnGry"><span class="sort srtDown">매출순</span></button></div>
					</dd>
				</dl>
			</div>
			<!-- //for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>