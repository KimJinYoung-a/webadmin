<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

	// textarea auto size
	$(".searchInput input").keyup(function () {
		$(this).parent().find('button').fadeIn();
	});

	// search box hidden scroll top auto change
	var schH = $(".artSearchTop").outerHeight();
	var tabT = $(".listTab").offset().top;
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH-tabT}, 'fast');
	}, 300);
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">작품 정보</h1>
			<div class="artDetailInfo">
				<ul class="listTab">
					<li class="current" onclick=""><div>기본 정보</div></li>
					<li onclick=""><div>수정</div></li>
				</ul>

				<div class="artDetailWrap">
					<div class="artDetail">
						<ul class="artList">
							<li class="artFlag1"><!-- 판매중(↓ 상태표시에 따라 클래스 artFlag1 ~ artFlag8 붙습니다) //-->
								<a href="">
									<div class="artStatus">
										<p><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus"><strong>판매중</strong></span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 작품 정보에선 모두 노출됩니다</strong>
										<div class="artTxt">
											<p><span>업체배송</span><span class="sepLine">l</span><span>주문제작 상품</span></p>
										</div>
									</div>
								</a>
							</li>
						</ul>
						<dl class="dfCompos">
							<dt>가격 정보 
								<i class="tag1">변경 대기중</i>
								<!-- <i class="tag2">요청 반려</i> -->
							</dt>
							<dd>
								<ul class="list">
									<li class="">
										<dfn><b>소비자가</b></dfn>
										<div><span class=""><s>30,000원</s></span><span class="lPad0-8r">27,000원</span><strong class="lPad0-8r cOr1">10%</strong></div>
									</li>
									<li class="">
										<dfn><b>공급가</b></dfn>
										<div class="cGy2"><span>2,000원</span><span class="sepLine">l</span><span>업체</span><span class="sepLine">l</span><span>마진 100%</span></div>
									</li>
								</ul>
							</dd>
							<dd class="tPad2r disabled"><!-- for dev msg : 반려 사유 노출시 비활성화 class : disabled 붙여주세요 -->
								<div class="boxUnit bdrTGry">
									<div class="boxHead">
										<b>가격 변경 요청</b>
										<p><span>90,000원</span><i class="chgArw"></i><span><strong class="cOr1">100,000원</strong></span></p>
									</div>
									<div class="boxCont">생산가가 높아져 변경을 요청드립니다. 빠른 처리 부탁드립니다.</div>
								</div>
							</dd>
							<dd class="tPad2r">
								<div class="boxUnit bdrTOr">
									<div class="boxHead">
										<b>반려 사유</b>
									</div>
									<div class="boxCont">현재 가격 정책 재조정중입니다. 재조정중에는 가격요청을 진행할 수 없습니다.</div>
								</div>
							</dd>
							<dd class="selectBtn tMar2-5r">
								<div><button type="button" class="btnM1 btnGry selected">가격 변경 요청</button></div>
								<!-- 가격변경 대기중일때 노출
								<div class="grid2"><button type="button" class="btnM1 btnGry selected">가격 재변경 요청</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry">변경 요청 취소</button></div>
								//-->
								<!-- 반려 사유 노출됬을때 
								<div class="grid2"><button type="button" class="btnM1 btnGry selected">반려 확인</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry">변경 요청 취소</button></div>
								//-->
							</dd>
						</dl>
						<dl class="dfCompos">
							<dt>재고 현황</dt>
							<dd>
								<ul class="list">
									<li class="cGy3">
										<dfn class="cGy1"><b>Z110</b></dfn>
										<div class="optName"><div>옵션없음</div></div>
										<div class="rt">20개</div>
									</li>
									<li class="cGy3">
										<dfn class="cGy1"><b>Z120</b></dfn>
										<div class="optName"><div>레드, Large</div></div>
										<div class="rt">20개</div>
									</li>
									<li class="cGy3">
										<dfn class="cGy1"><b>Z130</b></dfn>
										<div class="optName"><div>옵션명 나열됩니다. 글자수 많으면 점점점 처리되요</div></div>
										<div class="rt">2개</div>
									</li>
								</ul>
							</dd>
						</dl>
						<dl class="dfCompos">
							<dt>판매 관리</dt>
							<dd class="selectBtn">
								<div class="grid3"><button type="button" class="btnM1 btnGry selected">판매</button></div>
								<div class="grid3"><button type="button" class="btnM1 btnGry">일시품절</button></div>
								<div class="grid3"><button type="button" class="btnM1 btnGry">품절</button></div>
							</dd>
						</dl>
						<dl class="dfCompos">
							<dt>사용 여부</dt>
							<dd class="selectBtn">
								<div class="grid2"><button type="button" class="btnM1 btnGry selected">사용</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry">사용안함</button></div>
							</dd>
						</dl>
					</div>
				</div>





			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>