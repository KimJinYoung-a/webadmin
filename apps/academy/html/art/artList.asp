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
			<h1 class="hidden">작품 관리</h1>
			<div class="artManage">
				<ul class="listTab">
					<li class="current" onclick=""><div>판매중</div></li>
					<li onclick=""><div>판매종료</div></li>
					<li onclick=""><div>등록대기</div></li>
				</ul>
				<!-- for dev msg : 작품 리스트가 없을때 노출됩니다. -->
				<div class="artNo" style="display:none;">
					<div class="linkNotice">
						<p class="fs1-5r">오른쪽 상단 버튼을 선택해 <br />작품을 등록해주세요!</p>
						<!--
						<p class="fs1-5r">판매종료된 작품이 없습니다.</p>
						<p class="fs1-5r">등록대기중인 작품이 없습니다.</p>
						-->
					</div>
				</div>
				<!-- for dev msg : 작품 리스트가 없을때 노출됩니다. //-->
				<div class="artListCont">
					<div class="artSearchTop">
						<div class="searchInput">
							<input type="Search" placeholder="작품명, 코드, 키워드 검색" />
							<button type="button" class="btnSearch">검색</button>
							<!-- button type="button" class="btnTextDel">삭제</button -->
						</div>
						<div class="btnFilter filterActive"><!-- for dev msg : 필터링 된 후 클래스 filterActive 붙여주세요 -->
							<button type="button">필터</button>
						</div>
					</div>

					<div class="artListWrap">
						<ul class="artList">
							<li class="artFlag1"><!-- 판매중(↓ 상태표시에 따라 클래스 artFlag1 ~ artFlag8 붙습니다) //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus"><strong>판매중</strong></span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21 <i class="tag1">품절임박</i></p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag2"><!-- 일시품절 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">일시품절</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag3"><!-- 품절 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">품절</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag4"><!-- 사용안함 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">사용안함</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb whiteout"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag5"><!-- 반려 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">반려</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag6"><!-- 보류 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">보류</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag7"><!-- 등록대기 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">등록대기</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
							<li class="artFlag8"><!-- 임시저장 //-->
								<a href="">
									<div class="artStatus">
										<p><span>2016.03.21</span><span>ㅣ</span><span>A91006701777</span></p>
										<p class="rt"><span class="nowStatus">임시저장</span></p>
									</div>
									<div class="artInfo">
										<div class="artThumb"><img src="http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007935-1.jpg" alt="" /></div>
										<strong>플로리다 드림캐처 작품명 최대 한줄로 표현됩니다</strong>
										<div class="artTxt">
											<p><dfn>재고</dfn> 21</p>
											<p class="tPad1r"><span>30,000원</span><span class="saleRate">20%</span></p>
										</div>
									</div>
								</a>
							</li>
						</ul>
						<div class="paging">
							<a href="" class="btnPrev">이전 페이지</a>
							<span><input type="number" class="pageNum" value="1" /> / 30</span>
							<a href="" class="btnNext">다음 페이지</a>
						</div>
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