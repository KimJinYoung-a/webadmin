<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	// main banner
	var swiper01 = new Swiper(".basicImgRegist .swiper-container", {
		pagination:false,
		slidesPerView:'auto',
		spaceBetween:5
	});

	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

	// textarea auto size
	$("textarea.autosize").keyup(function () {
		$(this).css("height","1.96rem").css("height",($(this).prop("scrollHeight"))+"px");
	});
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">작품 등록</h1>
			<div class="artRegist">
				<div class="registUnit disabled"><!-- for dev msg : 비활성화 시 class : disabled 붙여주세요 -->
					<div class="basicImgRegist">
						<div class="swiper-container">
							<div class="swiper-wrapper">
								<div class="swiper-slide"><button type="button">이미지 등록1</button></div>
								<div class="swiper-slide"><button type="button">이미지 등록2</button></div>
								<div class="swiper-slide"><button type="button">이미지 등록3</button></div>
								<div class="swiper-slide"><button type="button">이미지 등록4</button></div>
							</div>
						</div>
					</div>
					<ul class="list">
						<li class="critical" onclick="#">
							<dfn><b>카테고리 설정</b></dfn>
							<div class="listButton btnCtgySet"><span class="setContView">홈/데코 외 1건</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="critical">
							<dfn><b>상품명</b></dfn>
							<div><input type="text" placeholder="22자 이하로 입력해주세요" /></div>
						</li>
						<li class="selectBtn">
							<div class="grid2"><button type="button" class="btnM1 btnGry selected">일반 상품</button></div>
							<div class="grid2"><button type="button" class="btnM1 btnGry">주문제작 상품</button></div>
						</li>
					</ul>
				</div>

				<!-- for dev msg : 주문제작 상품 선택시 노출됩니다. -->
				<div class="registUnit orderArt">
					<h2 class="critical"><b>주문제작 설정</b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid2"><button type="button" class="btnM1 btnGry">즉시 발송</button></div>
							<div class="grid2"><button type="button" class="btnM1 btnGry">제작 후 발송</button></div>
						</li>
						<li class="critical" onclick="#">
							<dfn><b>제작 기간</b></dfn>
							<div><input type="number" value="" placeholder="100" /></div>
							<div style="width:1.6rem">일</div>
						</li>
						<li class="" onclick="#">
							<dfn><b>특이사항</b></dfn>
							<div class="listButton btnCtgySet"><span class="">주문 제작 시 총 100일이 걸리지 말입니다.</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="selectBtn">
							<div class="grid2"><button type="button" class="btnM1 btnGry ckBtn">제작 메시지 필요</button></div>
							<div class="grid2"><button type="button" class="btnM1 btnGry ckBtn">제작 이미지 필요</button></div>
						</li>
						<li class="critical" onclick="#"><!-- for dev msg : 제작 이미지 필요 선택시 노출됩니다. -->
							<dfn><b>이미지 수신 메일</b></dfn>
							<div><input type="email" value="" placeholder="id1234@example.com" /></div>
						</li>
					</ul>
				</div>
				<!--// for dev msg : 주문제작 상품 선택시 노출됩니다. -->

				<div class="registUnit basicInfo">
					<h2>기본 정보</h2>
					<ul class="list">
						<li class="critical">
							<dfn><b>제작자</b></dfn>
							<div><input type="text" placeholder="작가명/법인을 입력해주세요" /></div>
						</li>
						<li class="critical">
							<dfn><b>원산지</b></dfn>
							<div><input type="text" placeholder="국가명을 입력해주세요" /></div>
						</li>
						<li class="critical">
							<dfn><b>재질</b></dfn>
							<div><input type="text" placeholder="예) 플라스틱, 합금, 은" /></div>
						</li>
						<li class="critical">
							<dfn><b>크기</b></dfn>
							<div><input type="number" value="" placeholder="예) 7.5 * 7.5" /></div>
							<div style="width:2.4rem">cm</div>
						</li>
						<li class="critical">
							<dfn><b>무게</b></dfn>
							<div><input type="number" placeholder="예) 785" /></div>
							<div style="width:1.4rem">g</div>
						</li>
						<li class="" onclick="#">
							<dfn><b>검색 키워드</b></dfn>
							<div class="listButton btnCtgySet"><span class="">7건 등록</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="registUnit salePrice">
					<h2 class="critical"><b>판매 가격 <span>(부가세 포함)</span></b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid2"><button type="button" class="btnM1 btnGry">과세</button></div>
							<div class="grid2"><button type="button" class="btnM1 btnGry">면세</button></div>
						</li>
						<li>
							<dfn><b>공급 마진</b></dfn>
							<div><input type="number" value="" placeholder="100" /></div>
							<div style="width:1.8rem">%</div>
						</li>
						<li class="critical">
							<dfn><b>판매가</b></dfn>
							<div><input type="number" placeholder="판매가(소비자가)를 입력해주세요" /></div>
						</li>
						<li>
							<dfn><b>공급가</b></dfn>
							<div><input type="number" value="" placeholder="0" /></div>
						</li>
					</ul>
				</div>
				<div class="registUnit quantity">
					<h2 class="critical"><b>수량 설정</b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid2"><button type="button" class="btnM1 btnGry">한정 수량</button></div>
							<div class="grid2"><button type="button" class="btnM1 btnGry">무제한</button></div>
						</li>
						<li><!--for dev msg : 한정수량 선택시 노출됩니다. -->
							<dfn><b>수량</b></dfn>
							<div><input type="number" value="1" placeholder="수량을 입력해주세요" /></div>
							<div style="width:1.6rem">개</div>
						</li>
					</ul>
				</div>
				<div class="registUnit option">
					<h2 class="critical"><b>옵션 설정</b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid3"><button type="button" class="btnM1 btnGry">사용안함</button></div>
							<div class="grid3"><button type="button" class="btnM1 btnGry">단일 옵션</button></div>
							<div class="grid3"><button type="button" class="btnM1 btnGry">이중 옵션</button></div>
						</li>
						<li class="critical" onclick="#"><!--for dev msg : 단일 옵션 or 이중 옵션 선택시 노출됩니다. -->
							<dfn><b>옵션 설정</b></dfn>
							<div class="listButton btnCtgySet"><span class="">설정됨</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="registUnit delivery">
					<h2 class="critical"><b>배송 설정 <span>(부가세 포함)</span></b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid4"><button type="button" class="btnM1 btnGry">무료</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry">조건부</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry">선불</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry">착불</button></div>
						</li>
						<li class="critical" onclick="#">
							<dfn><b>배송비 안내</b></dfn>
							<div class="listButton btnCtgySet"><span class="setContView">제품은 배송시 안전을 위해 배송비가 부과됩니다. 제품은 배송시 안전을 위해 배송비가 부과됩니다.</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="registUnit law">
					<h2 class="critical"><b>관련법 필수 입력 항목</b></h2>
					<ul class="list">
						<li class="critical" onclick="#">
							<dfn><b>상품정보제공고시</b></dfn>
							<div class="listButton btnCtgySet"><span class="">의류</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="critical" onclick="#">
							<dfn><b>안전인증대상</b></dfn>
							<div class="listButton btnCtgySet"><span class="">국가통합인증(KC마크)</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="detail">
					<div class="registUnit">
						<h2 class="critical"><b>상세 정보</b></h2>
						<ul class="list">
							<li>
								<p><button type="button" class="btnImgRegist">이미지 등록</button></p>
								<p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize"></textarea></p>
							</li>
							<li>
								<p><button type="button" class="btnImgRegist">이미지 등록</button></p>
								<p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize"></textarea></p>
							</li>
							<li>
								<p><img src="http://image.thefingers.co.kr/diyitem/contentsimage/00/M000002904_01.jpg" alt="" /></p>
								<p class="tMar1-5r">
									<textarea placeholder="내용을 입력해주세요" class="autosize">
수염틸란드시아 입니다
먼지먹는 식물로 이미 유명한 틸란드시아 입니다
									</textarea></p>
							</li>
						</ul>
					</div>
					<div class="addBtn">
						<button type="button" class="btnB1 btnDkGry"><span class="itemAdd">추가</span></button>
						<p class="tPad2r">최대 15개까지 등록 가능합니다.</p>
					</div>
				</div>
			</div>
		</div>
		<!--// content -->

		<!-- 알림 메세지 -->
		<div class="attentionBar" style="display:none">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_dot.png" alt="필수표시" style="width:0.4rem; height:0.4rem; margin:0.3rem 0.3rem 0 0" /> 표기는 필수 선택/입력 항목입니다. 꼭 입력해주세요.</p>
		</div>

		<div class="attentionBar" style="display:none">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_save.png" alt="저장표시" style="width:1.2rem; height:1.2rem; margin:0.3rem 0.3rem 0 0" /> 2016.10.10 – 14:20에 저장되었습니다.</p>
		</div>

		<div class="attentionBar" style="display:none">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_time.png" alt="시계표시" style="width:1.2rem; height:1.2rem; margin:0.3rem 0.3rem 0 0" /> 등록 대기중인 작품입니다. 관리자 승인 후 사이트에 게시됩니다.</p>
		</div>

		<div class="attentionBar badNotice" style="display:none">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_notice.png" alt="경고표시" style="width:1.2rem; height:1.1rem; margin:0.3rem 0.3rem 0 0;" /> 상품명에 굉장히 불건전하고 좋지 못한 단어가 있습니다. <br />꼭 수정해주시기 바랍니다.</p>
		</div>

		<div class="attentionBar badNotice" style="display:none">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_notice.png" alt="경고표시" style="width:1.2rem; height:1.1rem; margin:0.3rem 0.3rem 0 0;" /> 좋지 못한 사유로 인해 반려되었습니다. 반려된 작품은 다시 등록될 수 없습니다.</p>
		</div>

		<!-- 하단 플로팅 버튼 -->
		<div class="floatingBar">
			<p><button type="button" class="btnV16a btnWishV16a">임시저장</button></p>
			<p><button type="button" class="btnV16a btnRed2V16a">미리보기</button></p>
		</div>

		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>