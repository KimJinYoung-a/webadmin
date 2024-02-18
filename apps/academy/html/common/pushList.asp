<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	//sorting control
	$(".pushSort button").click(function(){
		if($(".pushSort ul").is(":hidden")){
			$(this).parent().children('ul').show();
			$(this).addClass("active");
		}else{
			$(this).parent().children('ul').hide();
			$(this).removeClass("active");
		};
	});

	$(".pushSort li a").click(function(e){
		e.preventDefault()
		var selectTxt = $(this).text();
		$(this).parents('.pushSort').children('button').text(selectTxt);
		$(".pushSort ul").hide();
		$(this).parents('.pushSort').children('button').removeClass("active");
	});

	//push setting
	$(".btnPushSet button").click(function(){
		$(this).toggleClass('settingOn');
	});
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">알림 메시지</h1>
			<div class="pushMsg">
				<div class="pushHead">
					<div class="pushSort">
						<button type="button" onclick="" class="">전체 알림</button>
						<ul>
							<li><a href="">전체 알림</a></li>
							<li><a href="">공지사항</a></li>
							<li><a href="">등록승인</a></li>
						</ul>
					</div>
					<div class="btnPushSet">
						<label>알림</label>
						<button type="button" onclick="">알림 설정</button>
					</div>
				</div>
				<ul class="pushList" style="display:none">
					<li class="flagNoti"><!-- for dev msg : 공지사항일경우 flagNoti / 등록승인일경우(기본일경우?) flagAprv class 명 붙여주세요 -->
						<a href="">
							<dfn>공지사항</dfn>
							<p>판매자 신규 약관 동의가 필요합니다.</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagAprv">
						<a href="">
							<dfn>등록승인</dfn>
							<p>등록승인 되었습니다. 아로마 디퓨저 세트와 맥모닝 세트를 먹나요?</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagAprv">
						<a href="">
							<dfn>등록승인</dfn>
							<p>등록승인 되었습니다. 아로마 디퓨저 세트와 맥모닝 세트를 먹나요?</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagAprv">
						<a href="">
							<dfn>등록승인</dfn>
							<p>등록승인 되었습니다. 아로마 디퓨저 세트와 맥모닝 세트를 먹나요?</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagNoti">
						<a href="">
							<dfn>공지사항</dfn>
							<p>판매자 신규 약관 동의가 필요합니다.</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagAprv">
						<a href="">
							<dfn>등록승인</dfn>
							<p>등록승인 되었습니다. 아로마 디퓨저 세트와 맥모닝 세트를 먹나요?</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagAprv">
						<a href="">
							<dfn>등록승인</dfn>
							<p>등록승인 되었습니다. 아로마 디퓨저 세트와 맥모닝 세트를 먹나요?</p>
							<span>2016.12.06</span>
						</a>
					</li>
					<li class="flagAprv">
						<a href="">
							<dfn>등록승인</dfn>
							<p>등록승인 되었습니다. 아로마 디퓨저 세트와 맥모닝 세트를 먹나요?</p>
							<span>2016.12.06</span>
						</a>
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