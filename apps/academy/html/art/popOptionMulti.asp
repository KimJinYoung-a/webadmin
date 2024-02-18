<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function() {
	//push setting
	$(".btnSetting button").click(function(){
		$(this).toggleClass('settingOn');
	});

	//option tab control
	function optSize() {
		var optLength = $('.optionTab li:visible').length;
		$('.optionTab li').each(function(){
			if (optLength == 1) {
				$(this).children('button').hide();
				$('.btnPlus').show();
			} else if (optLength == 2) {
				$(this).children('button').show();
				$('.btnPlus').show();
			} else if (optLength == 3) {
				$(this).children('button').show();
				$('.btnPlus').hide();
			}
		});
	}
	optSize();

	$('.optionTab li').click(function(){
		$('.optionTab li').removeClass('current');
		$(this).addClass('current');
	});

	$('.optionTab li button').click(function(e){
		e.preventDefault();
		$(this).parent('li').hide();
		optSize();
	});
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">이중 옵션 설정</h1>
			<ul class="list">
				<li class="" onclick="#">
					<dfn><b>옵션 1</b></dfn>
					<div class="listButton btnCtgySet"><span class="">색상</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
				</li>
				<li class="" onclick="#">
					<dfn><b>옵션 2</b></dfn>
					<div class="listButton btnCtgySet"><span class="setContView">색상</span></div>
				</li>
				<li class="" onclick="#">
					<dfn><b>옵션 3</b></dfn>
					<div class="listButton btnCtgySet"><span class="">색상</span></div>
				</li>
			</ul>
			<div class="registUnit optSet tMar2r">
				<h2><b>옵션별 수량 설정</b></h2>
				<ul class="list">
					<li class="">
						<dfn><em>Z110</em><b>일반, 레드, 스몰</b></dfn>
						<div><input type="number" placeholder="0" value="500" /></div>
						<div style="width:1.5rem">개</div>
						<div class="lPad3r">
							<span class="btnSetting">
								<label>옵션 사용여부</label>
								<button type="button" onclick="">옵션 사용여부 설정</button>
							</span>
						</div>
					</li>
					<li class="">
						<dfn><em>Z110</em><b>일반, 레드, 스몰</b></dfn>
						<div><input type="number" placeholder="0" /></div>
						<div style="width:1.5rem">개</div>
						<div class="lPad3r">
							<span class="btnSetting">
								<label>옵션 사용여부</label>
								<button type="button" onclick="">옵션 사용여부 설정</button>
							</span>
						</div>
					</li>
					<li class="">
						<dfn><em>Z110</em><b>일반, 레드, 스몰</b></dfn>
						<div><input type="number" placeholder="0" /></div>
						<div style="width:1.5rem">개</div>
						<div class="lPad3r">
							<span class="btnSetting">
								<label>옵션 사용여부</label>
								<button type="button" onclick="">옵션 사용여부 설정</button>
							</span>
						</div>
					</li>
				</ul>
			</div>
		</div>
		<!--// content -->
		<!-- 하단 플로팅 버튼 -->
		<div class="floatingBar">
			<p><button type="button" class="btnV16a btnWishV16a">초기화</button></p>
		</div>

		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>