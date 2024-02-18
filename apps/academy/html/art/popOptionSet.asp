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
			<h1 class="hidden">옵션 설정</h1>
			<div class="optionSetting">
				<div class="optSetListWrap">
					<div class="setList">
						<dl class="optionUnit">
							<dt>
								<div><input type="text" placeholder="옵션이름" style="width:80%;" /></div>
								<div class="btnSetting">
									<label>옵션 사용여부</label>
									<button type="button" onclick="">옵션 사용여부 설정</button>
								</div>
							</dt>
							<dd>
								<ul>
									<li>
										<div>
											<span><input type="text" placeholder="옵션 내 항목" style="width:100%;" /></span>
											<span class="btnSetting">
												<label>옵션 사용여부</label>
												<button type="button" onclick="">옵션 사용여부 설정</button>
											</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" placeholder="수량" style="width:100%;" /></span>
											<span>개</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" placeholder="추가금액" style="width:100%;" /></span>
											<span>원</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" placeholder="공급가" style="width:100%;" /></span>
											<span>원</span>
										</div>
									</li>
								</ul>
							</dd>
							<dd>
								<ul>
									<li>
										<div>
											<span><input type="text" placeholder="옵션 내 항목" style="width:100%;" /></span>
											<span class="btnSetting">
												<label>옵션 사용여부</label>
												<button type="button" onclick="">옵션 사용여부 설정</button>
											</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" placeholder="수량" style="width:100%;" /></span>
											<span>개</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" placeholder="추가금액" style="width:100%;" /></span>
											<span>원</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" placeholder="공급가" style="width:100%;" /></span>
											<span>원</span>
										</div>
									</li>
								</ul>
							</dd>
						</dl>
					</div>
					<div class="addBtn">
						<button type="button" class="btnB1 btnDkGry" disabled="disabled"><span class="itemAdd">추가</span></button><!-- for dev msg : 추가버튼 클릭시 setList division의 optionUnit dd가 추가(최대 9개까지)되면 됩니다.-->
					</div>
				</div>
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