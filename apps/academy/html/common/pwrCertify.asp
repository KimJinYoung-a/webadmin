<!-- #include virtual="/apps/academy/html/inc/head.asp" -->
<script>
$(function(){
	$(".certifyNumList li, .certifyNumList li button").click(function(e){
		e.preventDefault();
		$(".certifyNumList li").removeClass('selected');
		$(this).addClass('selected');
		$(".certifyNumInput").fadeIn();
		window.parent.$('html,body').animate({scrollTop:$('.certifyNumInput').offset().top},200);
	});
});
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<div class="pwrSearch">
				<h1 class="hidden">휴대폰 번호 인증</h1>
				<p class="tit">담당자 정보 확인 후 <br />인증번호받기 버튼을 선택해 주세요.</p>

				<form>
				<div class="certifyNumList">
					<ul>
						<li>
							<div>
								<b>영업담당자</b>
								<p><strong>권나*</strong> / 010-****-1234</p>
							</div>
							<div class="btnCertify"><button class="btnS1 btnWht">인증번호 받기</button></div>
						</li>
						<li>
							<div>
								<b>정산담당자</b>
								<p><strong>마동*</strong> / 010-****-1234</p>
							</div>
							<div class="btnCertify"><button class="btnS1 btnWht">인증번호 받기</button></div>
						</li>
						<li>
							<div>
								<b>배송담당자</b>
								<p><strong>박보*</strong> / 010-****-1234</p>
							</div>
							<div class="btnCertify"><button class="btnS1 btnWht">인증번호 받기</button></div>
						</li>
					</ul>
				</div>
				<div class="certifyNumInput" style="display:none;">
					<div class="textForm2"><label>인증번호 입력</label><input type="number" placeholder="인증번호를 입력해주세요" style="width:75%;" /><span class="timer">02:59</span></div>
					<div class="btnCertify"><button class="btnB1 btnGrn">확 인</button></div>
				</div>
				</form>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>