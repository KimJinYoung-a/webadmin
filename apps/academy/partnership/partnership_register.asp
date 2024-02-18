<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 작가 신청"
%>
<!-- #include Virtual="/apps/academy/lib/chkDevice.asp" -->
<!-- #include virtual="/apps/academy/lib/customapp.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!doctype html>
<html lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="format-detection" content="telephone=no" />
<title>2016 Fingers Academy Admin App</title><!-- for dev msg : 각 페이지 타이틀 넣어주세요 -->

<link rel="stylesheet" type="text/css" href="https://m.thefingers.co.kr/lib/css/common.css" />
<link rel="stylesheet" type="text/css" href="https://m.thefingers.co.kr/lib/css/content.css" />
<link rel="stylesheet" type="text/css" href="https://m.thefingers.co.kr/lib/css/myfingers.css" />

<script type="text/javascript" src="/js/jquery-2.2.2.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/apps/academy/lib/common.js"></script>

<script> 
function TnFindZipNew(frmname){
	$.ajax({
		//url: "/html/common/zipcode.asp",
		url: "/apps/academy/lib/searchzipNew.asp?target="+ frmname,
		cache: false,
		success: function(rst) {
			$("#hBoxes").html(rst);	//hBoxes는 incheader에 div있음.. 우편번호 완성시 이 주석삭제
		}
	});
}

$(function(){
	var scroll01 = new Swiper(".scroll01 .scrollArea .swiper-container", {
		scrollbar:'.swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode:true
	});

	// add file
	$('#findFile').change(function(){
		$('.fileVal').text($(this).val());
	});

	//휴대폰 번호 입력
	$("#teacherPhone").blur(function(){
		$(this).val($(this).val().replace(/[^0-9]/g,"").replace(/^(01[016789]{1}|02|0[3-9]{1}[0-9]{1})-?([0-9]{3,4})-?([0-9]{4})$/, "$1-$2-$3"));
	});

	//전번 번호 입력
	$("#teacherTel").blur(function(){
		$(this).val($(this).val().replace(/[^0-9]/g,"").replace(/^(01[016789]{1}|02|0[3-9]{1}[0-9]{1})-?([0-9]{3,4})-?([0-9]{4})$/, "$1-$2-$3"));
	});


});

function fnWriterSubmit(f){
	if($(':radio[name="gubun"]:checked').length < 1){
		alert('개인/공방/기업을 선택해주세요.');
		return;
	}
	if (f.writername.value == ''){
		alert('작가명을 입력해주세요.');
		f.writername.focus();
		return;
	}
	if (f.bunya.value == ''){
		alert('작품 분야를 입력해주세요.');
		f.bunya.focus();
		return;
	}
	if ($.trim($("#txZip").val()).length<5){
		alert('주소를 입력해주세요.');
		f.txZip.focus();
		return;
	}
	if (f.txAddr2.value == ''){
		alert('주소를 입력해주세요.');
		f.txAddr2.focus();
		return;
	}
	if (f.txCell.value == ''){
		alert('휴대폰 번호를 입력해주세요.');
		f.txCell.focus();
		return;
	}
	if (f.usermail.value == ''){
		alert('이메일을 입력해주세요.');
		f.usermail.focus();
		return;
	}	
	if (!check_form_email(f.usermail.value)){
	    alert("이메일 주소가 유효하지 않습니다.");
		f.usermail.focus();
		return ;
	}
	if (f.introduce.value == ''){
		alert('작품 소개글을 입력해주세요.');
		f.introduce.focus();
		return;
	}									
	if(!f.agreechk.checked){
		alert("문의에 필요한 정보 수접 및\n이용약관에 동의해주세요.");
		frm.agreechk.focus();
		return false;
	}
	var ret = confirm('정확히 입력하셨습니까?');
	if(ret){
		f.target = "FrameCKP";
		f.submit();
	}
}

function check_form_email(email){
	var pos;
	pos = email.indexOf('@');

	if (pos < 0)				//@가 포함되어 있지 않음
		return(false);
	else
		{
		pos = email.indexOf('@', pos + 1)
		if (pos >= 0)			//@가 두번이상 포함되어 있음
			return(false);
		}
	pos = email.indexOf('.');

	if (pos < 0)				//@가 포함되어 있지 않음
		return false;
	return(true);
}

<%
'//크로스 도메인 체크 URL(corpse2)
If (application("Svr_Info")	= "Dev") Then
%>
var allowOrigin = "http://testupload.10x10.co.kr;
<% Else %>
var allowOrigin = "https://upload.10x10.co.kr";
<% End If%>
window.addEventListener("message", postMessageController, true);
window.attachEvent("onmesage", postMessageController);

function postMessageController(e) {
	//alert(e.data);
	if(e.origin == allowOrigin){
		alert(e.data);
		fnAPPclosePopup();
	}
}

</script>
</head>
<body>
<div class="wrap">
	<div class="container headB bgGry1">
		<div class="content partnerRequest">
			<ul class="tabNav tab1">
				<li class="current"><a href="/apps/academy/partnership/partnership_register.asp" onclick="location.replace('/apps/academy/partnership/partnership_register.asp');return false;">작가 신청</a></li>
				<li><a href="/apps/academy/partnership/partnership_lecturer.asp" onclick="location.replace('/apps/academy/partnership/partnership_lecturer.asp');return false;">강사 신청</a></li>
			</ul>
			<div class="box1 lectureType">
			<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/academy/partnership/doPartnerShip_register_app.asp" EncType="multipart/form-data">
				<input type="hidden" name="mode" value="writer">
				<div class="restInfo box2">
					<dl>
						<dt>문의 전 확인사항</dt>
						<dd>
							<ul>
								<li>더핑거스에서 판매하고 싶은 작품을 제안해주세요!</li>
								<li>문의가 접수 완료되면 담당자 검토 후, 전화나 이메일로 연락 드립니다.</li>
								<li>이메일 문의 : <a href="mailto:customer@thefingers.co.kr" class="link">customer@thefingers.co.kr</a></li>
							</ul>
						</dd>
					</dl>
				</div>
				<div class="artistType flexInp">
					<p><span class="form1"><input type="radio" id="type1" name="gubun" checked value="1" /><label for="type1"><em>개인</em></label></span></p>
					<p><span class="form1"><input type="radio" id="type2" name="gubun" value="2" /><label for="type2"><em>공방</em></label></span></p>
					<p><span class="form1"><input type="radio" id="type3" name="gubun" value="3" /><label for="type3"><em>기업</em></label></span></p>
				</div>
				<div class="fingerTable tMar1-5r">
					<table>
						<caption>희망 수강내용 입력</caption>
						<tr>
							<th><label for="artistName">작가명<em class="needs">*</em><br /><span class="fs1-2r cGry6">(업체명)</span></label></th>
							<td><input type="text" id="artistName" name="writername" /></td>
						</tr>
						<tr>
							<th><label for="kind">작품분야<em class="needs">*</em></label></th>
							<td><input type="text" id="kind" name="bunya" /></td>
						</tr>
						<tr>
							<th>주소<em class="needs">*</em></th>
							<td>
								<button type="button" class="btn btnM1 btnGry2" onclick="TnFindZipNew('frm');return false;">우편번호 검색</button><input type="text" class="inpZipcode" name="txZip" id="txZip" value="" readonly />
								<p class="tMar0-5r"><input type="text" name="txAddr1" value="" readonly /></p>
								<p class="tMar0-5r"><input type="text" name="txAddr2" value="" /></p>
							</td>
						</tr>
						<tr>
							<th><label for="teacherPhone">휴대폰<em class="needs">*</em></label></th>
							<td><input type="tel" id="teacherPhone" name="txCell" /></td>
						</tr>
						<tr>
							<th><label for="teacherTel">전화</label></th>
							<td><input type="tel" id="teacherTel" name="userphone" /></td>
						</tr>
						<tr>
							<th><label for="artistMail">이메일<em class="needs">*</em></label></th>
							<td><input type="email" id="artistMail" placeholder="예)id@example.com" name="usermail" /></td>
						</tr>
						<tr>
							<th><label for="homepage">홈페이지</label></th>
							<td><input type="url" id="homepage" name="homepage" value="http://" /></td>
						</tr>
						<tr>
							<th><label for="introduce">작품소개<em class="needs">*</em></label></th>
							<td><textarea id="introduce" name="introduce" cols="20" rows="4"></textarea></td>
						</tr>
						<tr>
							<th><label for="etc">기타</label></th>
							<td><textarea id="etc" name="etc" cols="20" rows="4"></textarea></td>
						</tr>
					</table>
				</div>
			</div>
			<div class="box1">
				<div style="display:none;">
				<h2>첨부파일</h2>
				<div class="addFile">
					<div>
						<label for="findFile">찾아보기</label>
						<input type="file" id="findFile" name="writefile" />
						<p class="fileVal"></p>
					</div>
					<p class="tPad0-5r lPad0-5r cGry5">*10MB이하 GIF, JPEG, PNG 파일만 업로드가 가능합니다.</p>
				</div>
				</div>
				<div class="requestAgree tPad3r" style="padding-top:0 !important;">
					<div class="scroll01 box2 bMar1r">
						<div class="scrollArea">
							<div class="swiper-container">
								<div class="swiper-wrapper">
									<div class="swiper-slide">
										<div class="restInfo">
											<dl>
												<dt>문의를 위한 정보수집 및 이용동의</dt>
												<dd>
													<p class="bPad1r cGry2">(주)텐바이텐(이하 &quot;회사&quot;라 함)는 개인정보보호법, 정보통신망 이용촉진 및 정보보호 등에 관한 법률 등 관련 법령상의 개인정보보호 규정을 준수하며, 파트너의 개인정보 보호에 최선을 다하고 있습니다.</p>
													<ul>
														<li>
															1. 개인정보 수집 및 이용주체
															<p>강사신청 문의 신청을 통해 제공하신 정보는 &quot;회사&quot;가 직접 접수하고 관리합니다.</p>
														</li>
														<li>
															2. 동의를 거부할 권리 및 동의 거부에 따른 불이익
															<p>신청자는 개인정보제공 등에 관해 동의하지 않을 권리가 있습니다.(이 경우 강사신청 문의는 불가능합니다.)</p>
														</li>
														<li>
															3. 수집하는 개인정보 항목
															<p>신청자명, 신청자 주소, 신청자 연락처, 신청자 이메일 주소, 신청자 소개</p>
														</li>
														<li>
															4. 수집 및 이용목적
															<p>강사신청 검토, 강사신청 관리시스템의 운용, 공지사항의 전달 등</p>
														</li>
														<li>
															5. 보유기간 및 이용기간
															<p>수집된 정보는 강사 기간이 종료되는 시점까지 보관됩니다.</p>
														</li>
													</ul>
												</dd>
											</dl>
										</div>
									</div>
								</div>
								<div class="swiper-scrollbar"></div>
							</div>
						</div>
					</div>
					<span class="form4"><input type="checkbox" name="agreechk" id="agree" /><label for="agree"><em>위 내용에 동의합니다</em></label></span>
				</div>
			</form>
			</div>
			<div class="btnGroup pad1-5r">
				<button type="button" class="btn btnB1 btnYgn" onclick="javascript:fnWriterSubmit(frm);">문의하기</button>
			</div>
		</div>
		<div id="layerMask" class="layerMask"></div>
		<% if (application("Svr_Info")	= "Dev") then %>
		<iframe name="FrameCKP" src="about:blank" frameborder="1" width="300" height="500"></iframe>
		<% else %>
		<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
		<% end if %>
	<div id="hBoxes"></div>
	</div>
</div>
</body>
</html>