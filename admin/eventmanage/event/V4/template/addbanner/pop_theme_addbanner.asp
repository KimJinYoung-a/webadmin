<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
.popWinV17 {overflow:hidden; position:absolute; left:0; top:0; right:0; bottom:0; width:100%; height:100%; font-family:"malgun Gothic","맑은고딕", Dotum, "돋움", sans-serif;}
.popWinV17 h1 {height:40px; padding:12px 15px 0; color:#fff; font-size:17px; background:#c80a0a; border-bottom:1px solid #d80a0a}
.popWinV17 h2 {position:relative; padding:12px 15px; color:#333; font-size:12px; font-weight: bold; background-color:#444; border-top:1px solid #666; font-family:"malgun Gothic","맑은고딕", Dotum, "돋움", sans-serif; z-index:55; color:#fff;}
.popContainerV17 {position:absolute; left:0; top:40px; right:0; bottom:90px; width:100%; border-bottom:1px solid #ddd;}
.contL {position:relative; width:65%; height:100%; border-right:1px solid #ddd; z-index:10; overflow-y:auto;}
.contR {position:absolute; right:0; top:0; bottom:0; width:30%; height:100%; border-left:1px solid #ddd;}
.tbListWrap {position:relative; width:100%; height:100%;}
.tbDataList, .thDataList {display:table; width:100%;}
.tbDataList li, .thDataList li {display:table; width:100%; margin-top:-1px; border-top:1px solid #ddd; border-bottom:1px solid #ddd; }
.thDataList li {height:33px; background-color:#eaeaea; border-top:2px solid #ccc; font-weight:bold;}
.tbDataList li {background-color:#fff; z-index:100;}
.tbDataList li p, .thDataList li p {display:table-cell; padding:7px; text-align:center; vertical-align:middle; line-height:1.4;}
.thDataList li p {white-space:nowrap;}
.handling {background-color:rgba(42,42,57,0.2) !important; height:30px; border:none;}
#sortable li {cursor:move;}
.popBtnWrap {position:absolute; left:0; bottom:0; width:100%; height:60px; text-align:center;}
.textOverflow {width:100%; display:block; text-overflow:ellipsis; overflow:hidden; white-space:nowrap;}
.btnMove {position:absolute; left:67.5%; top:50%; width:40px; height:70px; margin-top:-35px; margin-left:-20px; padding:0; border:none; background:transparent url(/images/btn_move_arrow.png) no-repeat 50% 50%; z-index:1000; cursor:pointer;}
</style>
</head>
<body>
<!-- 팝업 사이즈 : 최소 1100*750 -->
<div class="popWinV17">
	<h1>Unit 검색</h1>
	<div class="popContainerV17">
		<div class="contL">
			<h2>Unit 선택</h2>
			<div class="tab" style="margin:-1px 0 0 -1px;">
				<ul>
					<li class="col11 selected"><a href="#unitType01">상품</a></li>
					<li class="col11 "><a href="#unitType02">이벤트</a></li>
					<li class="col11 "><a href="#unitType03">컨텐츠</a></li>
				</ul>
			</div>
			<!-- 상품 Tab -->
			<div id="unitType01" class="unitPannel">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">카테고리 :</label>
								<select class="formSlt" id="deal" title="옵션 선택">
									<option>전체</option>
									<option>서비스</option>
									<option>카테고리</option>
									<option>브랜드</option>
									<option>상품</option>
									<option>키워드</option>
								</select>
								<select class="formSlt" id="deal3" title="옵션 선택">
									<option>전체</option>
									<option>서비스</option>
									<option>카테고리</option>
									<option>브랜드</option>
									<option>상품</option>
									<option>키워드</option>
								</select>
								<select class="formSlt" id="deal4" title="옵션 선택">
									<option>전체</option>
									<option>서비스</option>
									<option>카테고리</option>
									<option>브랜드</option>
									<option>상품</option>
									<option>키워드</option>
								</select>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">검색어 :</label>
								<input type="text" class="formTxt" id="schWord" style="width:400px" placeholder="상품ID 또는 상품명을 입력하여 검색하세요." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="검색" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="ftLt lPad10">
						<select class="formSlt" id="deal" title="옵션 선택">
							<option>신상품순</option>
							<option>인기순</option>
						</select>
					</div>
					<div class="ftRt pad10">
						<span>검색결과 : <strong>999,999</strong></span> <span class="lMar10">페이지 : <strong>1 / 30,000</strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell10">상품 ID</p>
							<p class="cell10">이미지</p>
							<p>상품명</p>
							<p class="cell10">가격</p>
							<p class="cell10">업체 ID</p>
							<p class="cell10">판매여부</p>
						</li>
					</ul>
					<ul class="tbDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell10">17026941</p>
							<p class="cell10"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">[간지나베베] 썸머플레인원피스(12346) 썸머플레인원피스(12346)썸머플레인원피스(12346)</p>
							<p class="cell10">316,000</p>
							<p class="cell10">milliens</p>
							<p class="cell10">홍길동</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell10">17026942</p>
							<p class="cell10"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">두산베어스 철웅이</p>
							<p class="cell10">316,000</p>
							<p class="cell10">milliens</p>
							<p class="cell10">안영이</p>
						</li>
					</ul>
					<div class="ct tPad20 cBk1">
						<a href="">[prev]</a>
						<a href=""><span class="cRd1">[1]</span></a>
						<a href="">[2]</a>
						<a href="">[3]</a>
						<a href="">[4]</a>
						<a href="">[5]</a>
						<a href="">[6]</a>
						<a href="">[7]</a>
						<a href="">[8]</a>
						<a href="">[9]</a>
						<a href="">[10]</a>
						<a href="">[next]</a>
					</div>
				</div>
			</div>
			<!--// 상품 Tab -->
			<!-- 이벤트 Tab -->
			<div id="unitType02" class="unitPannel" style="display:none;">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">기간 :</label>
								<select class="formSlt" title="옵션 선택">
									<option>시작일</option>
									<option>종료일</option>
								</select>
								<input type="text" class="formTxt" id="term1" style="width:100px" placeholder="시작일" />
								<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" />
								~
								<input type="text" class="formTxt" id="term2" style="width:100px" placeholder="종료일" />
								<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" />
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<p class="formTit">이벤트 유형 :</p>
								<select class="formSlt" id="deal" title="옵션 선택">
									<option>전체</option>
									<option>쇼핑찬스</option>
								</select>
							</li>
							<li>
								<p class="formTit">카테고리 :</p>
								<select class="formSlt" id="deal" title="옵션 선택">
									<option>전체</option>
									<option>디자인문구</option>
									<option>디지털</option>
								</select>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">검색어 :</label>
								<input type="text" class="formTxt" id="schWord" style="width:400px" placeholder="이벤트코드 또는 이벤트명을 입력하여 검색하세요." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="검색" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="rt pad10">
						<span>검색결과 : <strong>999,999</strong></span> <span class="lMar10">페이지 : <strong>1 / 30,000</strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">이벤트 코드</p>
							<p class="cell12">이벤트 유형</p>
							<p class="cell12">배너</p>
							<p>이벤트명</p>
							<p class="cell12">카테고리</p>
							<p class="cell12">시작일</p>
							<p class="cell12">종료일</p>
						</li>
					</ul>
					<ul class="tbDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">쇼핑찬스</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">무한도전 X 스티키몬스터랩 무도몬 2종</p>
							<p class="cell12">디자인문구</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">쇼핑찬스</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">무한도전 X 스티키몬스터랩 무도몬 2종</p>
							<p class="cell12">디자인문구</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">쇼핑찬스</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">무한도전 X 스티키몬스터랩 무도몬 2종</p>
							<p class="cell12">디자인문구</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">쇼핑찬스</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">무한도전 X 스티키몬스터랩 무도몬 2종</p>
							<p class="cell12">디자인문구</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">쇼핑찬스</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">무한도전 X 스티키몬스터랩 무도몬 2종</p>
							<p class="cell12">디자인문구</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
					</ul>
					<div class="ct tPad20 bPad20 cBk1">
						<a href="">[prev]</a>
						<a href=""><span class="cRd1">[1]</span></a>
						<a href="">[2]</a>
						<a href="">[3]</a>
						<a href="">[4]</a>
						<a href="">[5]</a>
						<a href="">[6]</a>
						<a href="">[7]</a>
						<a href="">[8]</a>
						<a href="">[9]</a>
						<a href="">[10]</a>
						<a href="">[next]</a>
					</div>
				</div>
			</div>
			<!--// 이벤트 Tab -->
			<!-- 컨텐츠 Tab -->
			<div id="unitType03" class="unitPannel" style="display:none;">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">오픈일 :</label>
								<input type="text" class="formTxt" id="term1" style="width:100px" placeholder="시작일" />
								<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" />
								~
								<input type="text" class="formTxt" id="term2" style="width:100px" placeholder="종료일" />
								<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" />
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit">카테고리 :</label>
								<select class="formSlt" id="deal" title="옵션 선택">
									<option>전체</option>
									<option>서비스</option>
									<option>카테고리</option>
									<option>브랜드</option>
									<option>상품</option>
									<option>키워드</option>
								</select>
								<select class="formSlt" id="deal3" title="옵션 선택">
									<option>전체</option>
									<option>서비스</option>
									<option>카테고리</option>
									<option>브랜드</option>
									<option>상품</option>
									<option>키워드</option>
								</select>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">검색어 :</label>
								<input type="text" class="formTxt" id="schWord" style="width:400px" placeholder="타이틀을 검색하세요." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="검색" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="rt pad10">
						<span>검색결과 : <strong>999,999</strong></span> <span class="lMar10">페이지 : <strong>1 / 30,000</strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">Idx</p>
							<p class="cell12">카테고리1</p>
							<p class="cell15">카테고리2</p>
							<p class="cell12">이미지</p>
							<p>타이틀</p>
							<p class="cell12">오픈일</p>
						</li>
					</ul>
					<ul class="tbDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">히치하이커</p>
							<p class="cell15">MOVIE</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">vol.63 &lt;두근두근 설레임&gt;</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">플레잉</p>
							<p class="cell15">TALK &gt; AZIT&</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">눈으로 다녀오는 홍콩 영화 여행!</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">히치하이커</p>
							<p class="cell15">!NSPIRATION &gt; DESIGN</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">vol.63 &lt;두근두근 설레임&gt;</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">히치하이커</p>
							<p class="cell15">THING. &gt; thingthing</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">vol.63 &lt;두근두근 설레임&gt;</p>
							<p class="cell12">2017-05-30</p>
						</li>
					</ul>
					<div class="ct tPad20 bPad20 cBk1">
						<a href="">[prev]</a>
						<a href=""><span class="cRd1">[1]</span></a>
						<a href="">[2]</a>
						<a href="">[3]</a>
						<a href="">[4]</a>
						<a href="">[5]</a>
						<a href="">[6]</a>
						<a href="">[7]</a>
						<a href="">[8]</a>
						<a href="">[9]</a>
						<a href="">[10]</a>
						<a href="">[next]</a>
					</div>
				</div>
			</div>
			<!--// 컨텐츠 Tab -->
		</div>

		<input type="button" class="btnMove" title="선택해서 담기" />

		<div class="contR">
			<h2 style="margin-left:-1px;">Unit 선택 정보</h2>
			<div class="tbListWrap">
				<ul class="thDataList">
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">Unit 구분</p>
						<p>Unit명</p>
					</li>
				</ul>
				<ul id="sortable" class="tbDataList">
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">상품</p>
						<p class="lt"><span class="textOverflow">[간지나베베] 썸머플레인원피스(12346)</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">컨텐츠</p>
						<p class="lt"><span class="textOverflow">sunny tote bag yellow</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">이벤트</p>
						<p class="lt"><span class="textOverflow">두산베어스 철웅이</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">컨텐츠</p>
						<p class="lt"><span class="textOverflow">sunny tote bag yellow</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">이벤트</p>
						<p class="lt"><span class="textOverflow">두산베어스 철웅이</span></p>
					</li>
				</ul>
				<div class="pad10 rt">
					<input type="button" class="btn" value="선택삭제" onclick="" />
				</div>
			</div>
		</div>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="선택완료" onclick="" class="cRd1" style="width:100px; height:30px;" />
		<input type="button" value="취소" onclick="" style="width:100px; height:30px;" />
	</div>
</div>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function() {
	$("#sortable").sortable({
		placeholder:"handling"
	}).disableSelection();

	$(".tab li").click(function() {
		$(".tab li").removeClass('selected');
		$(this).addClass('selected');
		$('.unitPannel').hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
		return false;
	});
});
</script>
</body>
</html>