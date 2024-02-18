<%@ language=vbscript %>
    <% option explicit %>
        <% '#############################################################
' Description : 유료 사은품 관리 '	History		: 2022.01.05 이전도 생성
' ############################################################# %>
            <!-- #include virtual="/lib/function.asp"-->
            <!-- #include virtual="/lib/db/dbopen.asp" -->
            <!-- #include virtual="/lib/util/htmllib.asp" -->
            <!-- #include virtual="/admin/incSessionAdmin.asp" -->
            <!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
            </p>
            <link rel="stylesheet" type="text/css" href="/css/commonV20.css" />
            <div>
                <!-- 구매혜택 만들기 -->
                <div class="content paidGift add">
                    <div class="paidGift_top">
                        <div class="search_wrap">
                            <select name="" id="select01">
                                <option value="">구분</option>
                                <option value="">플러스세일</option>
                                <option value="">무료사은품</option>
                            </select>
                            <select name="" id="select02">
                                <option value="">진행상태</option>
                                <option value="">진행중</option>
                                <option value="">진행예정</option>
                                <option value="">종료</option>
                                <option value="">사용안함</option>
                            </select>
                            <div class="input_wrap">
                                <select name="" id="select03">
                                    <option value="">담당부서</option>
                                    <option value="">담당자</option>
                                </select>
                                <span></span>
                                <button class="btn_select">선택하기</button>
                                <li class="selected" style="display:none;">마케팅<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                             <div class="input_wrap">
                                <select name="" id="select04">
                                    <option value="">노출조건</option>
                                    <option value="">생애 첫 결제</option>
                                    <option value="">최근 결제금액</option>
                                    <option value="">회원등급</option>
                                    <option value="">구매금액</option>
                                    <option value="">구매횟수</option>
                                    <option value="">상품</option>
                                    <option value="">카테고리</option>
                                    <option value="">브랜드</option>
                                    <option value="">기획전/이벤트</option>
                                </select>
                                <li class="selected" style="display:none;">기획전/이벤트<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                                <li class="selected" style="display:none;">상품<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                            <div class="input_wrap">
                                <select name="" id="select05">
                                    <option value="">구매혜택/상품/사은품명</option>
                                    <option value="">구매혜택 번호</option>
                                    <option value="">사은품코드</option>
                                    <option value="">상품코드</option>
                                </select>
                                <span></span>
                                <input type="text" placeholder="검색어를 입력해주세요">
                            </div>
                            <button class="btn_search"><img src="https://webadmin.10x10.co.kr/images/icon/search.png">검색하기</button>
                        </div>
                        <div class="tgl_wrap">
                            <span>내가 등록한 구매혜택만 보기</span>
                            <div class="tgl_btn">
                                <input type="checkbox" id="tgl_btn_my">
                                <label for="tgl_btn_my"></label>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_aside">
                        <div class="list_wrap">
                            <div class="list_top">
                                <li>총 <span>103</span>건</li>
                            </div>
                            <div class="list_cont">
                                <a href=""><div class="cont cont_new">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[구분] 새 구매혜택</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>생애 첫 결제</span></li>
                                        <li><span></span></li>
                                        <li class="state st01" style="display:none;"></li>
                                        <p>12:31:20 자동저장</p>
                                    </ul>
                                    <button class="delete"><img src="https://webadmin.10x10.co.kr/images/icon/trash_red.png"></button>
                                </div></a>
                                <!-- 임시저장 구매혜택 -->
                                <a href=""><div class="cont" style="display:none;">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st03">임시저장</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st01">오픈예정</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st02">진행중</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st02">진행중</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st02">진행중</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st01">오픈예정</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st03">종료</li>
                                    </ul>
                                </div></a>
                            </div>
                            <div class="list_bottom">
                                <ul class="pagination">
                                    <li class="on"><a>1</a></li>
                                    <li class=""><a>2</a></li>
                                    <li class=""><a>3</a></li>
                                    <li class=""><a>4</a></li>
                                    <li class=""><a>5</a></li> 
                                    <li class=""><a>></a></li>
                                    <li class=""><a>>></a></li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_section">
                    <!-- 조건노출 : on -->
                        <div class="steps">
                            <li class="step on">혜택 노출조건 설정</li>
                            <li class="step on"><span></span>기간설정</li>
                            <li class="step on"><span></span>혜택설정</li>
                            <li class="step"><span></span>최종점검</li>
                        </div>
                        <div class="step_wrap step03">
                            <div class="step_noti on"><span>설정된 고객에게 어떤 혜택을 제공할까요?</span></div>
                            <div class="btn_group">
                                <button class="type01 step03_01 on"><span>플러스 세일</span><p class="img"><img></p></button>
                                <button class="type01 step03_02"><span>무료사은품</span><p class="img"><img></p></button>
                            </div>
                            <!-- 플러스세일 -->
                            <div class="step_cont step03_01">
                             <li>구매자에게 지정된 상품의 할인 혜택을 제공합니다. 혜택을 구성해주세요.</li>
                             <button class="btn_que"><img src="https://webadmin.10x10.co.kr/images/icon/question.png" alt=""></button>
                                <div class="step_detail">
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>그룹없이 상품 등록</h3>
                                            <li>그룹 설정 없이 상품을 노출합니다.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>단순 그룹 설정</h3>
                                            <li>상품그룹을 분류하고, 그룹에 이름을 붙일 수 있습니다.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>금액별 그룹 설정</h3>
                                            <li>구매금액별로 상품을 분류하여, 조건을 만족한 경우에만 구매 가능하게 합니다.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                                <!-- 그룹없이 상품 등록 -->
                                <div class="step_detail group group01_01" style="display:block;">
                                    <div class="table_wrap">
                                        <div class="ttop">
                                            <span>총 0건</span><span class="selected">선택된 상품 1건</span>
                                            <button>선택상품 삭제</button>
                                        </div>
                                        <div class="table">
                                            <ul class="thead">
                                                <li><input type="checkbox"></li>
                                                <li>상품코드</li>
                                                <li>대표이미지/브랜드/상품명</li>
                                                <li>추가정보</li>
                                                <li>재고 소진율<span>소진수량/전체수량</span></li>
                                                <li>판매가</li>
                                                <li>매입가</li>
                                                <li>마진</li>
                                                <li>계약구분</li>
                                                <li></li>
                                            </ul>
                                            <div class="tbody_wrap add">
                                                <ul class="tbody">
                                                    <a href=""><span>
                                                        + 상품 추가하기
                                                    </span></a>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge">뱃지[MD특가]</p>
                                                        <p class="info">유의사항</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent">99%</p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99/100</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% 할인<span>38,000</span></p>
                                                            <p class="price p_type03">7% 쿠폰<span>38,000</span></p>
                                                            <p class="price p_type04">15% 플러스세일<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        매입
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>수정</button>
                                                            <button>삭제</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info"></p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent"></p>
                                                            <ul class="bar" style="display:none;">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% 할인<span>38,000</span></p>
                                                            <p class="price p_type03">7% 쿠폰<span>38,000</span></p>
                                                            <p class="price p_type04">15% 플러스세일<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        매입
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>수정</button>
                                                            <button>삭제</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- 단순 그룹 설정 -->
                                 <div class="step_detail group group02_01" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>다이어리에는 뭐니뭐니해도 스티커지</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>치트키는 떡메모지</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+그룹 추가하기</li></button>
                                    </div>
                                    <!-- 그롭없이 상품 등록과 동일<div class="table_wrap"></div> -->
                                </div>
                                <!-- 금액별 그룹 설정 -->
                                 <div class="step_detail group group03_01" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><li>10,000원~<span>&#65372;</span><span class="gr_cond">구매금액</span><span class="gr_name">10,000원 이상 구매했다면</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><li>10,000원~<span>&#65372;</span><span class="gr_cond">구매금액</span><span class="gr_name">10,000원 이상 구매했다면</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+금액별 그룹 추가하기</li></button>
                                    </div>
                                    <!-- 그롭없이 상품 등록과 동일<div class="table_wrap"></div> -->
                                </div>
                            </div>
                            <!-- 무료사은품 -->
                            <div class="step_cont step03_02 on">
                             <li>구매자에게 지정된 사은품을 무료로 제공합니다. 구매자는 단 하나의 사은품을 선택하여 받을 수 있습니다. 혜택을 구성해주세요.</li>
                             <button class="btn_que"><img src="https://webadmin.10x10.co.kr/images/icon/question.png" alt=""></button>
                                <div class="step_detail">
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>그룹없이 상품 등록</h3>
                                            <li>그룹 설정 없이 상품을 노출합니다.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>단순 그룹 설정</h3>
                                            <li>사은품을 분류하고, 그룹에 이름을 붙일 수 있습니다.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>금액별 그룹 설정</h3>
                                            <li>구매금액별로 사은품을 분류하여, 조건을 만족한 경우에만 구매 가능하게 합니다.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                                <!-- 그룹없이 상품 등록 -->
                                <div class="step_detail group group01_02" style="display:block;">
                                    <div class="table_wrap">
                                        <div class="ttop">
                                            <span>총 0건</span><span class="selected">선택된 상품 1건</span>
                                            <button>선택상품 삭제</button>
                                        </div>
                                        <div class="table">
                                            <ul class="thead">
                                                <li><input type="checkbox"></li>
                                                <li>사은품코드</li>
                                                <li>대표이미지/브랜드/상품명</li>
                                                <li>추가정보</li>
                                                <li>재고 소진율<span>소진수량/전체수량</span></li>
                                                <li>판매가</li>
                                                <li>매입가</li>
                                                <li>마진</li>
                                                <li>계약구분</li>
                                                <li></li>
                                            </ul>
                                            <div class="tbody_wrap add">
                                                <ul class="tbody">
                                                    <div>
                                                        + 사은품 추가하기
                                                        <div class="add_btn">
                                                            <button>사은품 신규등록</button><span></span>
                                                            <button>기존 사은품 불러오기</button>
                                                        </div>
                                                    </div>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge">뱃지[MD특가]</p>
                                                        <p class="info">유의사항</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent">99%</p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99/100</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% 할인<span>38,000</span></p>
                                                            <p class="price p_type03">7% 쿠폰<span>38,000</span></p>
                                                            <p class="price p_type04">15% 플러스세일<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        매입
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>수정</button>
                                                            <button>삭제</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info"></p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent"></p>
                                                            <ul class="bar" style="display:none;">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% 할인<span>38,000</span></p>
                                                            <p class="price p_type03">7% 쿠폰<span>38,000</span></p>
                                                            <p class="price p_type04">15% 플러스세일<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        매입
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>수정</button>
                                                            <button>삭제</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- 단순 그룹 설정 -->
                                 <div class="step_detail group group02_02" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>다이어리에는 뭐니뭐니해도 스티커지</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>치트키는 떡메모지</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+그룹 추가하기</li></button>
                                    </div>
                                    <!-- 그롭없이 상품 등록과 동일<div class="table_wrap"></div> -->
                                </div>
                                <!-- 금액별 그룹 설정 -->
                                 <div class="step_detail group group03_02" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><li>10,000원~<span>&#65372;</span><span class="gr_cond">구매금액</span><span class="gr_name">10,000원 이상 구매했다면</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><li>10,000원~<span>&#65372;</span><span class="gr_cond">구매금액</span><span class="gr_name">10,000원 이상 구매했다면</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+금액별 그룹 추가하기</li></button>
                                    </div>
                                    <div class="table_wrap t02">
                                        <div class="ttop">
                                            <span>총 0건</span><span class="selected">선택된 상품 1건</span>
                                            <button>선택상품 삭제</button>
                                        </div>
                                        <div class="table">
                                            <ul class="thead">
                                                <li><input type="checkbox"></li>
                                                <li>상품코드</li>
                                                <li>대표이미지/브랜드/상품명</li>
                                                <li>추가정보</li>
                                                <li>재고 소진율<span>소진수량/전체수량</span></li>
                                                <li>배송 QR 쿠폰/마일리지 정보</li>
                                                <li></li>
                                            </ul>
                                            <!-- 그룹 추가 전 -->
                                            <div class="tbody_wrap none">
                                                <ul class="tbody">
                                                    그룹을 먼저 추가하면 사은품을 추가할 수 있어요!
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap add">
                                                <ul class="tbody">
                                                    <div>
                                                        + 사은품 추가하기
                                                        <div class="add_btn">
                                                            <button>사은품 신규등록</button><span></span>
                                                            <button>기존 사은품 불러오기</button>
                                                        </div>
                                                    </div>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info">유의사항</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent">100%</p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">100/100</p>
                                                        </ul>
                                                    </li>
                                                   <li>
                                                        <ul>
                                                            <p class="info">텐바이텐</p>
                                                            <p class="coupon"></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>수정</button>
                                                            <button>삭제</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info">유의사항</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent"></p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">50,000/50,000</p>
                                                        </ul>
                                                    </li>
                                                   <li>
                                                        <ul>
                                                            <p class="info"></p>
                                                            <p class="coupon">보너스쿠폰 지금 3,000</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>수정</button>
                                                            <button>삭제</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="steps_bottom">
                            <ul>
                                <button class="delete">취소</button>
                                <!-- 버튼 비활성화 : next on -->
                                <button class="next">다음</button>
                            </ul>
                        </div>
                    </div>

                    <!-- 레이어팝업 노출 : paidGift_lyr on -->
                    <div class="lyr paidGift_lyr">
                        <div class="lyr_overlay"></div>
                     <!-- lyr_wrap lyr01 on -->
                        <!-- 그룹 추가하기 -->
                        <div class="lyr_wrap lyr10 on">
                            <div class="lyr_top">
                                <li>그룹 추가하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 그룹 수정하기 -->
                        <div class="lyr_wrap lyr11">
                            <div class="lyr_top">
                                <li>그룹 수정하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>그룹 이름</h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <button class="initial">그룹삭제</button>
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 금액별 그룹 추가하기 -->
                        <div class="lyr_wrap lyr12">
                            <div class="lyr_top">
                                <li>금액별 그룹 추가하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <!-- 스페셜 이벤트 - 금액별 그룹 추가하기 -->
                                 <div class="lyr_cont" style="display:none;">
                                    <li class="type01 noti">혜택 노출 조건을 스페셜이벤트로 설정한 경우<br> 금액 기준은 설정한 스페셜 이벤트의 기준과 동일하게 적용됩니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>그룹 이름<span class="gray">29/30</span></h4></h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액 설정</h4>
                                    <div class="cont_wrap">
                                        <input class="type04" type="checkbox" id="lyr12_01"><label for="lyr12_01">총 구매금액</label>
                                        <input class="type04" type="checkbox" id="lyr12_02"><label for="lyr12_02">카테고리</label>
                                        <input class="type04" type="checkbox" id="lyr12_03"><label for="lyr12_03">브랜드</label>
                                        <input class="type04" type="checkbox" id="lyr12_04"><label for="lyr12_04">기획전/이벤트</label>
                                    </div>
                                </div>
                                <!-- 총 구매금액 선택시 노출 -->
                                <div class="lyr_cont">
                                    <h4>배송 구분</h4>
                                    <input type="checkbox" id="lyr12_05" class="type01"><label for="lyr12_05"><span class="circle"></span>전체상품</label>
                                    <input type="checkbox" id="lyr12_06" class="type01"><label for="lyr12_06"><span class="circle"></span>텐바이텐 배송 포함</label>
                                </div>
                                <!-- 카테고리 선택시 노출 -->
                                <div class="lyr_cont" style="display:block;">
                                    <h4>카테고리 지정<span>*복수선택 가능합니다.</span></h4>
                                    <select name="" id="">
                                        <option value="">1depth</option>
                                    </select>
                                    <select name="" id="">
                                        <option value="">2depth</option>
                                    </select>
                                    <button class="add btn_blue">카테고리 추가</button>
                                    <div class="option">
                                        <div class="option_added">
                                            <li>디자인문구</li>
                                            <li>></li>
                                            <li>다이어리/플래너</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                        <div class="option_added">
                                            <li>디자인문구</li>
                                            <li>></li>
                                            <li>데코레이션</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                    </div>
                                </div>
                                <!-- 브랜드 선택시 노출 -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>브랜드 지정<span>*복수선택 가능합니다.</span></h4>
                                    <div class="cont_wrap search brand">
                                         <input type="text" placeholder="브랜드ID를 입력해주세요">
                                        <button class="add btn_blue">추가하기</button>
                                        <button class="search btn_white">브랜드 찾기</button>
                                    </div>
                                </div>
                                <!-- 기획전/이벤트 선택시 노출 -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>기획전/이벤트 지정<span>*복수선택 가능합니다.</span></h4>
                                    <div class="cont_wrap search event">
                                        <input type="text" placeholder="이벤트 코드를 입력해주세요">
                                        <button class="add btn_blue">추가하기</button>
                                    </div>
                                </div>
                                <!-- 기획전/이벤트 선택시 노출 -->
                                 <div class="lyr_cont" style="display:block;">
                                    <h4>상품 구매 조건</h4>
                                    <input type="checkbox" id="lyr12_07" class="type01"><label for="lyr12_07"><span class="circle"></span>1개라도 구매 시</label>
                                    <input type="checkbox" id="lyr12_08" class="type01"><label for="lyr12_08"><span class="circle"></span>모든 상품 구매 시</label>
                                    <li class="type01 noti">지정한 상품 중 1개라도 구매하면 조건을 만족합니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>구매금액 조건</h4>
                                    <li>주문금액이</li>
                                    <input type="text" id="lyr12_09" placeholder="0"><label for="lyr12_09"></label>
                                    <span>이상일 경우 구매 가능</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 금액별 그룹 수정하기 -->
                        <div class="lyr_wrap lyr13">
                            <div class="lyr_top">
                                <li>금액별 그룹 수정하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <!-- 스페셜 이벤트 - 금액별 그룹 추가하기 -->
                                 <div class="lyr_cont" style="display:none;">
                                    <li class="type01 noti">혜택 노출 조건을 스페셜이벤트로 설정한 경우<br> 금액 기준은 설정한 스페셜 이벤트의 기준과 동일하게 적용됩니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>그룹 이름<span class="gray">29/30</span></h4></h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액 설정</h4>
                                    <div class="cont_wrap">
                                        <input class="type04" type="checkbox" id="lyr13_01"><label for="lyr13_01">총 구매금액</label>
                                        <input class="type04" type="checkbox" id="lyr13_02"><label for="lyr13_02">카테고리</label>
                                        <input class="type04" type="checkbox" id="lyr13_03"><label for="lyr13_03">브랜드</label>
                                        <input class="type04" type="checkbox" id="lyr13_04"><label for="lyr13_04">기획전/이벤트</label>
                                    </div>
                                </div>
                                <!-- 총 구매금액 선택시 노출 -->
                                <div class="lyr_cont">
                                    <h4>배송 구분</h4>
                                    <input type="checkbox" id="lyr13_05" class="type01"><label for="lyr13_05"><span class="circle"></span>전체상품</label>
                                    <input type="checkbox" id="lyr13_06" class="type01"><label for="lyr13_06"><span class="circle"></span>텐바이텐 배송 포함</label>
                                </div>
                                <!-- 카테고리 선택시 노출 -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>카테고리 지정<span>*복수선택 가능합니다.</span></h4>
                                    <select name="" id="">
                                        <option value="">1depth</option>
                                    </select>
                                    <select name="" id="">
                                        <option value="">2depth</option>
                                    </select>
                                    <button class="add btn_blue">카테고리 추가</button>
                                    <div class="option">
                                        <div class="option_added">
                                            <li>디자인문구</li>
                                            <li>></li>
                                            <li>다이어리/플래너</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                        <div class="option_added">
                                            <li>디자인문구</li>
                                            <li>></li>
                                            <li>데코레이션</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                    </div>
                                </div>
                                <!-- 브랜드 선택시 노출 -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>브랜드 지정<span>*복수선택 가능합니다.</span></h4>
                                    <div class="cont_wrap search brand">
                                         <input type="text" placeholder="브랜드ID를 입력해주세요">
                                        <button class="add btn_blue">추가하기</button>
                                        <button class="search btn_white">브랜드 찾기</button>
                                    </div>
                                </div>
                                <!-- 기획전/이벤트 선택시 노출 -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>기획전/이벤트 지정<span>*복수선택 가능합니다.</span></h4>
                                    <div class="cont_wrap search event">
                                        <input type="text" placeholder="이벤트 코드를 입력해주세요">
                                        <button class="add btn_blue">추가하기</button>
                                    </div>
                                </div>
                                <!-- 기획전/이벤트 선택시 노출 -->
                                 <div class="lyr_cont" style="display:block;">
                                    <h4>상품 구매 조건</h4>
                                    <input type="checkbox" id="lyr13_07" class="type01"><label for="lyr13_07"><span class="circle"></span>1개라도 구매 시</label>
                                    <input type="checkbox" id="lyr13_08" class="type01"><label for="lyr13_08"><span class="circle"></span>모든 상품 구매 시</label>
                                    <li class="type01 noti">지정한 상품 중 1개라도 구매하면 조건을 만족합니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>구매금액 조건</h4>
                                    <li>주문금액이</li>
                                    <input type="text" id="lyr13_09" placeholder="0"><label for="lyr13_09"></label>
                                    <span>이상일 경우 구매 가능</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 플러스세일 설정하기 -->
                        <div class="lyr_wrap lyr14">
                            <div class="lyr_top">
                                <li>플러스세일 설정하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>선택한 상품</h4>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p class="prd_code">12345678</p>
                                                        <p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% 할인<span>38,000</span></p>
                                                            <p class="price p_type03">7% 쿠폰<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                 <div class="lyr_cont">
                                    <h4>플러스 할인가 설정</h4>
                                    <div class="option t02">
                                        <ul>
                                            <li>할인가</li>
                                            <input type="text" id="lyr14_01"><label for="lyr14_01"></label>
                                        </ul>
                                        <ul class="percent">
                                            <li>할인율</li>
                                            <input type="text" id="lyr14_02"><label for="lyr14_02"></label>
                                        </ul>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>할인 시 공급율</h4>
                                    <select name="" id="">
                                        <option value="">텐바이텐부담</option>
                                        <option value="">텐바이텐부담</option>
                                    </select>
                                    <span>(판매가) 32,300</span>
                                    <span>(할인매입가) 23,902</span>
                                    <b>26%</b>
                                </div>
                                <div class="lyr_cont">
                                    <h4>사은품 수량 설정</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr14_03" class="type01" ><label for="lyr14_03"><span class="circle"></span>한정 수량<input type="text" placeholder="0">개</li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr14_04" class="type01"><label for="lyr14_04"><span class="circle"></span>비한정</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>최대 구매 수 제한</h4>
                                    <div class="option t03 lyr14_05">
                                        <input type="text" id="lyr14_05"><label for="lyr14_05"></label>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <button class="badge_add">
                                        <li>특별한 상품에 대한 한내를 위해<br>뱃지 및 유의사항 추가하기</li>
                                        <li class="arrow"></li>
                                    </button>
                                </div>
                                <!-- badge_add 뱃지 추가 버튼 클릭시 -->
                                <div class="lyr_cont badge">
                                    <div class="cont_wrap">
                                        <ul>
                                            <h4>상품 뱃지 설정<span class="gray">29/30</span></h4>
                                            <input type="text" placeholder="뱃지에 들어갈 문구를 입력해주세요">
                                            <h4 class="badge_noti">유의사항<span class="gray">29/30</span></h4>
                                            <textarea placeholder="상품 유닛 하단에 표시될 유의사항을입력해주세요"></textarea>
                                        </ul>
                                        <ul>
                                            <h4>뱃지와 유의사항 미리보기</h4>
                                            <li class="badge_img"><img src="https://webadmin.10x10.co.kr/images/icon/unit.png"></li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 사은품 신규등록 -->
                        <div class="lyr_wrap lyr15">
                            <div class="lyr_top">
                                <li>사은품 신규등록</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>사은품 구분</h4>
                                    <div class="cont_wrap type04">
                                        <input class="type04" type="checkbox" id="lyr15_01"><label for="lyr15_01">상품</label>
                                        <input class="type04" type="checkbox" id="lyr15_02"><label for="lyr15_02">보너스 쿠폰</label>
                                        <input class="type04" type="checkbox" id="lyr15_03"><label for="lyr15_03">마일리지</label>
                                    </div>
                                    <li class="noti">마일리지는 전체 증정 이벤트에서만 사용할 수 있습니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>사은품명</h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l" placeholder="사은품명을 입력해주세요">
                                    </div>
                                </div>
                                <!-- 상품 선택시 노출 -->
                                <div class="lyr15_1" style="display:none;">
                                    <div class="lyr_cont">
                                        <h4>배송방법</h4>
                                        <input type="checkbox" id="lyr15_04" class="type01"><label for="lyr15_04"><span class="circle"></span>텐바이텐 배송</label>
                                        <input type="checkbox" id="lyr15_05" class="type01"><label for="lyr15_05"><span class="circle"></span>업체배송</label>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>브랜드 ID</h4>
                                        <div class="cont_wrap search brand">
                                            <input type="text" placeholder="브랜드ID를 입력해주세요">
                                            <button class="add btn_blue">브랜드 ID 검색</button>
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>당첨상품코드</h4>
                                        <div class="cont_wrap search">
                                            <input type="text" class="input_l" placeholder="사은품명을 입력해주세요">
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>물류코드</h4>
                                        <div class="cont_wrap search code">
                                            <input type="text" class="code01">
                                            <input type="text" class="code02">
                                            <input type="text" class="code03">
                                            <button class="add btn_blue btn01">검색</button>
                                            <button class="add btn_blue btn02">사은품 물류코드 자동생성</button>
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>대표 이미지</h4>
                                        <li class="noti">사은품 구분을 위한 이미지로 활용되니 등록을 권장합니다.</li>
                                        <div class="cont_wrap img">
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                        </div>
                                    </div>
                                </div>
                                <!-- 보너스 쿠폰 선택시 노출 -->
                                <div class="lyr15_2" style="display:none;">
                                    <div class="lyr_cont">
                                        <h4>쿠폰코드</h4>
                                        <div class="cont_wrap search">
                                            <input type="text" class="input_l"placeholder="연결할 쿠폰 코드를 입력해주세요">
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>대표 이미지</h4>
                                        <li class="noti">사은품 구분을 위한 이미지로 활용됩니다.</li>
                                        <li class="type02">
                                            <input class="type02" type="checkbox" id="lyr15_06" ><label for="lyr15_06"><span class="circle"></span>기본 이미지</label>
                                            <div class="cont_wrap img">
                                                <img src="https://webadmin.10x10.co.kr/images/icon/coupon.png">
                                            </div>
                                        </li>
                                       <li class="type02">
                                            <input class="type02" type="checkbox" id="lyr15_07" ><label for="lyr15_07"><span class="circle"></span>이미지 직접 등록</label>
                                            <div class="cont_wrap img">
                                                <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                            </div>
                                       </li>
                                    </div>
                                </div>
                                <!-- 마일리지 선택시 노출 -->
                                <div class="lyr15_3" style="display:block;">
                                    <div class="lyr_cont">
                                        <h4>마일리지 지급 금액</h4>
                                        <div class="cont_wrap search">
                                            <input type="text" class="input_l"placeholder="금액을 입력해주세요">
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>유효기간 설정</h4>
                                        <ul>
                                            <li class="type02">
                                            <input class="type02" type="checkbox" id="lyr15_08" ><label for="lyr15_08"><span class="circle"></span>기간설정</label>
                                            </li>
                                            <li class="type02">
                                                <input class="type02" type="checkbox" id="lyr15_09" ><label for="lyr15_09"><span class="circle"></span>회수일자 지정</label>
                                            </li>
                                        </ul>
                                        <!-- 기간설정 -->
                                        <div class="option lyr15_08">
                                            <input type="text" placeholder="30"> 일 이후 마일리지 소멸 
                                        </div>
                                        <!-- 회수일자 지정 -->
                                        <div class="option">
                                            <div class="date_wrap">
                                                <ul class="date">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- 날짜 선택 -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2월 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    <input type="text" placeholder="00" class="time">:<input type="text" placeholder="00" class="time"><img src="https://webadmin.10x10.co.kr/images/icon/clock.png">
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>대표 이미지</h4>
                                        <li class="noti">사은품 구분을 위한 이미지로 활용됩니다.</li>
                                        <ul>
                                            <li class="type02">
                                                <input class="type02" type="checkbox" id="lyr15_10" ><label for="lyr15_10"><span class="circle"></span>기본 이미지</label>
                                                <div class="cont_wrap img">
                                                    <img src="https://webadmin.10x10.co.kr/images/icon/mileage.png">
                                                </div>
                                            </li>
                                            <li class="type02">
                                                <input class="type02" type="checkbox" id="lyr15_11" ><label for="lyr15_11"><span class="circle"></span>이미지 직접 등록</label>
                                                <div class="cont_wrap img">
                                                    <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                                </div>
                                            </li>
                                       </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 사은품 설정 -->
                        <div class="lyr_wrap lyr16">
                            <div class="lyr_top">
                                <li>사은품 설정</li>
                            </div>
                            <!-- 상품 -->
                            <div class="lyr_cont_wrap" style="display:none;">
                                <div class="lyr_cont">
                                    <h4>사은품 정보</h4>
                                    <div class="add_btn">
                                        <button>사은품 신규등록</button><span></span>
                                        <button>기존 사은품 불러오기</button>
                                    </div>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                        <a>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <ul>
                                                            <span>상품</span>
                                                            <p>12345678</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p class="prd_brand">dailylife</p>
                                                        <p class="prd_name">신년축하 감사카드 패키지패키지패키지</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p>업체배송업체배송업체배송</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                                </a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>구매 횟수 설정</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_01" ><label for="lyr16_01"><span class="circle"></span>한정수량 <input type="text" placeholder="0" > 개</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_02" ><label for="lyr16_02"><span class="circle"></span>비한정</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>사은품 이미지 관리</h4>
                                    <div class="cont_wrap img">
                                        <li>결제 페이지 썸네일</li>
                                        <ul>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                        </ul>
                                        <li>상세 팝업 이미지</li>
                                        <ul class="img_wrap">
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <!-- 보너스쿠폰 -->
                            <div class="lyr_cont_wrap" style="display:none;">
                                <div class="lyr_cont">
                                    <h4>사은품 정보</h4>
                                    <div class="add_btn">
                                            <button>사은품 신규등록</button><span></span>
                                            <button>기존 사은품 불러오기</button>
                                    </div>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                        <a>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <ul>
                                                            <span>보너스쿠폰</span>
                                                            <p>12345678</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p>10x10giftmall</p>
                                                        <p class="prd_name">텐바이텐이 쏜다! 배송비 무료 쿠폰!</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p>보너스쿠폰<br>3000원 지급</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                                </a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>구매 횟수 설정</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_03" ><label for="lyr16_03"><span class="circle"></span>한정수량 <input type="text" placeholder="0" > 개</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_04" ><label for="lyr16_04"><span class="circle"></span>비한정</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>사은품 이미지 관리</h4>
                                    <div class="cont_wrap img">
                                        <li>결제 페이지 썸네일</li>
                                        <ul>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <!-- 마일리지 -->
                            <div class="lyr_cont_wrap" style="display:block;">
                                <div class="lyr_cont">
                                    <h4>사은품 정보</h4>
                                    <div class="add_btn">
                                            <button>사은품 신규등록</button><span></span>
                                            <button>기존 사은품 불러오기</button>
                                    </div>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                        <a>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <ul>
                                                            <span>마일리지</span>
                                                            <p>12345678</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p>10x10giftmall</p>
                                                        <p class="prd_name">텐바이텐이 쏜다! 1,010지원금!</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p>마일리지<br>1,010P 지급</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                                </a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>구매 횟수 설정</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_05" ><label for="lyr16_05"><span class="circle"></span>한정수량 <input type="text" placeholder="0" > 개</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_06" ><label for="lyr16_06"><span class="circle"></span>비한정</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>사은품 이미지 관리</h4>
                                    <div class="cont_wrap img">
                                        <li>결제 페이지 썸네일</li>
                                        <ul>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>이미지 등록</button>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <script>
                </script>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->