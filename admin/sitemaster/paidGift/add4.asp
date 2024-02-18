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
                            <li class="step on"><span></span>최종점검</li>
                        </div>
                        <div class="step_wrap step04">
                            <div class="step_noti on"><span>저장하기 전 마지막 확인을 해주세요!</span></div>
                            <div class="bg">
                                <div class="step_cont step04_01 on">
                                    <h2>기본정보 입력</h2>
                                    <div class="step_detail">
                                        <h3>지원채널 선택</h3>
                                        <div class="step_detail_list">
                                            <input type="checkbox" class="type01" id="step04_01"><label for="step04_01"><span class="circle"></span>PC WEB</label>
                                            <input type="checkbox" class="type01" id="step04_02"><label for="step04_02"><span class="circle"></span>App(iOS/Android)</label>
                                            <input type="checkbox" class="type01" id="step04_03"><label for="step04_03"><span class="circle"></span>Mobile WEB</label>
                                        </div>
                                    </div>
                                    <div class="step_detail">
                                        <h3>제목<span>22/99</span></h3>
                                        <textarea>통합혜택관리를 런칭 기념 구매혜택 테스트</textarea>
                                    </div>
                                    <div class="step_detail">
                                        <h3>부제목<span>22/99</span></h3>
                                        <textarea>통합혜택관리를 런칭 기념 구매혜택 테스트</textarea>
                                    </div>
                                    <div class="step_detail">
                                        <h3>담당자</h3>
                                        <div class="step_detail_list">
                                            <li>
                                                <input type="text" value="마케팅팀 - 김텐텐">
                                                <a href="" class="close"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"></a>
                                            </li>
                                            <button class="btn_blue">변경</button>
                                        </div>
                                    </div>
                                    <div class="step_detail">
                                        <h3>사용여부</h3>
                                        <div class="step_detail_list">
                                            <input type="checkbox" class="type01" id="step04_04"><label for="step04_04"><span class="circle"></span>사용함</label>
                                            <input type="checkbox" class="type01" id="step04_05"><label for="step04_05"><span class="circle"></span>사용안함</label>
                                        </div>
                                    </div>
                                </div>
                                <div class="step_cont step04_02 on">
                                    <h2>주문/결제 화면 확인<li><span>2022.01.03 - 2022.01.20</span> 진행되는 구매혜택의 순서를 조정해주세요</li></h2>
                                    <div class="step_detail">
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box on">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>통합혜택관리를 런칭 기념 구매혜택 테스트</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="step_cont step04_03 on">
                                <h2>안내사항<li>자세히 안내하고자 하는 내용이 있다면 작성해주세요. 팝업으로 노출됩니다!</li></h2>
                                <div class="step_detail">
                                    <textarea>· 2022 다이어리 스토리 상품 포함 20,000원 이상 구매 시 증정됩니다.&#10;· 쿠폰, 할인카드 등 사용 후 구매확정 금액 기준입니다.&#10;· 2022 다이어리 스토리 상품 포함 20,000원 이상 구매 시 증정됩니다.&#10;· 쿠폰, 할인카드 등 사용 후 구매확정 금액 기준입니다.&#10;· 2022 다이어리 스토리 상품 포함 20,000원 이상 구매 시 증정됩니다.&#10;· 쿠폰, 할인카드 등 사용 후 구매확정 금액 기준입니다.&#10;
                                    </textarea>
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
                </div>
                <script>
                </script>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->