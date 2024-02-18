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
                            <li class="step"><span></span>기간설정</li>
                            <li class="step"><span></span>혜택설정</li>
                            <li class="step"><span></span>최종점검</li>
                        </div>
                        <div class="step_wrap step01">
                            <div class="step_noti on"><span>가장 먼저 주문결제 페이지에서 구매혜택을 노출할 대상자를 지정합니다. 어떤 구매자에게 혜택을 제공하고 싶나요?</span></div>
                            <div class="btn_group">
                                <button class="type01 step01_01 on"><span>모든 구매자</span><p class="img"></p></button>
                                <button class="type01 step01_02"><span>특정 구매자</span><p class="img"></p></button>
                            </div>
                            <!-- 모든 구매자 step_cont step01_01 on -->
                            <div class="step_cont step01_01">
                                <li>모든 구매자에게 구매혜택을 제공합니다. 아래 '다음' 버튼을 눌러주세요.</li>
                            </div>
                            <!-- 모든 구매자 step_cont step01_02 on -->
                            <div class="step_cont step01_02 on">
                                <li>선택된 조건에 모두 해당하는 구매자에게만 혜택을 제공합니다. 조건을 정한 뒤 '다음' 버튼을 눌러주세요.</li>
                                <div class="step_detail">
                                    <div class="step_detail_top">
                                        <h3>조건 설정하기</h3>
                                        <span>*조건은 복수 선택 가능하며, 스페셜이벤트와 함께 선택할 수 없습니다.</span>
                                    </div>
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>생애 첫 결제</h3>
                                            <li>가입 후 첫 결제일 경우</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>최근 결제금액 기준</h3>
                                            <li>최근 5개월 내 배송완료된 주문건의 총 결제금액</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>회원등급 해당</h3>
                                            <li>매월 갱신되는 회원등급 기준</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>구매금액 도달</h3>
                                            <li>현재 주문 건의 구매금액 기준</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>구매횟수 도달</h3>
                                            <li>최근 5개월 내 배송완료된 주문횟수 기준</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>지정한 상품 구매</h3>
                                            <li>특정 상품 구매 시</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>지정한 카테고리 상품 구매</h3>
                                            <li>현재 또는 과거 주문 건에서 특정 카테고리 상품 구매 시</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>지정한 브랜드 상품 구매</h3>
                                            <li>현재 또는 과거 주문 건에서 특정 브랜드 상품 구매시</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>지정한 기획전/이벤트 상품 구매</h3>
                                            <li>특정 기획전/이벤트 상품 구매 시</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                                <div class="step_detail">
                                    <div class="step_detail_top">
                                        <h3>스페셜 이벤트</h3>
                                        <span>*1개만 선택 가능하며, 조건 설정하기와 함께 선택할 수 없습니다.</span>
                                         <li>직접 설정할 수 없는 특별한 조건이거나, 사용빈도가 높은 구매혜택은 커스텀 항목으로 지정합니다.</li>
                                    </div>
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>다이어리 스토리</h3>
                                            <li>다이어리 스토리 사은품 지급 대상</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>주년이벤트</h3>
                                            <li>N 주년이벤트 사은품 지급 대상</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="steps_bottom">
                            <ul>
                                <button class="delete">취소</button>
                                <!-- 버튼 비활성화 : next on -->
                                <button class="next on">다음</button>
                            </ul>
                        </div>
                    </div>

                    <!-- 레이어팝업 노출 : paidGift_lyr on -->
                    <div class="lyr paidGift_lyr">
                        <div class="lyr_overlay"></div>
                     <!-- 혜택별 레이어팝업 노출 : lyr_wrap lyr01 on -->
                        <!-- 최근 결제금액 기준 -->
                        <div class="lyr_wrap lyr01 on">
                            <div class="lyr_top">
                                <li>최근 결제금액 기준</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>기간 설정</h4>
                                    <input type="checkbox" id="lyr01_01" class="type01"><label for="lyr01_01"><span class="circle"></span>최근 5개월 동안</label>
                                    <input type="checkbox" id="lyr01_02" class="type01"><label for="lyr01_02"><span class="circle"></span>기간 직접설정</label>
                                    <!-- 기간 직접설정 선택시 노출 -->
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
                                                    </div></button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"><button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
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
                                                    </div></button>
                                                </ul>
                                            </div>
                                        <li class="noti">최근 5개월 이내의 날짜만 선택할 수 있습니다.</li>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액 설정</h4>
                                    <li>선택한 기간동안 결제한 금액의 합이</li>
                                    <input type="text" id="lyr01_03" placeholder="0"><label for="lyr01_03"></label>
                                    <span>이상 구매한 경우 혜택 노출</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 회원등급 -->
                        <div class="lyr_wrap lyr02">
                            <div class="lyr_top">
                                <li>회원등급</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <input type="checkbox" id="lyr02_01" class="type01"><label for="lyr02_01"><span class="circle"></span>WHITE</label>
                                    <input type="checkbox" id="lyr02_02" class="type01"><label for="lyr02_02"><span class="circle"></span>RED</label>
                                    <input type="checkbox" id="lyr02_03" class="type01"><label for="lyr02_03"><span class="circle"></span>VIP</label>
                                    <input type="checkbox" id="lyr02_04" class="type01"><label for="lyr02_04"><span class="circle"></span>VIP GOLD</label>
                                    <input type="checkbox" id="lyr02_05" class="type01"><label for="lyr02_05"><span class="circle"></span>VVIP</label>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 구매금액 도달 -->
                        <div class="lyr_wrap lyr03">
                            <div class="lyr_top">
                                <li>구매금액 도달</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>배송 구분</h4>
                                    <input type="checkbox" id="lyr03_01" class="type01"><label for="lyr03_01"><span class="circle"></span>전체상품</label>
                                    <input type="checkbox" id="lyr03_02" class="type01"><label for="lyr03_02"><span class="circle"></span>텐텐배송상품 포함</label>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액 설정</h4>
                                    <li>현재 주문건의 구매금액이</li>
                                    <input type="text" id="lyr01_03" placeholder="0"><label for="lyr01_03"></label>
                                    <span>이상 이상일 경우 노출</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 구매 횟수 -->
                        <div class="lyr_wrap lyr04">
                            <div class="lyr_top">
                                <li>구매 횟수</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>기간 설정</h4>
                                    <input type="checkbox" id="lyr04_01" class="type01"><label for="lyr04_01"><span class="circle"></span>최근 5개월 동안</label>
                                    <input type="checkbox" id="lyr04_02" class="type01"><label for="lyr04_02"><span class="circle"></span>기간 직접설정</label>
                                    <!-- 기간 직접설정 선택시 노출 -->
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
                                                    </button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
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
                                                    </button>
                                                </ul>
                                            </div>
                                        <li class="noti">최근 5개월 이내의 날짜만 선택할 수 있습니다.</li>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>구매 횟수 설정</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr04_03" ><label for="lyr04_03"><span class="circle"></span><input type="text" placeholder="0" > 회 주문한 고객 대상</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr04_04" ><label for="lyr04_04"><span class="circle"></span><input type="text" placeholder="0" > 회 이상 주문한 고객 대상</label></li>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 지정한 상품 구매 -->
                        <div class="lyr_wrap lyr05">
                            <div class="lyr_top">
                                <li>상품 선택하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>상품 선택<span>*복수선택 가능합니다.</span></h4>
                                    <div class="cont_wrap search">
                                        <input type="text" placeholder="상품코드를 입력해주세요">
                                        <button class="add btn_blue">추가하기</button>
                                        <button class="search btn_white">상품 찾기</button>
                                    </div>
                                </div>
                                <!-- 상품추가 시 노출 -->
                                <div class="lyr_cont">
                                    <div class="ttop">
                                        <span>총 2건</span>
                                        <button>선택항목 삭제</button>
                                    </div>
                                    <div class="table">
                                        <ul class="thead">
                                            <li><input type="checkbox"></li>
                                            <li>상품코드</li>
                                            <li>대표이미지/브랜드/상품명</li>
                                            <li></li>
                                            <li>판매가</li>
                                            <li></li>
                                        </ul>
                                    <div class="tbody_wrap">
                                            <ul class="tbody">
                                                <li><input type="checkbox"></li>
                                                <li>12345678</li>
                                                <li>
                                                    <img class="prd_img">
                                                    <ul><p class="prd_brand">dailylike</p>
                                                    <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                </li>
                                                <li>7% 할인</li>
                                                    <li>
                                                    <p class="price01">38,000</p>
                                                    <p class="price02">32,300</p>
                                                </li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                            </ul>
                                        </div>
                                        <div class="tbody_wrap">
                                            <ul class="tbody">
                                                <li><input type="checkbox"></li>
                                                <li>12345678</li>
                                                <li>
                                                    <img class="prd_img" >
                                                    <ul><p class="prd_brand">dailylike</p>
                                                    <p class="prd_name">신년축하 감사카드 패키지</p></ul>
                                                </li>
                                                <li>7% 할인</li>
                                                    <li>
                                                    <p class="price01">38,000</p>
                                                    <p class="price02">32,300</p>
                                                </li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </ul>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>상품 구매 조건</h4>
                                    <input type="checkbox" id="lyr05_01" class="type01"><label for="lyr05_01"><span class="circle"></span>1개라도 구매 시</label>
                                    <input type="checkbox" id="lyr05_02" class="type01"><label for="lyr05_02"><span class="circle"></span>모든 상품 구매 시</label>
                                    <li class="type01 noti">지정한 상품 중 1개라도 구매하면 조건을 만족합니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>수량 조건</h4>
                                     <!-- 구매수량 설정 선택시 노출 -->
                                    <div class="option t02">
                                        <li>지정한 상품을</li>
                                        <input type="text" id="lyr05_05" value="1"><label for="lyr05_05"></label>
                                        <span>이상 구매할 경우 조건 만족</span>
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
                        <!-- 지정한 카테고리 상품 구매 -->
                        <div class="lyr_wrap lyr06">
                            <div class="lyr_top">
                                <li>지정한 카테고리 상품 구매 시</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <div class="cont_wrap type03">
                                        <input class="type03" type="checkbox" id="lyr06_01"><label for="lyr06_01">현재 주문건 기준</label>
                                        <input class="type03" type="checkbox" id="lyr06_02"><label for="lyr06_02">과거 주문건 기준<li>최근 5개월 이내</li></label>
                                    </div>
                                </div>
                                <!-- 과거 주문건 기준 선택시 노출 -->
                                 <div class="lyr_cont">
                                        <h4>기간 설정</h4>
                                        <input type="checkbox" id="lyr06_03" class="type01"><label for="lyr06_03"><span class="circle"></span>최근 5개월 동안</label>
                                        <input type="checkbox" id="lyr06_04" class="type01"><label for="lyr06_04"><span class="circle"></span>기간 직접설정</label>
                                        <!-- 기간 직접설정 선택시 노출 -->
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
                                                    </button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
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
                                                    </button>
                                                </ul>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>카테고리 지정<span>*복수선택 가능합니다.</span></h4>
                                        <select name="" id="">
                                            <option value="">1depth</option>
                                        </select>
                                        <select name="" id="">
                                            <option value="">2depth</option>
                                        </select>
                                        <button class="add btn_blue">카테고리 추가</button>
                                         <!-- 카테고리 추가시 노출 -->
                                        <div class="option">
                                            <div class="option_added">
                                                <li>디자인문구</li>
                                                <li>></li>
                                                <li>다이어리/플래너</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                            <div class="option_added">
                                                <li>디자인문구</li>
                                                <li>></li>
                                                <li>데코레이션</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액/수량 조건</h4>
                                    <div class="cont_wrap lyr06">
                                        <input type="checkbox" id="lyr06_06" class="type01"><label for="lyr06_06"><span class="circle"></span>최소금액 설정</label>
                                        <input type="checkbox" id="lyr06_07" class="type01"><label for="lyr06_07"><span class="circle"></span>구매수량 설정</label>
                                        <!-- 과거 주문건 기준 선택시 노출 -->
                                        <input type="checkbox" id="lyr06_08" class="type01"><label for="lyr06_08"><span class="circle"></span>주문횟수 설정</label>
                                    </div>
                                    <!-- 설정 안함 선택시 노출 -->
                                    <div class="option">
                                        <li class="noti">지정한 카테고리 상품 중 1개라도 구매하면 조건을 만족합니다.</li>
                                    </div>
                                    <!-- 최소금액 설정 선택시 노출 -->
                                    <div class="option t02 lyr06_06">
                                        <li>지정한 카테고리 상품을</li>
                                        <input type="text" id="lyr06_09" placeholder="0" value="0"><label for="lyr06_09"></label>
                                        <span>이상 구매할 경우 혜택 노출</span>
                                    </div>
                                    <!-- 구매수량 설정 선택시 노출 -->
                                    <div class="option t02 lyr06_07">
                                        <li>지정한 카테고리 상품을</li>
                                        <input type="text" id="lyr06_10" value="1"><label for="lyr06_10"></label>
                                        <span>이상 구매할 경우 혜택 노출</span>
                                    </div>
                                    <!-- 주문횟수 설정 선택시 노출 -->
                                    <div class="option t02 lyr06_08">
                                        <li>선택한 기간동안 지정한 카테고리 상품을 포함하여</li>
                                        <input type="text" id="lyr06_11" value="0"><label for="lyr06_11"></label>
                                        <span>이상 주문한 경우 혜택 노출</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <!-- 버튼 비활성화 : submit on -->
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 지정한 브랜드 상품 구매 -->
                        <div class="lyr_wrap lyr07">
                            <div class="lyr_top">
                                <li>지정한 브랜드 상품 구매 시</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <div class="cont_wrap type03">
                                        <input class="type03" type="checkbox" id="lyr07_01"><label for="lyr07_01">현재 주문건 기준</label>
                                        <input class="type03" type="checkbox" id="lyr07_02"><label for="lyr07_02">과거 주문건 기준<li>최근 5개월 이내</li></label>
                                    </div>
                                </div>
                                <!-- 과거 주문건 기준 선택시 노출 -->
                                 <div class="lyr_cont">
                                        <h4>기간 설정</h4>
                                        <input type="checkbox" id="lyr07_03" class="type01"><label for="lyr07_03"><span class="circle"></span>최근 5개월 동안</label>
                                        <input type="checkbox" id="lyr07_04" class="type01"><label for="lyr07_04"><span class="circle"></span>기간 직접설정</label>
                                        <!-- 기간 직접설정 선택시 노출 -->
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
                                                    </button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
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
                                                    </button>
                                                </ul>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>브랜드 선택<span>*복수선택 가능합니다.</span></h4>
                                        <div class="cont_wrap search">
                                            <input type="text" placeholder="브랜드ID를 입력해주세요">
                                            <button class="add btn_blue">추가하기</button>
                                            <button class="search btn_white">브랜드 찾기</button>
                                        </div>
                                         <!-- 브랜드 추가시 노출 -->
                                        <div class="option">
                                            <div class="option_added">
                                                <li>PEANUTS10X10</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                            <div class="option_added">
                                                <li>SANRIO</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액/수량 조건</h4>
                                    <div class="cont_wrap lyr07">
                                        <input type="checkbox" id="lyr07_06" class="type01"><label for="lyr07_06"><span class="circle"></span>최소금액 설정</label>
                                        <input type="checkbox" id="lyr07_07" class="type01"><label for="lyr07_07"><span class="circle"></span>구매수량 설정</label>
                                        <!-- 과거 주문건 기준 선택시 노출 -->
                                        <input type="checkbox" id="lyr07_08" class="type01"><label for="lyr07_08"><span class="circle"></span>주문횟수 설정</label>
                                    </div>
                                    <!-- 설정 안함 선택시 노출 -->
                                    <div class="option">
                                        <li class="noti">지정한 브랜드 상품 중 1개라도 구매하면 조건을 만족합니다.</li>
                                    </div>
                                    <!-- 최소금액 설정 선택시 노출 -->
                                    <div class="option t02 lyr07_06">
                                        <li>지정한 브랜드 상품을</li>
                                        <input type="text" id="lyr07_09" value="0"><label for="lyr07_09"></label>
                                        <span>이상 구매할 경우 조건 노출</span>
                                    </div>
                                    <!-- 구매수량 설정 선택시 노출 -->
                                    <div class="option t02 lyr07_07">
                                        <li>지정한 브랜드 상품을</li>
                                        <input type="text" id="lyr07_10" value="1"><label for="lyr07_10"></label>
                                        <span>이상 구매할 경우 조건 노출</span>
                                    </div>
                                    <!-- 주문횟수 설정 선택시 노출 -->
                                    <div class="option t02 lyr07_08">
                                        <li>선택한 기간동안 지정한 브랜드 상품을 포함하여</li>
                                        <input type="text" id="lyr07_11" value="1"><label for="lyr07_11"></label>
                                        <span>이상 주문한 경우 혜택 노출</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <!-- 버튼 비활성화 : submit on -->
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 지정한 기획전/이벤트 상품 구매 -->
                        <div class="lyr_wrap lyr08">
                            <div class="lyr_top">
                                <li>지정한 기획전/이벤트 상품 구매 시</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>기획전/이벤트 선택<span>*복수선택 가능합니다.</span></h4>
                                        <div class="cont_wrap search">
                                            <input type="text" placeholder="이벤트 코드를 입력해주세요">
                                            <button class="add btn_blue">추가하기</button>
                                        </div>
                                         <!-- 이벤트코드 추가시 노출 -->
                                        <div class="option">
                                            <div class="option_added">
                                                <li class="e_img"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></li>
                                                <ul>
                                                    <li class="e_code">12345678</li>
                                                    <li class="e_name">신년축하 감사카드 패키지</li>
                                                </ul>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                             <div class="option_added">
                                                <li class="e_img"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></li>
                                                <ul>
                                                    <li class="e_code">12345678</li>
                                                    <li class="e_name">신년축하 감사카드 패키지</li>
                                                </ul>
                                                 <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>상품 구매 조건</h4>
                                        <input type="checkbox" id="lyr08_01" class="type01"><label for="lyr08_01"><span class="circle"></span>1개라도 구매 시</label>
                                        <input type="checkbox" id="lyr08_02" class="type01"><label for="lyr08_02"><span class="circle"></span>모든 상품 구매 시</label>
                                    <li class="noti">지정한 브랜드 상품 중 1개라도 구매하면 조건을 만족합니다.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>금액/수량 조건</h4>
                                        <input type="checkbox" id="lyr08_04" class="type01"><label for="lyr08_04"><span class="circle"></span>최소금액 설정</label>
                                        <input type="checkbox" id="lyr08_05" class="type01"><label for="lyr08_05"><span class="circle"></span>구매수량 설정</label>
                                    <!-- 최소금액 설정 선택시 노출 -->
                                    <div class="option t02 lyr08_04">
                                        <li>지정한 기획전 상품을</li>
                                        <input type="text" id="lyr08_06" value="0"><label for="lyr08_06"></label>
                                        <span>이상 구매할 경우 조건 만족</span>
                                    </div>
                                    <!-- 구매수량 설정 선택시 노출 -->
                                    <div class="option t02 lyr08_05">
                                        <li>지정한 기획전 상품을</li>
                                        <input type="text" id="lyr08_07" value="0"><label for="lyr08_07"></label>
                                        <span>이상 구매할 경우 조건 만족</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">취소</button>
                                    <!-- 버튼 비활성화 : submit on -->
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                        <!-- 구매금액 도달 -->
                        <div class="lyr_wrap lyr09">
                            <div class="lyr_top">
                                <li>구매금액 도달</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>배송 구분</h4>
                                        <input type="checkbox" id="lyr09_01" class="type01"><label for="lyr09_01"><span class="circle"></span>전체상품</label>
                                        <input type="checkbox" id="lyr09_02" class="type01"><label for="lyr09_02"><span class="circle"></span>텐텐배송상품 포함</label>
                                </div>
                                <div class="lyr_cont lyr09_03">
                                    <h4>금액 설정</h4>
                                    <input type="text" id="lyr09_03"><label for="lyr09_03"></label>
                                    <span>이상일 경우 노출</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <button class="initial">설정 초기화</button>
                                <ul>
                                    <button class="cancel">취소</button>
                                    <!-- 버튼 비활성화 : submit on -->
                                    <button class="submit">확인</button>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->