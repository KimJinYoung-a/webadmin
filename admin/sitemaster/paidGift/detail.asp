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
                <div class="content paidGift detail">
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
                            <div class="add_btn">
                                <button>+ 구매혜택 만들기</button>
                            </div>
                            <div class="list_top">
                                <li>총 <span>103</span>건</li>
                            </div>
                            <div class="list_cont">
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
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[플러스세일] 1월 첫구매 플러스세일</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>노출조건 : <span>조건A, 조건B</span></li>
                                        <li>담당부서 : <span>마케팅</span></li>
                                        <li class="state st04">사용안함</li>
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
                        <div class="section_top">
                            <li>구매혜택 번호 : <span>1234</span></li>
                            <div>
                                <button class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png" alt="">정렬순서 편집하기</button>
                                <button><img src="https://webadmin.10x10.co.kr/images/icon/trash_gray.png" alt="">삭제하기</button>
                            </div>
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <li class="dt_type">플러스세일</li>
                                <button class="btn_edit">수정하기</button>
                                <div class="dt_list_wrap">
                                    <textarea class="title01" rows="1">1월 첫구매 플러스 세일</textarea>
                                    <textarea class="title02" rows="1">1월에도 텐텐을 찾아주신 첫 구매 회원들에게만 드리는 특별한 선물</textarea>
                                </div>
                                <div class="dt_list_wrap">
                                    <li>운영기간</li>
                                     <div class="date_wrap">
                                        <!-- 선택 시 date start on / date end on -->
                                        <ul class="date start">
                                        <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png"></button>
                                        <!-- 날짜 선택 -->
                                            <div class="cal_month t02" style="display:none;">
                                                <div class="arrow"></div>
                                                    <table class="table-condensed table-bordered table-striped">
                                                            <thead>
                                                                <tr>
                                                                    <th colspan="7">
                                                                        <ul class="btn_group">
                                                                            <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></a>
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
                                        <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
                                        <ul class="date end">
                                            <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png"></button>
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
                                    <span class="line"></span>
                                    <li>사용여부</li>
                                    <select>
                                        <option>사용함</option>
                                        <option>사용안함</option>
                                    </select>
                                    <span class="line"></span>
                                    <li>담당자</li>
                                    <li>
                                        <input type="text" value="마케팅팀 - 김텐텐" class="name">
                                    </li>
                                </div>
                            </div> 
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <h3 class="dt_list_top">혜택 노출 대상자<span class="noti blue">항목을 선택하면 수정 또는 추가할 수 있어요</span></h3>
                                <div class="dt_list_wrap">
                                    <div class="info_wrap_top">
                                        <button>
                                            <li>모든 구매자</li>
                                            <li class="check">&#10003;</li>
                                        </button>
                                         <button class="on">
                                            <li>특정 구매자</li>
                                            <li class="check">&#10003;</li>
                                            <span class="arrow"></span>
                                        </button>
                                    </div>
                                    <div class="info_wrap_list">
                                        <li class="bold">혜택 노출조건 <span>2</span></li>
                                        <div class="btn_wrap">
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
                                    <!-- 스페셜 이벤트 -->
                                    <div class="info_wrap_list" style="display:block;">
                                        <li class="bold">스페셜 이벤트</li>
                                        <div class="btn_wrap">
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
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <h3 class="dt_list_top">혜택 설정</h3>
                                <!-- 단순그룹설정 시에만 노출 -->
                                <button class="btn_edit">그룹 순서 수정하기</button>
                                <div class="dt_list_wrap">
                                    <div class="info_wrap_top">
                                        <div class="btn_wrap t02">
                                            <button>그룹없이 상품 등록</button>
                                            <button class="on">단순 그룹 설정</button>
                                            <button>금액별 그룹 설정</button>
                                        </div>
                                    </div>
                                    <div class="info_wrap_list t02">
                                        <!-- 플러스세일 -->
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
                                                            <span>
                                                                + 상품 추가하기
                                                            </span>
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
                                        <!-- 무료사은품 -->
                                        <!-- 그룹없이 상품 등록 -->
                                        <div class="step_detail group group01_02" style="display:none;">
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
                                                <button class="added on"><span class="sort">ㅣㅣㅣ</span><li>다이어리에는 뭐니뭐니해도 스티커지</li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button class="added"><span class="sort">ㅣㅣㅣ</span><li>치트키는 떡메모지</li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button><li>+그룹 추가하기</li></button>
                                            </div>
                                            <!-- 그롭없이 상품 등록과 동일<div class="table_wrap"></div> -->
                                        </div>
                                        <!-- 금액별 그룹 설정 -->
                                        <div class="step_detail group group03_02" style="display:none;">
                                            <div class="tab">
                                                <button class="added on"><li>10,000원~<span>&#65372;</span><span class="gr_cond">구매금액</span><span class="gr_name">10,000원 이상 구매했다면</span></li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
                                                <button class="added"><li>10,000원~<span>&#65372;</span><span class="gr_cond">구매금액</span><span class="gr_name">10,000원 이상 구매했다면</span></li><img src="https://fiximage.10x10.co.kr/web2021/anniv2021/icon_edit.png" alt=""></button>
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
                            </div> 
                        </div>
                        <div class="step_wrap">
                            <div class="step_cont detail on">
                                <h3 class="dt_list_top">자세히 보기</h3>
                                <button class="btn_edit">수정하기</button>
                                <li class="noti gray">자세히 안내하고자 하는 내용이 있다면 작성해주세요!</li>
                                <div class="dt_list_wrap">
                                    <div class="info_wrap_list t03">
                                        <textarea>· 2022 다이어리 스토리 상품 포함 20,000원 이상 구매 시 증정됩니다.&#10;· 쿠폰, 할인카드 등 사용 후 구매확정 금액 기준입니다.&#10;· 2022 다이어리 스토리 상품 포함 20,000원 이상 구매 시 증정됩니다.&#10;· 쿠폰, 할인카드 등 사용 후 구매확정 금액 기준입니다.&#10;· 2022 다이어리 스토리 상품 포함 20,000원 이상 구매 시 증정됩니다.&#10;· 쿠폰, 할인카드 등 사용 후 구매확정 금액 기준입니다.&#10;
                                        </textarea>
                                    </div>
                                </div>
                            </div> 
                        </div>
                    </div>
                    <!-- 레이어팝업 노출 : paidGift_lyr on -->
                    <div class="lyr paidGift_lyr">
                        <div class="lyr_overlay"></div>
                     <!-- lyr_wrap lyr01 on -->
                        <!-- 구매혜택 정렬 순서 편집하기 -->
                        <div class="lyr_wrap lyr17 on">
                            <div class="lyr_top">
                                <li>구매혜택 정렬 순서 편집하기</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <li class="noti"><span>2022.01.03 - 2022.01.20</span> 진행되는 구매혜택의 순서를 조정해주세요</li>
                                    <div class="sort_wrap">
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box on">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>통합혜택관리를 런칭 기념 구매혜택 테스트</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img></a>
                                            <ul>
                                                <li>산리오 런칭기념 블라</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">구매금액 도달</li>
                                            <ul>
                                        </div>
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