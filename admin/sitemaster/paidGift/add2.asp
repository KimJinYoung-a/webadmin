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
             <link rel="stylesheet" type="text/css"
                href="https://uicdn.toast.com/tui-calendar/latest/tui-calendar.css" />
            <link rel="stylesheet" type="text/css"
                href="https://uicdn.toast.com/tui.date-picker/latest/tui-date-picker.css" />
            <link rel="stylesheet" type="text/css"
                href="https://uicdn.toast.com/tui.time-picker/latest/tui-time-picker.css" />
            <script src="https://uicdn.toast.com/tui.code-snippet/v1.5.2/tui-code-snippet.min.js"></script>
            <script src="https://uicdn.toast.com/tui.time-picker/latest/tui-time-picker.min.js"></script>
            <script src="https://uicdn.toast.com/tui.date-picker/latest/tui-date-picker.min.js"></script>
            <script src="https://uicdn.toast.com/tui-calendar/latest/tui-calendar.js"></script>
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
                            <li class="step"><span></span>혜택설정</li>
                            <li class="step"><span></span>최종점검</li>
                        </div>
                        <div class="step_wrap step02">
                            <div class="step_noti on"><span>캘린더 날짜 선택 영역에서 구매혜택 시작일과 종료일을 지정해주세요.</span>
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
                            </div>
                            <div class="step_cont step02 on">
                                <div id="calendar">
                                <!-- 캘린더 날짜 클릭 텍스트 tui-full-calendar-month-creation-guide 우측 상단에 위치-->
                                <p class="click_text click_st" style="display:none;">00:00 시작&nbsp;<span></span></p>
                                <p class="click_text click_sel" style="display:none;">종료일 선택</p>
                                <p class="click_text click_end" style="display:none;">00:00 시작&nbsp;<span></span></p>
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

                    var calendar = new tui.Calendar(document.getElementById('calendar'), {
                        defaultView: 'month',
                        // useDetailPopup: true,
                        // isReadOnly: true,
                        timezones: {
                            timezoneOffset: 540,
                            displayLabel: 'GMT+09:00',
                            tooltip: 'Seoul'
                        },
                        template: {
                            monthGridHeader: function (model) {
                                var date = new Date(model.date);
                                var day = date.getDate();
                                var format = ("00" + day.toString()).slice(-2);
                                var template = '<span class="tui-full-calendar-weekday-grid-date">' + format + '</span>';
                                return template;
                            }
                        },
                        month: {
                            daynames: ['일', '월', '화', '수', '목', '금', '토'],
                            startDayOfWeek: 0,
                        },
                    });

                    calendar.createSchedules([
                        {
                            id: '1',
                            calendarId: '1',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-02T22:30:00+09:00',
                            end: '2022-02-03T02:30:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:87%;">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>    
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:59%;">    
                                            <li class="percent">
                                                <span>59%</span>    
                                                <span>2,902/4,000</span>    
                                            </li> 
                                        </div>
                                    </ul> 
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar" style="display:none;">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>        
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>      
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>       
                                            </li> 
                                        </div>
                                    </ul>
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '1',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-02T22:30:00+09:00',
                            end: '2022-02-03T02:30:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '2',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-28T17:30:00+09:00',
                            end: '2022-03-01T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '2',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-09T17:30:00+09:00',
                            end: '2022-02-27T17:31:00+09:00',
                            state:`<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '2',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-09T17:30:00+09:00',
                            end: '2022-02-27T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '1',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-09T17:30:00+09:00',
                            end: '2022-02-11T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '2',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-18T17:30:00+09:00',
                            end: '2022-02-20T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '3',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-22T17:30:00+09:00',
                            end: '2022-02-26T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '3',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-22T17:30:00+09:00',
                            end: '2022-02-26T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        },
                        {
                            id: '1',
                            calendarId: '3',
                            title: '[전체사은품] 20주년 머그컵 - 마케팅',
                            category: 'time',
                            start: '2022-02-27T17:30:00+09:00',
                            end: '2022-03-01T17:31:00+09:00',
                            state: `<div class="cal_popup">
                                <div class="cal_popup_top">
                                    <span>플러스세일</span>
                                    <span class="button_wrap">
                                        <button class="detail">자세히 보기</button>
                                        <button class="close"><img src="https://fiximage.10x10.co.kr/web2017/common/btn_lyr_close.png" alt=""></button>
                                    </span>
                                </div>
                                <h3>1월 첫 구매 고객에게만 선물을 드릴게요!</h3>
                                <div class="cal_popup_cont">
                                    <h4>20,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">2022 데일리라이크 무드 다이어리</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black" style="width:23%;">    
                                            <li class="percent">
                                                <span>23%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                                <div class="cal_popup_cont">
                                    <h4>40,000원 이상 구매 시</h4>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul>
                                    <ul>
                                        <li class="prd_img"><img src="" alt=""></li>
                                        <li class="prd_name">시리얼 컵 스푼 세트</li>
                                        <div class="bar">
                                            <li class="bar_gray">   
                                            <li class="bar_black">    
                                            <li class="percent">
                                                <span>87%</span>    
                                                <span>2,902/4,000</span>     
                                            </li> 
                                        </div>
                                    </ul> 
                                </div>
                            </div>`
                        }
                    ]);

                    function onClickTodayBtn() {
                        calendar.today();
                    }
                    function moveToNextOrPrevRange(val) {
                        if (val === -1) {
                            calendar.prev();
                        } else if (val === 1) {
                            calendar.next();
                        }
                    }
                    calendar.setCalendarColor('1', {
                        color: '#687182',
                        bgColor: '#e9edf5',
                        borderColor: '#687182'
                    });
                    calendar.setCalendarColor('2', {
                        color: '#14804a',
                        bgColor: '#e1fcef',
                        borderColor: '#14804a'
                    });
                    calendar.setCalendarColor('3', {
                        color: '#c97a20',
                        bgColor: '#fcf2e6',
                        borderColor: '#c97a20',
                    });

                    calendar.setTheme({
                        'month.day.fontSize': '12px',
                        'month.schedule.height': '20px',
                        'common.holiday.color': '#333',
                        'month.holidayExceptThisMonth.color': 'rgba(51, 51, 51, 0.4)',
                    })
                </script>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->