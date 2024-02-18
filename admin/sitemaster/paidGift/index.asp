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
            <link rel="stylesheet" type="text/css" href="/css/commonV20.css"/>
            <div>
                <!-- 구매혜택관리 -->
                <div class="content paidGift">
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
                                <button>+ 구매혜택 만들기</button>
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
                                </div>
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
                        <div class="menu">
                            <div id="renderRange" class="render-range">2022.02</div>
                            <!-- 2022.02 클릭시 나오는 캘린더 -->
                            <div class="cal_month" style="display:none;">
                                <div class="arrow"></div>
                                        <table class="table-condensed table-bordered table-striped">
                                            <thead>
                                                <tr>
                                                    <th colspan="7">
                                                        <ul class="btn_group">
                                                            <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></a>
                                                            <li class="btn active">2022</li>
                                                            <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                        </ul>
                                                    </th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <tr>
                                                    <td>1월</td>
                                                    <td class="on">2월</td>
                                                    <td>3월</td>
                                                    <td>4월</td>
                                                </tr>
                                                <tr>
                                                    <td>5월</td>
                                                    <td>6월</td>
                                                    <td>7월</td>
                                                    <td>8월</td>
                                                </tr>
                                                <tr>
                                                    <td>9월</td>
                                                    <td>10월</td>
                                                    <td>11월</td>
                                                    <td>12월</td>
                                                </tr>
                                            </tbody>
                                        </table>
                            </div>
                            <li class="menu-navi">
                                <button type="button" class="btn btn-default btn-sm move-today" data-action="move-today"
                                    onclick="onClickTodayBtn();">오늘</button>
                                <button type="button" class="btn btn-default btn-sm move-day" data-action="move-prev"
                                    onclick="moveToNextOrPrevRange(-1);">
                                    <i class="calendar-icon ic-arrow-line-left" data-action="move-prev">
                                        </i>
                                </button>
                                <button type="button" class="btn btn-default btn-sm move-day" data-action="move-next"
                                    onclick="moveToNextOrPrevRange(1);">
                                    <i class="calendar-icon ic-arrow-line-right" data-action="move-next"></i>
                                </button>
                            </li>
                            <li>
                                <select name="" id="">
                                    <option value="">결제페이지 노출순</option>
                                    <option value="">최근등록순</option>
                                    <option value="">종료임박순</option>
                                    <option value="">오픈순</option>
                                </select>
                            </li>
                        </div>

                        <div id="calendar">
                            <!-- 구매혜택 선택시 타구매혜택 유닛은 opacity 낮게 노출 (타구매혜택 유닛(tui-full-calendar-weekday-schedule on)) -->
                            <!-- 해당 기간에 등록된 구매혜택 개수 제한 없이 노출
                            - 해당기간 구매혜택 추가시 높이 조절 -->
                        </div>
                    </div>
                </div>
                <script>

                    var calendar = new tui.Calendar(document.getElementById('calendar'), {
                        defaultView: 'month',
                        useDetailPopup: true,
                        isReadOnly: true,
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
                                        <button class="btn_detail">자세히 보기</button>
                                        <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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
                                        <button class="btn_detail">자세히 보기</button>
                                         <button class="btn_close"><img src="https://webadmin.10x10.co.kr/images/icon/close_black.png" alt=""></button>
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