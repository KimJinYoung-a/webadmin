<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : LINKER 관리
'	History		: 2021.10.14 이전도 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
</p>
<link rel="stylesheet" type="text/css" href="/css/linker.css">
<link rel="stylesheet" href="https://cdn.materialdesignicons.com/3.6.95/css/materialdesignicons.min.css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">

<div class="container">
    <header class="linker-title">
        <h3>LINKER 관리</h3>
        <div class="nickname-btn-area">
            <a href="#">닉네임 사전</a>
            <a href="#">닉네임 비속어 관리</a>
        </div>
    </header>

    <div class="linker-content">
        <div class="forum-list-container">
            <div class="title">
                <h5>포럼</h5>
                <div class="btn-area">
                    <button class="linker-btn">포럼 등록</button>
                    <button class="linker-btn">노출 순서관리</button>
                </div>
            </div>

            <!-- region 포럼 리스트 -->
            <ul class="forum-list">
                <li class="on">
                    <p class="number">03</p>
                    <div class="info">
                        <strong>텐바이텐 20주년을 축하해주세요!</strong>
                        <span>2021-09-09 ~ 2021-09-31 / 오픈</span>
                    </div>
                </li>
                <li>
                    <p class="number">02</p>
                    <div class="info">
                        <strong>텐킨버스데이! 도넛 받아가세요!</strong>
                        <span>2021-09-09 ~ 2021-09-31 / 오픈안함</span>
                    </div>
                </li>
            </ul>
            <!-- endregion -->

        </div>
        <div class="forum-content">

            <!-- region 상단 포럼 정보 -->
            <div class="forum-content-top">
                <div class="title-control">
                    <div class="title">
                        <p>아니 벌써?</p>
                        <h5>텐바이텐 20주년을 축하해주세요!</h5>
                    </div>
                    <div class="btn-area">
                        <button class="linker-btn">포럼 수정</button>
                        <button class="linker-btn">포럼설명 등록</button>
                        <button class="linker-btn">포럼 삭제</button>
                    </div>
                </div>
                <div class="title-info">
                    <span>운영기간 : 2021-09-09 ~ 2021-09-31</span>
                    <span>프론트 오픈여부 : 오픈</span>
                    <span>노출 순서 : 2</span>
                </div>
            </div>
            <!-- endregion -->

            <div class="forum-info">
                <div class="title">
                    <div>
                        <h3>포럼 안내</h3>
                        <span>포럼 안내는 5개까지만 등록할 수 있습니다.</span>
                    </div>
                    <div>
                        <button class="linker-btn">포럼 안내 등록</button>
                        <button class="linker-btn">정렬수정</button>
                        <button class="linker-btn">선택 항목 삭제</button>
                    </div>
                </div>

                <table class="forum-list-tbl">
                    <colgroup>
                        <col style="width: 50px;">
                        <col style="width: 100px;">
                        <col style="width: 300px;">
                        <col>
                    </colgroup>
                    <thead>
                        <tr>
                            <th><input type="checkbox"></th>
                            <th>노출순서</th>
                            <th>안내제목</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><input type="checkbox"></td>
                            <td>1</td>
                            <td class="tl" colspan="2">스무 살 이야기 #2 텐바이텐과 함께 자라온 '히치하이커'</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="forum-content-bottom">

                <div class="title">
                    <h3>포스팅 리스트</h3>
                    <a>신고 포스팅 관리</a>
                </div>

                <!-- region 검색 -->
                <div class="search">
                    <div>
                        <div class="search-group">
                            <label>회원구분:</label>
                            <select>
                                <option>전체</option>
                                <option>Host</option>
                                <option>Guest</option>
                                <option>User</option>
                            </select>
                        </div>
                        <div class="search-group">
                            <label>회원등급:</label>
                            <select>
                                <option>전체</option>
                                <option>STAFF</option>
                                <option>VVIP</option>
                                <option>VIP GOLD</option>
                            </select>
                        </div>
                        <div class="search-group">
                            <select>
                                <option>닉네임</option>
                            </select>
                            :
                            <input type="text">
                        </div>
                        <div class="search-group">
                            <label>등록일자:</label>
                            <input type="text" class="date" readonly> ~
                            <input type="text" class="date" readonly>
                        </div>
                    </div>
                    <button class="linker-btn">검색</button>
                </div>
                <!-- endregion -->

                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p>검색결과 : <span>1,312</span></p>
                        <div>
                            <button class="linker-btn">선택 항목 고정</button>
                            <button class="linker-btn">고정 포스팅 관리</button>
                        </div>
                    </div>

                    <!-- region 포스팅 리스트 -->
                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 50px;">
                            <col style="width: 70px;">
                            <col style="width: 180px;">
                            <col>
                            <col style="width: 110px;">
                            <col style="width: 150px;">
                            <col style="width: 170px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>idx</th>
                                <th>작성자 정보</th>
                                <th>작성내용</th>
                                <th>상단 고정여부</th>
                                <th>작성일지</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>Host / STAFF / 헤이즈짱</td>
                                <td>미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</td>
                                <td class="posting-red">Y</td>
                                <td>2021-09-08 15:13:15</td>
                                <td>
                                    <button class="linker-btn">수정</button>
                                    <button class="linker-btn">삭제</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1757</td>
                                <td>User / VIP GOLD / 헤이즈짱</td>
                                <td>미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</td>
                                <td>N</td>
                                <td>2021-09-08 15:13:15</td>
                                <td>
                                    <button class="linker-btn">수정</button>
                                    <button class="linker-btn">삭제</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <!-- endregion -->

                    <!-- region 페이징 -->
                    <ul class="pagination">
                        <li class="disabled"><a>&lt;</a></li>
                        <li class="on"><a>1</a></li>
                        <li><a>2</a></li>
                        <li><a>3</a></li>
                        <li><a>&gt;</a></li>
                    </ul>
                    <!-- endregion -->

                </div>
            </div>
        </div>
    </div>


    <!-- region 포럼 신규등록 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>포럼 신규등록</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:100px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>제목</th>
                                <td><input type="text" placeholder="포럼 제목을 입력해주세요"></td>
                            </tr>
                            <tr>
                                <th>부제목</th>
                                <td><input type="text" placeholder="포럼 부제목을 입력해주세요"></td>
                            </tr>
                            <tr>
                                <th>설명</th>
                                <td><textarea placeholder="포럼 설명을 입력해주세요"></textarea></td>
                            </tr>
                            <tr>
                                <th>백그라운드<br>PC</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="backPcImage" type="radio" checked>
                                        <label for="backPcImage">이미지</label>
                                        <input id="backPcVideo" type="radio">
                                        <label for="backPcVideo">동영상</label>
                                    </p>
                                    <button class="linker-btn">이미지 첨부</button>
                                </td>
                            </tr>
                            <tr>
                                <th>백그라운드<br>M</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="backPcImage" type="radio" checked>
                                        <label for="backPcImage">이미지</label>
                                        <input id="backPcVideo" type="radio">
                                        <label for="backPcVideo">동영상</label>
                                    </p>
                                    <input type="text" placeholder="영상 URL을 입력해주세요">
                                </td>
                            </tr>
                            <tr>
                                <th>운영기간</th>
                                <td>
                                    <span class="datepicker">
                                        <label for="datepicker1">
                                            <strong>시작일</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker1" readonly>

                                        <label for="datepicker2">
                                            <strong>종료일</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker2" readonly>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <th>프론트<br>노출여부</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="showY" type="radio" checked>
                                        <label for="showY">Y</label>
                                        <input id="showN" type="radio">
                                        <label for="showN">N</label>
                                    </p>
                                </td>
                            </tr>
                            <tr>
                                <th>정렬순서</th>
                                <td><input type="text" style="width: 100px;"></td>
                            </tr>
                            <tr>
                                <th>비고</th>
                                <td><textarea></textarea></td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">저장</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 포럼 노출 순서관리 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 1100px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>포럼 노출 순서 관리</h3>
            </div>
            <div class="modal-container">
                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p><span>1,312</span>개</p>
                        <div>
                            <button class="linker-btn">정렬순서 수정</button>
                        </div>
                    </div>

                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 30px;">
                            <col style="width: 60px;">
                            <col>
                            <col style="width: 200px;">
                            <col style="width: 120px;">
                            <col style="width: 70px;">
                            <col style="width: 160px;">
                            <col style="width: 140px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>NO.</th>
                                <th>포럼 제목</th>
                                <th>포럼 부제목</th>
                                <th>프론트 오픈 여부</th>
                                <th>정렬순서</th>
                                <th>운영기간</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>던킨과 함께하는 텐킨버스데이!</td>
                                <td>텐텐 20주년 이벤트</td>
                                <td>오픈</td>
                                <td><input type="text" class="forum-sort"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">수정</button>
                                    <button class="linker-btn">삭제</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>던킨과 함께하는 텐킨버스데이!</td>
                                <td>텐텐 20주년 이벤트</td>
                                <td>오픈</td>
                                <td><input type="text" class="forum-sort"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">수정</button>
                                    <button class="linker-btn">삭제</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 포럼 안내 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 750px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>포럼 안내</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>포럼 안내 개수</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="descrCount1" type="radio" name="descrCount" checked>
                                        <label for="descrCount1">1개</label>
                                        <input id="descrCount2" type="radio" name="descrCount">
                                        <label for="descrCount2">2개</label>
                                        <input id="descrCount3" type="radio" name="descrCount">
                                        <label for="descrCount3">3개</label>
                                        <input id="descrCount4" type="radio" name="descrCount">
                                        <label for="descrCount4">4개</label>
                                        <input id="descrCount5" type="radio" name="descrCount">
                                        <label for="descrCount5">5개</label>
                                    </p>
                                    <p class="descr">포럼 설명이 2개 이상일 경우, 팝업으로 표기됩니다</p>
                                </td>
                            </tr>
                            <tr>
                                <th>포럼 안내1</th>
                                <td>
                                    <p class="forum-descr-title">
                                        <input type="text" placeholder="포럼 안내 1번 제목을 입력해주세요">
                                        <button class="linker-btn">미리보기</button>
                                    </p>
                                    <textarea class="forum-descr-code" rows="8" placeholder="설명 내용을 코드로 입력해주세요"></textarea>
                                </td>
                            </tr>
                            <tr>
                                <th>포럼 안내2</th>
                                <td>
                                    <p class="forum-descr-title">
                                        <input type="text" placeholder="포럼 안내 2번 제목을 입력해주세요">
                                        <button class="linker-btn">미리보기</button>
                                    </p>
                                    <textarea class="forum-descr-code" rows="8" placeholder="설명 내용을 코드로 입력해주세요"></textarea>
                                </td>
                            </tr>
                            <tr>
                                <th>샘플코드</th>
                                <td>
                                    <textarea class="forum-descr-sample" rows="6" wrap="off" readonly></textarea>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">저장</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 포럼 안내 미리보기 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap forum-descr-preview" style="width: 400px;">
            <div class="modal-header"></div>
            <div class="modal-body">
                <div class="modal-cont">
                    <div class="ex_img">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_01.gif" alt="since 2001">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_02.jpg?v=2" alt="텐바이텐, 머그컵에 꽤나 진심인걸?">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_03.gif" alt="머그컵">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_04.jpg?v=2" alt="20번째 머그컵 드디어 공개!">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_05.gif?v=2.1" alt="머그컵들을 구경해 볼까요?">
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 포스팅 관리 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>포스팅 관리</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>idx</th>
                                <td class="content">1234</td>
                            </tr>
                            <tr>
                                <th>회원구분</th>
                                <td class="content">User</td>
                            </tr>
                            <tr>
                                <th>회원등급</th>
                                <td class="content">VIP</td>
                            </tr>
                            <tr>
                                <th>닉네임</th>
                                <td class="content">요란한다람쥐#123</td>
                            </tr>
                            <tr>
                                <th>작성내용</th>
                                <td><textarea rows="5">미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</textarea></td>
                            </tr>
                            <tr>
                                <th>연결컨텐츠</th>
                                <td><button class="linker-btn">이벤트 : 348282</button></td>
                            </tr>
                            <tr>
                                <th>상단 고정 여부</th>
                                <td>
                                    <input id="fixPostingY" type="radio" name="fixPosting" checked>
                                    <label for="fixPostingY">고정</label>
                                    <input id="fixPostingN" type="radio" name="fixPosting">
                                    <label for="fixPostingN">고정안함</label>
                                </td>
                            </tr>
                            <tr>
                                <th>작성일시</th>
                                <td class="content">
                                    <strong>2021-09-08 15:13:15</strong>
                                    <span class="posting-update">2021-09-10 14:11:12 수정</span>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">저장</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 포스팅 고정 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>포스팅 고정</h3>
                <p class="add-descr">여러 항목을 선택했을 경우에는 일괄 적용됩니다.</p>
            </div>
            <div class="modal-container">
                <div>

                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th><span class="required">고정 순서<i></i></span></th>
                                <td><input type="text" style="width: 50px;"></td>
                            </tr>
                            <tr>
                                <th><span class="required">고정 기간<i></i></span></th>
                                <td>
                                    <span class="datepicker">
                                        <label for="datepicker3">
                                            <strong>시작일</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker3" readonly>

                                        <label for="datepicker4">
                                            <strong>종료일</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker4" readonly>
                                    </span>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">저장</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 고정 포스팅 관리 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 1100px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>고정 포스팅 관리</h3>
            </div>
            <div class="modal-container">
                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p><span>1,312</span>개</p>
                        <div>
                            <button class="linker-btn">고정 해제</button>
                            <button class="linker-btn">노출순서 수정</button>
                        </div>
                    </div>

                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 30px;">
                            <col style="width: 60px;">
                            <col style="width: 170px;">
                            <col>
                            <col style="width: 100px;">
                            <col style="width: 160px;">
                            <col style="width: 170px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>idx</th>
                                <th>작성자 정보</th>
                                <th>작성내용</th>
                                <th>고정 노출 순서</th>
                                <th>고정 기간</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>Host / STAFF / 랄랄라</td>
                                <td>미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</td>
                                <td><input type="text" class="forum-sort" value="1"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">수정</button>
                                    <button class="linker-btn long">고정 해제</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>User / VIP / 헤이즈짱</td>
                                <td>미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</td>
                                <td><input type="text" class="forum-sort" value="2"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">수정</button>
                                    <button class="linker-btn long">고정 해제</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 신고 포스팅 관리 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 1100px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>신고 포스팅 관리</h3>
            </div>
            <div class="modal-container">
                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p><span>1,312</span>개</p>
                        <div>
                            <button class="linker-btn">블락 해제</button>
                            <button class="linker-btn">선택 포스팅 삭제</button>
                        </div>
                    </div>

                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 30px;">
                            <col style="width: 60px;">
                            <col style="width: 80px;">
                            <col style="width: 170px;">
                            <col>
                            <col style="width: 170px;">
                            <col style="width: 200px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>idx</th>
                                <th>프로필 이미지</th>
                                <th>작성자 정보</th>
                                <th>작성내용</th>
                                <th>연결된 컨텐츠</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td><img src="//fiximage.10x10.co.kr/web2015/common/img_profile_04.png" class="thumb"></td>
                                <td>Host / STAFF / 랄랄라</td>
                                <td>미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</td>
                                <td><button class="linker-btn link">이벤트 : 348282</button></td>
                                <td>
                                    <button class="linker-btn long">블락 해제</button>
                                    <button class="linker-btn long">포스팅 삭제</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td><img src="//fiximage.10x10.co.kr/web2015/common/img_profile_04.png" class="thumb"></td>
                                <td>Host / STAFF / 랄랄라</td>
                                <td>미치겠어요 우연히 치명적인 귀여움을 찾아버렸는데 아까워서 나눠봐요. 하 귀여운건 진짜 최고야 XD</td>
                                <td><button class="linker-btn link">외부 URL</button></td>
                                <td>
                                    <button class="linker-btn long">블락 해제</button>
                                    <button class="linker-btn long">포스팅 삭제</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 닉네임 사전 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 900px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>닉네임 사전</h3>
            </div>
            <div class="modal-container">
                <div>
                    <div class="search">
                        <div>
                            <div class="search-group">
                                <select>
                                    <option>단어1</option>
                                    <option>단어2</option>
                                </select>
                                :
                                <input type="text">
                            </div>
                        </div>
                        <button class="linker-btn">검색</button>
                    </div>

                    <div class="modal-nicknames-area">
                        <div class="modal-nicknames-content">
                            <div class="nicknames-btn-area">
                                <button class="linker-btn">신규등록</button>
                                <button class="linker-btn">삭제</button>
                            </div>

                            <table class="forum-list-tbl">
                                <colgroup>
                                    <col style="width: 50px;">
                                    <col style="width: 60px;">
                                    <col>
                                    <col style="width: 150px;">
                                </colgroup>
                                <thead>
                                    <tr>
                                        <th><input type="checkbox"></th>
                                        <th>NO.</th>
                                        <th>단어1</th>
                                        <th></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>요란한</td>
                                        <td>
                                            <button class="linker-btn">수정</button>
                                            <button class="linker-btn">삭제</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>건들거리는</td>
                                        <td>
                                            <button class="linker-btn">수정</button>
                                            <button class="linker-btn">삭제</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>순수한</td>
                                        <td>
                                            <button class="linker-btn">수정</button>
                                            <button class="linker-btn">삭제</button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>

                        <div class="modal-nicknames-content">
                            <div class="nicknames-btn-area">
                                <button class="linker-btn">신규등록</button>
                                <button class="linker-btn">삭제</button>
                            </div>

                            <table class="forum-list-tbl">
                                <colgroup>
                                    <col style="width: 50px;">
                                    <col style="width: 60px;">
                                    <col>
                                    <col style="width: 150px;">
                                </colgroup>
                                <thead>
                                    <tr>
                                        <th><input type="checkbox"></th>
                                        <th>NO.</th>
                                        <th>단어2</th>
                                        <th></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>다람쥐</td>
                                        <td>
                                            <button class="linker-btn">수정</button>
                                            <button class="linker-btn">삭제</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>햄스터</td>
                                        <td>
                                            <button class="linker-btn">수정</button>
                                            <button class="linker-btn">삭제</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>옥수수</td>
                                        <td>
                                            <button class="linker-btn">수정</button>
                                            <button class="linker-btn">삭제</button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 단어 등록 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>단어1 등록</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>단어1</th>
                                <td>
                                    <textarea rows="3"></textarea>
                                    <p class="descr">여러 단어를 추가할 경우 ',(쉼표)'로 구분하여 입력해주세요.</p>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">저장</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region 닉네임 비속어 관리 모달 -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 900px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>닉네임 비속어 관리</h3>
            </div>
            <div class="modal-container">
                <div>
                    <div class="search">
                        <div>
                            <div class="search-group">
                                <label>비속어:</label>
                                <input type="text">
                            </div>
                        </div>
                        <button class="linker-btn">검색</button>
                    </div>

                    <div>
                        <div class="nicknames-btn-area">
                            <button class="linker-btn">신규등록</button>
                            <button class="linker-btn">삭제</button>
                        </div>

                        <table class="forum-list-tbl">
                            <colgroup>
                                <col style="width: 50px;">
                                <col style="width: 60px;">
                                <col>
                                <col style="width: 150px;">
                            </colgroup>
                            <thead>
                                <tr>
                                    <th><input type="checkbox"></th>
                                    <th>NO.</th>
                                    <th>비속어</th>
                                    <th></th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td><input type="checkbox"></td>
                                    <td>112</td>
                                    <td>^&%&</td>
                                    <td>
                                        <button class="linker-btn">수정</button>
                                        <button class="linker-btn">삭제</button>
                                    </td>
                                </tr>
                                <tr>
                                    <td><input type="checkbox"></td>
                                    <td>112</td>
                                    <td>^%&^&^</td>
                                    <td>
                                        <button class="linker-btn">수정</button>
                                        <button class="linker-btn">삭제</button>
                                    </td>
                                </tr>
                                <tr>
                                    <td><input type="checkbox"></td>
                                    <td>112</td>
                                    <td>@*#(*</td>
                                    <td>
                                        <button class="linker-btn">수정</button>
                                        <button class="linker-btn">삭제</button>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>

                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->


    <script>
        document.querySelector('body').classList.add('noscroll');
        $(function(){
            for( let i=1 ; i<=4 ; i++ ) {
                $('#datepicker' + i).datepicker( {
                    inline: true,
                    showOtherMonths: true,
                    showMonthAfterYear: true,
                    monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
                    dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
                    dateFormat: 'yy-mm-dd',
                });
            }

            document.querySelector('.forum-descr-sample').value = ''
                + '<div class="ex_img">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_01.gif" alt="since 2001">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_02.jpg" alt="텐바이텐, 머그컵에 꽤나 진심인걸?">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_03.gif" alt="머그컵">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_04.jpg" alt="20번째 머그컵 드디어 공개!">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_05.gif" alt="머그컵들을 구경해 볼까요?">\n'
                + '</div>';
        });
    </script>

</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->