var app = new Vue({
    el: "#app",
    store: store,
    template: `
        <div>
            <div style="margin: 5px 30px 30px 30px;">
                <input type="button" @click="change_show_type('list')" value="리스트" class="btn"/>
                <input type="button" @click="change_show_type('contents')" value="컨텐츠" class="btn"/>
                <div style="float: right">
                    <p style="display: contents; font-size: 15px;"><b>{{nickname.occupation}}</b> {{nickname.nickname}}</p>
                    <input type="button" @click="popup_nickname" value="닉네임 변경" class="btn" />
                </div>
            </div>
            
            <div v-if="show_type == 'list'">
                <!-- 검색 테이블 -->
                <table class="table table-dark table-search">
                    <colgroup>
                        <col style="width:10%"/>
                        <col style="width:30%"/>
                        <col style="width:10%"/>
                        <col style="width:30%"/>
                        <col style="width:20%"/>
                    </colgroup>
                    <thead class="thead-tenbyten">
                        <tr>
                            <th>기간</th>
                            <td style="text-align:left;display: flex;" colspan="4">
                                <select id="period" class="form-control inline small" style="margin-right: 5px;">
                                    <option value="1">시작일 기준</option>
                                    <option value="2">종료일 기준</option>
                                    <option value="3">등록일 기준</option>
                                </select>
                                
                                <input id="startDate" class="form-control small" size="10" maxlength="10" style="margin-right: 5px;" />
                                 <p>~</p>
                                <input id="endDate" class="form-control small" size="10" maxlength="10" style="margin-left: 5px;" />
                            </td>
                            
                            <th>레이아웃</th>
                            <td style="text-align:left;">
                                <select id="uiNumber" class="form-control inline small">
                                    <option value="0">전체</option>
                                    <option value="1">리스트형</option>
                                    <option value="2">상세형</option>
                                    <option value="3">동영상형</option>
                                    <option value="4">이벤트형</option>
                                </select>
                            </td>
                            
                            <td rowspan="3">
                                <button @click="do_search" type="button" class="button dark">검색</button> <br/><br/>
                                <button @click="reload" type="button" class="button secondary">검색조건Reset</button>
                            </td>
                        </tr>
                        <tr>                            
                            <th>컨텐츠</th>
                            <td style="text-align:left;">
                                <select id="contentsNumber" class="form-control inline small">
                                    <option value="0">전체</option>
                                    <option value="1">마스터피스</option>
                                    <option value="2">탐구생활</option>
                                    <option value="3">DAY.FILM</option>
                                    <option value="4">THING.배지</option>
                                    <option value="5">PLAY.GOODS</option>
                                    <option value="7">WEEKLY WALLPAPER</option>
                                </select>
                            </td>
                            
                            <th>진행상태</th>
                            <td style="text-align:left;">
                                <select id="stateFlag" class="form-control inline small">
                                    <option value="0" selected="selected">선택</option>
                                    <option value="1">등록대기</option>
                                    <option value="2">디자인요청</option>
                                    <option value="3">퍼블리싱요청</option>
                                    <option value="4">개발요청</option>
                                    <option value="5">오픈요청</option>
                                    <option value="7">오픈</option>
                                    <option value="8">보류</option>
                                    <option value="9">종료</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>키워드 검색</th>
                            <td style="text-align:left;">
                                <select id="searchKey" class="form-control inline small">
                                    <option value="1">번호</option>
                                    <option value="2">컨텐츠명</option>
                                    <option value="3">작성자</option>
                                </select>
                            </td>
                            
                            <th></th>
                            <td style="text-align:left;">
                            </td>
                          </tr>                      
                    </thead>                
                </table>
                
                <p class="p-table">
                    <span>검색결과 : <strong>{{content_count}}</strong></span>
                    <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                        <option v-for="n in 5" :value="n*10">{{n*10}}개씩 보기</option>
                    </select>
                    <button id="reg_new_content" @click="popup_content('')" type="button" class="button dark">신규 등록</button>
                </p>
                
                <p style="margin-bottom: 50px;">
                    <table class="table table-dark">
                        <colgroup>
                            <col style="width:33%"/>
                            <col style="width:33%"/>
                            <col style="width:33%"/>
                        </colgroup>
                        <thead>
                          <tr>
                              <th>오프닝 1</th>
                              <th>오프닝 2</th>
                              <th>오프닝 3</th>
                          </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td v-for="item in opening_list">
                                    <div v-if="item.pidx">
                                        {{item.pidx}} {{item.contentTitleName}} <br/>
                                      <b>{{item.titlename}}</b>
                                      <p>{{item.startdate}} ~ {{item.enddate}} 오픈 </p>
                                      <input type="button" @click="deleteOpening(item.pidx)" value="제거"/>     
                                    </div>                                                           
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </p>
    
                <!-- 리스트 테이블 -->
                <table class="table table-dark">
                    <colgroup>
                        <col style="width:20%"/>
                        <col style="width:40%"/>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                        <col style="width:20%"/>
                    </colgroup>
                    <thead>
                        <tr>
                            <th>썸네일</th>
                            <th>상세</th>
                            <th>조회수</th>
                            <th>오픈 여부</th>
                            <th>기타</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr v-for="content in contents" :key="content.content_idx">
                            <td>
                                <p v-if="content.openingflag !== 0">{{content.openingflag}}</p>
                                <img @click="popup_thumbnail" :src="content.listimage" class="img-thumbnail link" style="width:50px;height:50px;" onerror="this.src='/images/loading.gif'" />
                            </td>
                            <td @click="popup_content(content.pidx)" class="link">
                                {{content.pidx}} {{content.contentTitleName}} <br/>
                                <b>{{content.titlename}}</b>
                                <p>{{content.occupation}} {{content.nickname}} {{content.regdate}} 등록</p>
                                <p v-if="content.lastupdate">{{content.lastOccupation}} {{content.lastNickname}} {{content.lastupdate}} 최종수정</p> 
                                <p>{{content.startdate}} ~ {{content.enddate}} 오픈</p>
                            </td>
                            <td>{{content.viewcount}}</td>
                            <td>{{content.stateflag_name}}</td>
                            <td><input type="button" @click="popup_content(content.pidx)" value="수정"/> <input type="button" @click="delete_playlist(content.pidx)" value="보류"/></td>
                        </tr>
                    </tbody>
                </table>
    
                <!-- 페이지 -->
                <Pagination
                    @click_page="click_page"
                    :current_page="current_page"
                    :last_page="last_page"
                ></Pagination>
    
                <!-- 등록/수정 모달 -->
                <Modal v-show="show_write_modal" @save="save_list_content" @close="show_write_modal = false" modal_width="830px" header_title="PLAY 컨텐츠 등록/수정">
                    <List-Write slot="body" :pop_content="pop_content" :pop_content_items="pop_content_items" :pop_content_tag="pop_content_tag" 
                        @change_content_tag="change_content_tag"
                        ref="write"/>
                </Modal>
    
                <!-- 썸네일 모달 -->
                <Modal v-show="show_thumbnail_modal" @close="show_thumbnail_modal = false"
                    modal_width="400px" :show_header_yn="false" :show_footer_yn="false"
                    :close_background_click_yn="true"
                >
                    <img width="100%" :src="popup_thumbnail_src" slot="body" />
                </Modal>
            </div>
            
            <div v-else-if="show_type == 'contents'">
                <table class="table table-dark table-search">
                    <colgroup>
                        <col style="width: 30%" />
                        <col style="width: 70%" />
                    </colgroup>
                    <thead class="thead-tenbyten">
                        <tr>
                            <th>운영여부</th>
                            <td style="text-align:left;display: flex;" colspan="4">
                                <select id="isUsing" class="form-control inline small">
                                    <option value="3">전체</option>
                                    <option value="1">운영중</option>
                                    <option value="0">운영안함</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>노출여부</th>
                            <td style="text-align:left;">
                                <select id="isView" class="form-control inline small">
                                    <option value="3">전체</option>
                                    <option value="1">공개</option>
                                    <option value="0">비공개</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>키워드 검색</th>
                            <td style="text-align:left;">
                                <select id="keywordSearchType" class="form-control inline small">
                                    <option value="number">번호</option>
                                    <option value="contentName">컨텐츠명</option>
                                    <option value="author">작성자</option>
                                </select>
                                
                                <input type="text" id="keywordSearch" />
                            </td>
                          </tr>         
                          <tr>
                            <td class="td-button align-right">
                                <button @click="reload" type="button" class="button secondary">검색조건Reset</button>
                                <button @click="do_contents_search" type="button" class="button dark">검색</button>
                            </td>
                          </tr>            
                    </thead>
                </table>
                <p class="p-table">
                    <span>검색결과 : <strong>{{content_count}}</strong></span>
                    <i class='fas fa-sync' @click="reload"></i>
                    <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                        <option v-for="n in 5" :value="n*10">{{n*10}}개씩 보기</option>
                    </select>
                    <button id="reg_new_content" @click="popup_content('')" type="button" class="button dark">신규 등록</button>
                </p>            
    
                <!-- 리스트 테이블 -->
                <table class="table table-dark">
                    <colgroup>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                        <col style="width:20%"/>
                        <col style="width:10%"/>
                        <col style="width:20%"/>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                    </colgroup>
                    <thead>
                        <tr>
                            <th>번호</th>
                            <th>운영여부</th>
                            <th>컨텐츠 명</th>
                            <th>최초 등록정보</th>
                            <th>최종 수정정보</th>
                            <th>노출 순서</th>
                            <th>노출 여부</th>
                            <th>수정 / 삭제</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr v-for="item in contents">
                            <td>{{item.cidx}}</td>                        
                            <td>{{item.isusing == 1 ? '운영함' : '운영안함'}}</td>
                            <td>{{item.titlename}}</td>
                            <td>
                                <p>{{item.occupation}} {{item.nickname}}</p>
                                <p>{{item.regdate}}</p>
                            </td>
                            <td>
                                <p>{{item.lastOccupation}} {{item.lastNickname}}</p>
                                <p>{{item.lastupdate}}</p>
                            </td>
                            <td>{{item.sortnum}}</td>
                            <td>{{item.isview == 1 ? '노출함' : '노출안함'}}</td>
                            <td><input type="button" @click="popup_content(item.cidx)" value="수정"/> <input type="button" @click="delete_content(item.cidx)" value="삭제"/></td>
                        </tr>
                    </tbody>
                </table>
    
                <!-- 페이지 -->
                <Pagination
                    @click_page="click_page"
                    :current_page="current_page"
                    :last_page="last_page"
                ></Pagination>
    
                <!-- 등록/수정 모달 -->
                <Modal v-show="show_write_modal" @save="save_contents" @close="show_write_modal = false" modal_width="830px" header_title="PLAY 컨텐츠 등록/수정">
                    <Content-Write slot="body" :pop_content="pop_content" ref="write"/>
                </Modal>
            </div>
            
            <Modal v-show="show_nickname_modal" @save="save_nickname" @close="show_nickname_modal = false" modal_width="400px" header_title="닉네임 수정">
                <Nickname-Write slot="body" :nickname="nickname" ref="nickname"/>
            </Modal>
        </div>
    `,
    data() {
        return {
            show_write_modal: false // 등록 모달 노출여부
            , show_thumbnail_modal: false // 썸네일 모달 노출여부
            , popup_thumbnail_src: "" // 썸네일 모달 이미지 src
            , show_nickname_modal : false //닉네임 모달 노출여부
            , check_ok:false // 유효검증 플래그값
            , is_saving : false //비동기통신 중복처리 방지 플래그값
        };
    },
    created() {
        this.$store.dispatch("GET_CONTENTS"); // 컨텐츠 리스트 조회
        this.$store.dispatch("GET_OPENING_LIST");
        this.$store.dispatch("GET_NICKNAME");
    },
    computed: {
        content_count() {
          // 컨텐츠 총 개수
          return this.$store.getters.content_count;
        },
        contents() {
          // 컨텐츠 리스트
          return this.$store.getters.contents;
        },
        current_page() {
          // 현재 페이지
          return this.$store.getters.current_page;
        },
        last_page() {
          // 마지막 페이지
          return this.$store.getters.last_page;
        },
        pop_content() {
          // 기존 컨텐츠
          return this.$store.getters.pop_content;
        }
        , opening_list(){
            return this.$store.getters.opening_list;
        }
        , pop_content_items(){
            return this.$store.getters.pop_content_items;
          }
        , pop_content_tag(){
            return this.$store.getters.pop_content_tag;
        }
        , show_type(){
            return this.$store.getters.show_type;
        }
        , nickname(){
            return this.$store.getters.nickname;
        }
    },
    methods: {
        click_page(page) {
            // 페이지 클릭 이벤트
            this.$store.commit("SET_CURRENT_PAGE", page);
            this.$store.dispatch("GET_CONTENTS");
            window.scrollTo(0, 0);
        },
        do_search() {
            // 검색버튼 클릭 이벤트
            this.$store.commit("SET_SEARCH_PARAMETER", {
              period: $("#period").val()
              , startdate: $("#startDate").val()
              , enddate: $("#endDate").val()
              , uinumber: $("#uiNumber").val()
              , contentsnumber: $("#contentsNumber").val()
              , stateflag: $("#stateFlag").val()
              , searchkey: $("#searchKey").val()
              , searchstring: $("#searchString").val()
            });
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_CONTENTS");
        }
        , do_contents_search() {
            this.$store.commit("SET_CONTENTS_SEARCH_PARAMETER", {
                isusing: $("#isUsing").val()
                , isview: $("#isView").val()
                , keywordsearchtype: $("#keywordSearchType").val()
                , keywordsearch: $("#keywordSearch").val()
            });
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_CONTENTS");
        }
        ,change_page_size() {
            // 페이지당 컨텐츠 노출 수 변경
            this.$store.commit("SET_PAGE_SIZE", document.getElementById("page_size").value);
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_CONTENTS");
        },
        popup_content(pidx) {
            if(pidx != ""){
              this.$store.dispatch("GET_POP_CONTENT", pidx);
            }else{
              this.$store.commit("SET_POP_CONTENT", {});
              this.$store.commit("SET_POP_CONTENT_ITEMS", []);
              this.$store.commit("SET_POP_CONTENT_TAG", []);
            }
            this.show_write_modal = true;
        },
        popup_thumbnail(e) {
            // popup 썸네일 모달
            this.popup_thumbnail_src = e.target.src;
            this.show_thumbnail_modal = true;
        }
        , popup_nickname(){
            this.show_nickname_modal = true;
        }
        , save_list_content() {
            if(this.is_saving){
                return false;
            }

            this.validate_content_data();
            if(this.check_ok){
                const _this = this;
                _this.is_saving = true;

                _this.save_image().then(function (data){
                    // 컨텐츠 저장
                    const form_data = new FormData(document.play_content);
                    const api_data = {};
                    form_data.forEach((value, key) => {
                        api_data[key] = value;
                    });

                    // let apiType = "Put"; //수정
                    let apiType = "Post"; //수정
                    let url = "/mobileSite/play/update/list-content";
                    if(!$("input[name=pidx]").val()){
                        apiType="Post"; //등록
                        url = "/mobileSite/play/list-content";
                    }
                    callApiHttps(apiType, url, api_data, function (data) {
                        _this.is_saving = false;

                        alert("저장 되었습니다.");
                        _this.$store.dispatch("GET_CONTENTS");
                        _this.$store.dispatch("GET_OPENING_LIST");
                        _this.show_write_modal = false;
                    }, function (xhr){
                        _this.is_saving = false;
                        console.log("ajax error", xhr)
                    });
                });
            }
        }
        , save_contents(){
            if(this.is_saving){
                return false;
            }

            this.validate_content_data();
            if(this.check_ok){
                const _this = this;

                _this.save_image().then(function (data){
                    // 컨텐츠 저장
                    const form_data = new FormData(document.play_content);
                    const api_data = {};
                    form_data.forEach((value, key) => {
                        api_data[key] = value;
                    });

                    let apiType = "Put"; //수정
                    if(!$("input[name=cidx]").val()){
                        apiType="Post"; //등록
                    }
                    callApiHttps(apiType, "/mobileSite/play/contents-content", api_data, function (data) {
                        _this.is_saving = false;

                        alert("저장 되었습니다.");
                        _this.$store.dispatch("GET_CONTENTS");
                        _this.show_write_modal = false;
                    }, function (xhr){
                        _this.is_saving = false;
                        console.log("ajax error", xhr)
                    });
                });
            }
        }
        , save_image(){
            const _this = this;
            return new Promise(function (resolve, reject) {
                console.log("this.show_type", _this.show_type);
                let imgChangeF = "";
                let file, filePath;
                if(_this.show_type == "list"){
                    imgChangeF = $("input[name=listimageChangeF]").val();
                    file = document.getElementById("addListimage").files[0];
                    filePath = "list_listimage";
                }else if(_this.show_type == "contents"){
                    imgChangeF = $("input[name=mainimageChangeF]").val();
                    file = document.getElementById("addMainimage").files[0];
                    filePath = "contents_mainimage";
                }

                if(imgChangeF == "Y"){
                    //리스트 이미지 저장
                    const imgData = new FormData();
                    imgData.append('sfImg', file);
                    imgData.append('sName', filePath);

                    let api_url;
                    if (location.hostname.startsWith('webadmin')) {
                        api_url = '//upload.10x10.co.kr';
                    } else {
                        api_url = '//testupload.10x10.co.kr';
                    }
                    $.ajax({
                        url: api_url + "/linkweb/play/play_admin_imgreg_json.asp"
                        , type: "POST"
                        , processData: false
                        , contentType: false
                        , data: imgData
                        , crossDomain: true
                        , success: function (data) {
                            const response = JSON.parse(data);

                            if (response.response === 'ok') {
                                if(_this.show_type == "list"){
                                    app.$refs.write.current_content.listimage = response.imgurl;
                                    console.log(app.$refs.write.current_content.listimage);
                                }else if(_this.show_type == "contents"){
                                    app.$refs.write.current_content.mainimage = response.imgurl;
                                    console.log(app.$refs.write.current_content.mainimage);
                                }

                                return resolve();
                            } else {
                                alert('이미지 저장 중 오류가 발생했습니다. (Err: 001)');
                                return reject();
                            }
                        }
                    });
                }else{
                    return resolve();
                }
            });
        }
        , validate_content_data(content) {
            const _this = this;

            this.check_ok = true;
            $(".must").each(function(){
                if($(this).val().trim() == ""){
                    _this.check_ok = false;
                    let th_name = $(this).parent().parent().find("th")[0].innerText;
                    alert("필수항목 " + th_name + "를 입력하지 않으셨습니다.");
                    $(this).focus();

                    return false;
                }
            });

            if(this.check_ok && $("select[name=stateflag]").val() == "0"){
                _this.check_ok = false;
                alert("진행상황을 선택해주세요.");
                $("select[name=stateflag]").focus();

                return false;
            }
        },
        reload() {
            window.location.reload(true);
        }
        , deleteOpening(pidx){
            const _this = this;

            if(confirm("제거하시겠습니까?")){
                const api_data = {
                    pidx : pidx
                    , lastadminid : sessionStorage.getItem("ssBctId")
                };

                callApiHttps("delete", "/mobileSite/play/openinglist", api_data, function (data) {
                    alert("제거 됐습니다.");
                    _this.$store.dispatch("GET_OPENING_LIST");
                    _this.show_write_modal = false;
                });
            }
        }
        , delete_playlist(pidx) {
            const _this = this;

            if (confirm("보류하시겠습니까?")) {
                const api_data = {
                    pidx : pidx
                    , lastadminid : sessionStorage.getItem("ssBctId")
                };

                callApiHttps("DELETE", "/mobileSite/play/list", api_data, function (data) {
                    alert("보류 됐습니다.");
                    _this.$store.dispatch("GET_CONTENTS");
                });
            }
        }
        , change_show_type(show_type){
            this.$store.commit("SET_SHOW_TYPE", show_type);
            this.$store.dispatch("GET_CONTENTS");
        }
        , delete_content(cidx){
            const _this = this;

            if (confirm("삭제하시겠습니까?")) {
                const api_data = {
                    cidx : cidx
                };

                callApiHttps("DELETE", "/mobileSite/play/contents", api_data, function (data) {
                    alert("삭제 됐습니다.");
                    _this.$store.dispatch("GET_CONTENTS");
                });
            }
        }
        , save_nickname(){
            const _this = this;
            let api_data = $("#play_nickname").serialize();

            callApiHttps("PUT", "/mobileSite/play/nickname", api_data, function (data) {
                alert("저장 되었습니다.");
                _this.$store.dispatch("GET_NICKNAME");
                _this.show_nickname_modal = false;
            }, function (xhr){console.log("ajax error", xhr)});
        }
        , change_content_tag(data){
            this.$store.commit("SET_POP_CONTENT_TAG", data);
        }
    }
    , mounted() {
        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $(".thead-tenbyten #startDate").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
        });
        $(".thead-tenbyten #endDate").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
        });
    }
});
