var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div style="height: 300px;">
            <table class="table table-dark table-search" style="width:80%;margin-left: 130px;">
                <colgroup>
                    <col style="width:150px">
                    <col style="width:300px">
                    <col style="width:200px">
                    <col style="width:300px">
                </colgroup>
                <tbody>
                    <p>기본설정</p>
                    <tr>
                        <th>이벤트 코드</th>
                        <td>
                            <input type="text" name="evt_code" v-model="search_param.evt_code"/>
                        </td>
                        <th>스케쥴 시퀀스</th>
                        <td>
                            <input type="text" name="schedule_idx" />
                        </td>
                    </tr>
                    <tr>
                        <th>이벤트 기간</th>
                        <td>
                            <input v-model="search_param.start_date" type="text" id="start_date" style="float: left; width: 90px;"/>
                            <p style="float: left;">&nbsp ~ &nbsp</p>
                            <input v-model="search_param.end_date"  type="text" id="end_date" style="float: left; width: 90px;"/>
                        </td>
                    </tr>
                    <tr>
                        <button @click="reload" type="button" class="button secondary">검색조건Reset</button>
                        <button @click="go_search" type="button" class="button dark">검색</button>
                    </tr>
                </tbody>
            </table>
            <p class="p-table">
                <span>검색결과 : <strong>{{search_count}}</strong></span>
                <i class='fas fa-sync' @click="reload"></i>
                <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                    <option v-for="n in 5" :value="n*10">{{n*10}}개씩 보기</option>
                </select>
                <button id="reg_new_content" @click="popup_detail('')" type="button" class="button dark">신규 등록</button>
                <button @click="reset_cache" type="button" class="button dark">캐시 지우기</button>
            </p>
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:20%;">
                    <col style="width:20%;">
                    <col style="width:30%;">
                </colgroup>
                <thead>
                    <tr>
                        <th>이벤트코드</th>
                        <th>티저 이벤트코드</th>
                        <th>등록일</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in lists" @click="popup_detail(item.evt_code)" class="link">
                        <td>{{item.evt_code}}</td>
                        <td>{{item.tz_evt_code}}</td>
                        <td>{{item.regDate}}</td>
                    </tr>
                </tbody>
            </table>
            
            <Pagination
                @click_page="click_page"
                :current_page="current_page"
                :last_page="last_page"
            ></Pagination>
            
            <Scroll-Modal v-show="show_write_modal"                
                :show_footer_yn="false" header_title="타임세일 등록/수정"
            >
                <Timesale-Detail slot="body" :content="content" :content_schedule="content_schedule" :timesale_is_write="timesale_is_write"
                    @change_image_flag="change_image_flag" @reload="reload" 
                    ref="detail"
                    @save="go_save" @close="show_write_modal = false"
                />
            </Scroll-Modal>
        </div>
    `
    , data() {
        return{
            search_param : {
                evt_code : ""
                , start_date : ""
                , end_date : ""
            }
            , show_write_modal: false
            , image_flag : false
        }
    }
    , created() {
        this.$store.dispatch("GET_LISTS");
    }
    , mounted(){
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#start_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#end_date").datepicker('setDate', min_date);
                $("#end_date").datepicker('option', "minDate", min_date);

                _this.search_param.start_date = document.getElementById("start_date").value;
            }
        });
        $("#end_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.search_param.end_date = document.getElementById("end_date").value;
            }
        });
    }
    , computed: {
        search_count(){
            return this.$store.getters.search_count;
        }
        , lists(){
            return this.$store.getters.lists;
        }
        , content(){
            return this.$store.getters.content;
        }
        , content_schedule() {
            return this.$store.getters.content_schedule;
        }
        , timesale_is_write(){
            return this.$store.getters.timesale_is_write;
        }

        , current_page() {
            return this.$store.getters.current_page;
        }
        , last_page() {
            return this.$store.getters.last_page;
        }
    }
    , methods: {
        go_search(){

        }
        , reload(evt_code) {
            this.$store.dispatch("GET_LISTS");
            if(evt_code){
                console.log("check");
                this.$store.dispatch("GET_CONTENT", evt_code);
            }
        }
        , popup_detail(evt_code){
            this.$store.dispatch("GET_CONTENT", evt_code);
            this.show_write_modal = true;
        }
        , change_page_size() {
            this.$store.commit("SET_PAGE_SIZE", document.getElementById("page_size").value);
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_LISTS");
        }
        , go_save() {
            const _this = this;
            let form_data = $("#timesale_detail").serialize();
            let file1 = document.getElementById("katalkImage").files[0];

            this.save_image(file1).then(function(data){
                form_data += "&katalkImage=" + data;
                //console.log("form_data", form_data);

                if(_this.content.evt_code && _this.content.evt_code.trim() != ""){
                    callApiHttps("PUT", "/event/timedeal", form_data, function(data){
                        alert("수정되었습니다.");
                    });
                }else{
                    callApiHttps("post", "/event/timedeal", form_data, function(data){
                        alert("저장되었습니다.");
                    });
                }
            });
        }
        , click_page(page) {
            this.$store.commit("SET_CURRENT_PAGE", page);
            this.$store.dispatch("GET_LISTS");
            window.scrollTo(0, 0);
        }
        , change_image_flag(){
            this.image_flag = true;
        }
        , save_image(file1){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                if(_this.image_flag){
                    imgData.append('imgFile1', file1);
                    imgData.append("imgFolder", "timedeal");
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/event_admin/timedeal_admin_imgreg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);

                        let imgurl = response.imgurl1 ? response.imgurl1 : _this.content.katalkImage;

                        return resolve(imgurl);
                    }
                    , error : function (request,status,error){
                        console.log("code", request.status);
                        console.log("message", request.responseText);
                        console.log("error", error);

                        return reject();
                    }
                });
            });
        }
        , reset_cache(){
            if(confirm("캐시를 비우시겠습니까?")){
                callApiHttps("GET", "/event/timedeal-delete-cache", null, function(data){
                    alert("캐시를 비웠습니다.");
                });
            }
        }
    }
});
