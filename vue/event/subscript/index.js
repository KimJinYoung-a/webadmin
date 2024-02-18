var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div style="height: 300px;">
            <table class="table table-dark table-search" style="width:30%;margin-left: 130px;">
                <colgroup>
                    <col style="width:150px">
                    <col style="width:300px">
                </colgroup>
                <tbody>
                    <h2>검색</h2>
                    <tr>
                        <th>이벤트 코드</th>
                        <td>
                            <input type="text" name="evt_code" v-model="search_param.evt_code"/>
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
                        <th>옵션1</th>
                        <td>
                            <input type="text" name="sub_opt1" v-model="search_param.sub_opt1"/>
                        </td>
                    </tr>
                    <tr>
                        <th>옵션2</th>
                        <td>
                            <input type="text" name="sub_opt2" v-model="search_param.sub_opt2"/>
                        </td>
                    </tr>
                    <tr>
                        <th>옵션3</th>
                        <td>
                            <input type="text" name="sub_opt3" v-model="search_param.sub_opt3"/>
                        </td>
                    </tr>
                    <tr>
                        <button @click="go_search" type="button" class="button dark">검색</button>
                    </tr>
                </tbody>
            </table>
            <p class="p-table">
                <span>검색결과 : <strong>{{total_count}}</strong></span>
                <button @click="go_excel_down" type="button" class="button secondary">엑셀 다운</button>
            </p>
            
            <table class="table table-dark">
                <colgroup>
                </colgroup>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>회원등급</th>
                        <th>회원번호</th>
                        <th>옵션1</th>
                        <th>옵션2</th>
                        <th>옵션3</th>
                        <th>가입일</th>
                        <th>응모신청일</th>
                        <th>당첨횟수</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in subscript_list" class="link">
                        <td>{{item.userid}}</td>
                        <td>{{item.userlevel}}</td>
                        <td>{{item.encoded_useq}}</td>
                        <td>{{item.sub_opt1}}</td>
                        <td>{{item.sub_opt2}}</td>
                        <td>{{item.sub_opt3}}</td>
                        <td>{{item.regdate}}</td>
                        <td>{{item.subscriptDate}}</td>
                        <td>{{item.win_count}}</td>
                    </tr>
                </tbody>
            </table>
            
            <Pagination @click_page="click_page" :current_page="search_param.page" :last_page="last_page"></Pagination>
        </div>
    `
    , data() {
        return{
            search_param : {
                evt_code : ""
                , start_date : ""
                , end_date : ""
                , sub_opt1 : ""
                , sub_opt2 : ""
                , sub_opt3 : ""

                , page : 1
                , page_size : 100
            }
            , subscript_list : []
            , total_count : 0
            , last_page : 1
        }
    }
    , created() {

    }
    , mounted(){
        const _this = this;
    }
    , computed: {
    }
    , methods: {
        go_search(){
            const _this = this;

            const api_data = {
                evt_code : this.search_param.evt_code
                , page : this.search_param.page
                , page_size : this.search_param.page_size
            };
            if(this.search_param.sub_opt1.trim() != ""){
                api_data.sub_opt1 = this.search_param.sub_opt1;
            }
            if(this.search_param.sub_opt2.trim() != ""){
                api_data.sub_opt2 = this.search_param.sub_opt2;
            }
            if(this.search_param.sub_opt3.trim() != ""){
                api_data.sub_opt3 = this.search_param.sub_opt3;
            }

            callApiHttps("GET", "/event/subscript-list", api_data, function (data){
                _this.subscript_list = data.subscript_list;
                _this.total_count = data.total_count;
                _this.last_page = data.last_page;
            });
        }
        , click_page(page){
            console.log(page);
            this.search_param.page = page;
            this.go_search();
            window.scrollTo(0, 0);
        }
        , go_excel_down(){
            const _this = this;

            if(this.search_param.evt_code == ""){
                return false;
            }

            let url_param = "evt_code=" + this.search_param.evt_code;
            if(this.search_param.sub_opt1.trim() != ""){
                url_param += "&sub_opt1=" + this.search_param.sub_opt1;
            }
            if(this.search_param.sub_opt2.trim() != ""){
                url_param += "&sub_opt2=" + this.search_param.sub_opt2;
            }
            if(this.search_param.sub_opt3.trim() != ""){
                url_param += "&sub_opt3=" + this.search_param.sub_opt3;
            }

            let api_url;
            if (location.hostname.startsWith("webadmin")) {
                api_url = "//fapi.10x10.co.kr/api/admin/v1";
            } else {
                //api_url = "//testfapi.10x10.co.kr:8080/api/admin/v1";
                api_url = "//localhost:8080/api/admin/v1";
            }

            window.open(api_url + '/event/subscript-list-excel?' + url_param,'excel','width=420,height=200');
        }
    }
});
