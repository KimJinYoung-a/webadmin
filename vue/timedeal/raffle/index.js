var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div style="height: 300px;">            
            <p class="p-table">
                <span>총 수 : <strong>{{search_count}}</strong></span>
                <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                    <option v-for="n in 5" :value="n*10">{{n*10}}개씩 보기</option>
                </select>
                <button id="reg_new_content" @click="popup_detail()" type="button" class="button dark">신규 등록</button>
                <button @click="reset_cache" type="button" class="button dark">캐시 지우기</button>
            </p>
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:40%;">
                    <col style="width:60%;">
                </colgroup>
                <thead>
                    <tr>
                        <th>이벤트코드</th>
                        <th>등록일</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in lists" @click="popup_detail(item.evt_code)" class="link">
                        <td>{{item.evt_code}}</td>
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
                <Timesale-Detail slot="body" ref="detail" @reload="reload"                     
                    @save="go_save" @close="show_write_modal = false"
                    :content="content" :content_schedule="content_schedule" 
                    :timesale_is_write="timesale_is_write" 
                />
            </Scroll-Modal>
        </div>
    `
    , data() {
        return{
            show_write_modal: false
            , is_saving : false
        }
    }
    , created() {
        this.$store.dispatch("GET_LISTS");
    }
    , mounted(){
        const _this = this;
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
        reload(evt_code) {
            this.$store.dispatch("GET_LISTS");
            if(evt_code){
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

            if(_this.content.evt_code && _this.content.evt_code.trim() != ""){
                callApiHttps("PUT", "/event/timedeal", form_data, function(data){
                    alert("수정되었습니다.");
                });
            }else{
                callApiHttps("post", "/event/timedeal", form_data, function(data){
                    alert("저장되었습니다.");
                });
            }
        }
        , click_page(page) {
            this.$store.commit("SET_CURRENT_PAGE", page);
            this.$store.dispatch("GET_LISTS");
            window.scrollTo(0, 0);
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
