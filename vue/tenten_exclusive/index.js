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
                    <p>단독상품 관리</p>
                    <tr>
                        <th>상품 코드</th>
                        <td>
                            <input v-model="search_param.itemid" type="text" name="itemid" style="float: left;"/>
                        </td>
                        <th>프론트 노출여부</th>
                        <td>
                            <select v-model="search_param.display_yn" name="display_yn">
                                <option value="">전체</option>
                                <option value="Y">Y</option>
                                <option value="N">N</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <th>상품오픈일</th>
                        <td>
                            <input v-model="search_param.open_date" type="text" id="search_open_date" style="float: left; width: 90px;"/>
                        </td>
                        <th>상태</th>
                        <td>
                            <select v-model="search_param.state">
                                <option value="">전체</option>
                                <option value="1">판매예정</option>
                                <option value="2">판매중</option>
                                <option value="3">판매완료</option>
                                <option value="4">노출대기</option>
                                <option value="5">비노출</option>
                            </select>
                        </td>
                    </tr>                    
                </tbody>                
            </table>
            <div style="float: right; margin-right: 11px;">
                <button @click="reload" type="button" class="button secondary">검색조건Reset</button>
                <button @click="go_search" type="button" class="button dark">검색</button>
            </div>
            <p class="p-table">
                <span>검색결과 : <strong>{{total_count}}</strong></span>
                <button id="reg_new_content" @click="popup_write('')" type="button" class="button dark">신규 등록</button>
                <button @click="update_sort" type="button" class="button dark">순서 변경</button>                
                <button @click="go_delete_item" type="button" class="button dark">선택 삭제</button>
            </p>
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:5%;">
                    <col style="width:5%;">
                    <col style="width:10%;">
                    <col style="width:20%;">
                    <col style="width:20%;">
                    <col style="width:10%;">
                    <col style="width:10%;">
                    <col style="width:20%;">
                </colgroup>
                <thead>
                    <tr>
                        <th><input @click="go_check_all" type="checkbox" id="check_all"></th>
                        <th>idx</th>
                        <th>상품코드</th>
                        <th>상품명</th>
                        <th>상품오픈일</th>
                        <th>판매중/판매예정 노출여부</th>
                        <th>상태(노출영역)</th>
                        <th>컨텐츠 등록</th>
                    </tr>
                </thead>
                <tbody id="sorting_row">
                    <tr v-for="item in items" class="link" :data-idx="item.exclusive_idx">
                        <td><input :value="item.exclusive_idx" type="checkbox" name="item_checkbox"></td>
                        <td @click="popup_write(item.exclusive_idx)" >{{item.exclusive_idx}}</td>
                        <td @click="popup_write(item.exclusive_idx)" >{{item.itemid}}</td>
                        <td @click="popup_write(item.exclusive_idx)" >{{item.itemname}}</td>
                        <td>{{item.open_date}}</td>
                        <td>{{item.display_yn}} / {{item.pre_display_yn}}</td>
                        <td>{{item.state}}</td>
                        <td>
                            <input type="button" value="메인 리스트" @click="popup_main($event, item.exclusive_idx)"/>
                            <input type="button" value="상품 상세페이지" @click="popup_detail($event, item.exclusive_idx)"/>
                        </td>
                    </tr>
                </tbody>
            </table>
            
            <Scroll-Modal v-show="show_write_modal" :show_footer_yn="false" header_title="텐텐단독 등록/수정">
                <Tenten-Exclusive-Write slot="body" ref="write" 
                    :content="write" :is_written="is_written"
                    @change_image_flag="change_image_flag" @reload="reload"                     
                    @close="show_write_modal = false"
                />
            </Scroll-Modal>
            
            <Scroll-Modal v-show="show_main_modal" :show_footer_yn="false" header_title="텐텐단독 상품메인">
                <Tenten-Exclusive-Main slot="body"
                    :content="main" :is_written="is_written"
                    @change_image_flag="change_image_flag" @reload="reload"                     
                    @save="go_main_save" @close="show_main_modal = false"
                />
            </Scroll-Modal>
            
            <Scroll-Modal v-show="show_detail_modal" :show_footer_yn="false" header_title="텐텐단독 상품 상세">
                <Tenten-Exclusive-Detail slot="body"
                    :content="detail" :is_written="is_written"
                    @reload="reload"                     
                    @close="show_detail_modal = false"
                />
            </Scroll-Modal>
        </div>
    `
    , data() {
        return{
            search_param : {
                itemid : ""
                , search_open_date : ""
                , display_yn : ""
                , state : ""
            }
            , is_saving : false

            , show_write_modal: false
            , show_main_modal : false
            , show_detail_modal : false

            , item_image_flag : false

            , sorted_arr : []
        }
    }
    , created() {
        this.$store.dispatch("GET_ITEMS");
    }
    , mounted(){
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#search_open_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.search_param.open_date = document.getElementById("search_open_date").value;
            }
        });

        $("#sorting_row").sortable({
            delay: 150
            , stop: function() {
                let sortedArrVar = new Array();

                $('#sorting_row > tr').each(function() {
                    sortedArrVar.push($(this).attr("data-idx"));
                });
                _this.sorted_arr = sortedArrVar;
            }
        });

        $("input[name=item_checkbox]").click(function(){
            $("#check_all").attr("checked", false);
        });
    }
    , computed: {
        total_count(){
            return this.$store.getters.total_count;
        }
        , items(){
            return this.$store.getters.items;
        }
        , write(){
            return this.$store.getters.write;
        }
        , main(){
            return this.$store.getters.main;
        }
        , detail(){
            return this.$store.getters.detail;
        }
        , is_written(){
            return this.$store.getters.is_written;
        }
    }
    , methods: {
        go_search(){
            this.$store.dispatch("GET_ITEMS", this.search_param);
        }
        , reload() {
            this.search_param = {
                itemid : ""
                , search_open_date : ""
                , display_yn : ""
                , state : ""
            };
            this.$store.dispatch("GET_ITEMS");
        }
        , popup_write(exclusive_idx){
            this.$store.dispatch("GET_WRITE", exclusive_idx);
            this.show_write_modal = true;
        }
        , popup_main(e, popup_main){
                e.cancelBubble = true;

            this.$store.dispatch("GET_MAIN", popup_main);
            this.show_main_modal = true;
        }
        , popup_detail(e, itemid) {
            e.cancelBubble = true;

            this.$store.dispatch("GET_DETAIL", itemid);
            this.show_detail_modal = true;
        }
        , go_main_save(){
            const _this = this;

            if(this.is_saving){
                return false;
            }

            let form_data = $("#tenten_exclusive_main").serialize();
            let file1 = document.getElementById("item_img").files[0];

            this.save_image(file1).then(function(data){
                if(data){
                    form_data += "&item_img=" + data.photo1;
                }
                console.log("form_data", form_data);

                callApiHttps("post", "/tenten-exclusive/item-main", form_data, function(data){
                    alert("저장되었습니다.");
                    _this.show_main_modal = false;
                    _this.is_saving = false;
                });
            });
        }
        , change_image_flag(type){
            if(type == "item_img"){
                this.item_image_flag = true;
            }
        }
        , save_image(file1){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                if(_this.item_image_flag){
                    imgData.append('photo1', file1);
                    imgData.append("folderName", "item_img");
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/tenten_exclusive/tenten_exclusive_reg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);

                        if (response.response === 'ok') {
                            return resolve(response);
                        } else if(response.response === 'none'){
                            return resolve();
                        }else {
                            alert('이미지 저장 중 오류가 발생했습니다. (Err: 001)');
                            return reject();
                        }
                    }
                    , error : function (e){
                        console.log(e);

                        return reject();
                    }
                });
            });
        }
        , go_delete_item(){
            const _this = this;
            let delete_exclusive_idx = [];
            if(confirm("삭제 하시겠습니까?")){
                $("input[name=item_checkbox]").each(function(){
                   if(this.checked == true){
                       delete_exclusive_idx.push(parseInt(this.value));
                   }
                });

                let api_data = {"delete_exclusive_idx" : delete_exclusive_idx.toString()};
                callApiHttps("delete", "/tenten-exclusive/item", api_data, function(data){
                    alert("삭제되었습니다.");
                    _this.$store.dispatch("GET_ITEMS");
                });
            }
        }
        , update_sort(){
            const _this = this;
            let sort_data = {"sort_idx" : this.sorted_arr};

            callApiHttps("PUT", "/tenten-exclusive/item", sort_data, function (data) {
                alert("순서 변경 완료");
            });
        }
        , go_check_all(){
            const check_all_value = $("#check_all").prop("checked");
            console.log("check_all_value", check_all_value);
            if(check_all_value){
                $("input[name=item_checkbox]").each(function(){
                    this.checked = true;
                });
            }else{
                $("input[name=item_checkbox]").each(function(){
                    this.checked = false;
                });
            }

        }
    }
});
