let app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div>
            <form id="schedule_form">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:150px">
                        <col style="width:300px">
                        <col style="width:200px">
                        <col style="width:300px">
                    </colgroup>
                    <tbody>
                        <h2>스케쥴 정보</h2>
                            
                        <tr>
                            <th>이벤트 코드</th>
                            <td>
                                <input :value="now_schedule.evt_code" type="text" name="evt_code" readonly class="must" min="1"/>                                
                            </td>
                            
                            <input v-model="now_schedule.schedule_idx" type="number" name="schedule_idx" class="must" min="1" style="display: none;"/>
                        </tr>
                        <tr>
                            <th>기간</th>
                            <td colspan="3">
                                <input type="text" id="start_day" v-model="start_day" style="float: left; width: 90px;" class="must" autocomplete="false"/>
                                <input type="text" id="start_time" v-model="start_time" required size="8" style="float: left;" class="must" autocomplete="false"/>
                                <input type="text" name="startDate" v-model="now_schedule.startDate" class="must" style="display: none"/>
                                <p style="float: left;">&nbsp ~ &nbsp</p>
                                <input type="text" id="end_day" v-model="end_day" style="float: left; width: 90px;" class="must" autocomplete="false"/>
                                <input type="text" id="end_time" v-model="end_time" required size="8" style="float: left;" class="must" autocomplete="false"/>
                                <input type="text" name="endDate" v-model="now_schedule.endDate" class="must" style="display: none"/>                           
                            </td>
                        </tr>                                          
                    </tbody>
                </table>  
            </form>  
            
            <div style="margin: 30px 0px  30px 747px;">                
                <button @click="go_save" class="button dark">{{is_write ? '수정' : '저장'}}</button>
                <button @click="go_close" class="button secondary">취소</button>
            </div> 
            
            <br/>
            <h2>미끼상품</h2>          
            <p style="color: red;"> * 미끼상품은 저장후에 등록해주세요.</p>      
            <input type="button" value="추가" @click="add_mikki()" />
            <template v-for="item in mikki_list">
                <table class="table table-write table-dark" style="margin-top: 15px;">
                    <colgroup>
                        <col style="width:150px">
                        <col style="width:300px">
                        <col style="width:200px">
                        <col style="width:300px">
                    </colgroup>
                    <tbody>                            
                        <tr>
                            <th>기간</th>
                            <td>
                                {{item.startDate.substr(0, 16)}} ~ {{item.endDate.substr(0, 16)}}                                
                            </td>
                            <td colspan="2">
                                <input type="button" value="수정" @click="edit_mikki(item.startDate, item.endDate)"/>
                                <input type="button" value="삭제" @click="delete_mikki(item.startDate, item.endDate)"/>
                            </td>
                        </tr>
                        <tr>
                            <th>상품 코드</th>
                            <td>
                                {{item.itemid}}
                            </td>
                            
                            <th>상품 이름</th>
                            <td>
                                {{item.itemName}}
                            </td>
                        </tr>     
                        <tr>
                            <th>상품 이미지</th>
                            <td>
                                <img :src="item.itemImage" style="height: 80px;"/>
                            </td>
                            
                            <th>상품 개수</th>
                            <td>
                                {{item.itemCnt}}
                            </td>
                        </tr>
                        <tr>
                            <th>상품 판매가</th>
                            <td>
                                {{format_price(item.orgPrice)}}
                            </td>
                            
                            <th>상품 할인가</th>
                            <td>
                                {{format_price(item.sellCash)}}
                            </td>
                        </tr>
                        <tr>
                            <th>할인율</th>
                            <td>
                                {{item.saleValue}}
                            </td>
                            
                            <th>할인구분</th>
                            <td>
                                {{item.saleType == 1 ? '비율' : '금액'}}
                            </td>
                        </tr>                               
                    </tbody>
                </table>  
            </template>
                    
            <br/>
            <h2>일반상품</h2>
            <button v-show="normal_list.length > 0" @click="update_sort" class="button secondary">순서 적용</button>
            <button v-show="normal_list.length > 0" @click="go_normal_one_save" class="button secondary">상품 추가</button>
            <table class="table table-write table-dark" style="margin-top: 15px;">
                <colgroup>
                    <col style="width:40px">
                    <col style="width:30px">
                    <col style="width:50px">
                    <col style="width:50px">
                    <col style="width:50px">
                    <col style="width:200px">
                    <!--<col style="width:60px">
                    <col style="width:200px">
                    <col style="width:60px">
                    <col style="width:200px">-->
                    <col style="width:50px">
                    <col style="width:20px">
                    <col style="width:50px">
                </colgroup>
                <tbody id="sorting_row">                            
                    <tr v-if="normal_list.length == 0">
                        <th>상품코드</th>
                        <td colspan="3" >
                            <p style="color: red;"> * 첫등록은 아이템코드.아이템코드,...  형식으로 저장합니다.</p>
                            <textarea v-model="itemid_list" rows="5" style="width:100%;" name="itemid_list" class="must"/>
                            <button @click="go_normal_save" class="button secondary">저장</button>
                        </td>
                    </tr>   
                    <tr v-else v-for="item in normal_list" :data-idx="item.itemid">
                        <th>정렬순서</th>
                        <td style="text-align: center;">
                            {{item.sortNo}}
                        </td>
                        
                        <th>상품코드</th>
                        <td style="text-align: center;">
                            {{item.itemid}}
                        </td>
                        
                        <th>상품명</th>
                        <td>
                            {{item.itemname}}
                        </td>
                        
                        <!--<th>커스텀 상품명</th>
                        <td>
                            {{item.custom_name}}
                        </td>
                        
                        <th>커스텀 이미지</th>
                        <td>
                            {{item.custom_image}}
                        </td>-->
                        
                        <th>상품 분류</th>
                        <td style="text-align: center;">
                            {{getItemdivNameLink(item.itemdiv)}}
                        </td>
                        
                        <td>
                            <button @click="go_normal_edit(item.itemid)" class="button secondary">수정</button>
                            <button @click="go_normal_delete(item.itemid)" class="button secondary">삭제</button>
                        </td>
                    </tr>       
                </tbody>
            </table>
            
            <Scroll-Modal v-show="show_mikki_write_modal" header_title="미끼상품 등록/수정" 
                :show_footer_yn="false"
            >
                <MIKKI slot="body" :content_mikki="content_mikki" :mikki_is_write="mikki_is_write"
                    :evt_code="now_schedule.evt_code" :schedule_idx="parseInt(now_schedule.schedule_idx)" :mikki_list="mikki_list"
                    :schedule_start="schedule.startDate" :schedule_end="schedule.endDate"
                    @reload="reload" @close="close_mikki_modal"
                />
            </Scroll-Modal>
            
            <Modal v-show="show_normal_write_modal" header_title="일반상품 등록" 
                :show_footer_yn="false"
            >
                <NORMAL slot="body" :evt_code="now_schedule.evt_code" :schedule_idx="parseInt(now_schedule.schedule_idx)"
                    :normal_list="normal_list"
                    @reload="reload" @close="close_normal_modal"
                />
            </Modal>
        </div>
    `
    , data() {
        return{
            now_schedule : {
                evt_code : ""
                , schedule_idx : ""
                , startDate : ""
                , endDate : ""
            }
            , now_product : []
            , now_mikki_detail: {}
            , show_mikki_write_modal: false
            , show_normal_write_modal : false
            , mikki_count : 1
            , itemid_list : ""

            , start_day : ""
            , start_time : ""
            , end_day : ""
            , end_time : ""

            , is_write : false
            , sorted_arr : []
        }
    }
    , created() {
        let query_param = new URLSearchParams(window.location.search);
        this.now_schedule.evt_code = query_param.get("evt_code");
        this.now_schedule.schedule_idx = query_param.get("schedule_idx");

        this.$store.commit("SET_EVT_CODE", this.now_schedule.evt_code);
        this.$store.commit("SET_SCHEDULE_IDX", this.now_schedule.schedule_idx);
        this.$store.dispatch("GET_SCHEDULE", this.now_schedule.schedule_idx);
        this.$store.dispatch("GET_MIKKI_LIST", this.now_schedule.schedule_idx);
        this.$store.dispatch("GET_NORMAL_LIST", this.now_schedule.schedule_idx);
    }
    , mounted(){
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#start_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#end_day").datepicker('setDate', min_date);
                $("#end_day").datepicker('option', "minDate", min_date);

                _this.start_day = document.getElementById("start_day").value;
            }
        });
        $("#end_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.end_day = document.getElementById("end_day").value;
            }
        });
        $("#start_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.start_time = document.getElementById("start_time").value;
            }
        });
        $("#end_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.end_time = document.getElementById("end_time").value;
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
    }
    , computed: {
        schedule(){
            return this.$store.getters.schedule;
        }
        , content_mikki() {
            return this.$store.getters.content_mikki;
        }
        , mikki_is_write(){
            return this.$store.getters.mikki_is_write;
        }
        , mikki_list(){
            return this.$store.getters.mikki_list;
        }
        , normal_list(){
            return this.$store.getters.normal_list;
        }
    }
    , methods: {
        go_save() {
            const _this = this;
            return new Promise(function (resolve, reject) {
                _this.now_schedule.startDate = _this.start_day + " " + _this.start_time;
                _this.now_schedule.endDate = _this.end_day + " " + _this.end_time;

                return resolve();
            }).then(function(){
                window.opener.app.$refs.detail.child_save_flag = true;
                let form_data = $("#schedule_form").serialize();

                if(_this.is_write){
                    callApiHttps("PUT", "/event/timedeal-schedule", form_data, function(data){
                        alert("수정되었습니다.");
                    });
                }else{
                    callApiHttps("POST", "/event/timedeal-schedule", form_data, function(data){
                        alert("저장되었습니다.");
                    });
                }
            });
        }
        , go_close(){
            window.close();
        }
        , add_mikki(){
            this.$store.dispatch("GET_MIKKI_DETAIL");
            this.show_mikki_write_modal = true;
        }
        , edit_mikki(startDate, endDate){
            let param = [startDate, endDate];
            this.$store.dispatch("GET_MIKKI_DETAIL", param);
            this.show_mikki_write_modal = true;
        }
        , delete_mikki(startDate, endDate){
            if(confirm("삭제하시겠습니까?")){
                let param = [startDate, endDate];
                this.$store.dispatch("DELETE_MIKKI_DETAIL", param);
                this.reload();
            }
        }
        , reload() {
            this.$store.dispatch("GET_SCHEDULE", this.now_schedule.schedule_idx);
            this.$store.dispatch("GET_MIKKI_LIST", this.now_schedule.schedule_idx);
            this.$store.dispatch("GET_NORMAL_LIST", this.now_schedule.schedule_idx);
        }
        , close_mikki_modal(){
            this.show_mikki_write_modal = false;
        }
        , close_normal_modal(){
            this.show_normal_write_modal = false;
        }
        , go_normal_save(){
            const _this = this;
            callApiHttps("POST", "/event/timedeal-normal-list", {"evt_code": this.schedule.evt_code, "schedule_idx": this.now_schedule.schedule_idx, "itemid_list": this.itemid_list}, function(data){
                alert("저장되었습니다.");
                _this.$store.dispatch("GET_NORMAL_LIST", _this.now_schedule.schedule_idx);
            });
        }
        , update_sort(){
            const _this = this;
            let sord_data = {"sort_idx" : this.sorted_arr, "evt_code" : this.schedule.evt_code, "schedule_idx" : this.now_schedule.schedule_idx};

            callApiHttps("PUT", "/event/timedeal-normal-list-sort", sord_data, function (data) {
                alert("순서 변경 완료");
                _this.$store.commit("SET_NORMAL_LIST_EMPTY");
                _this.$store.dispatch("GET_NORMAL_LIST", _this.now_schedule.schedule_idx);
            });
        }
        , go_normal_delete(itemid){
            const _this = this;
            if(confirm("삭제하시겠습니까?")){
                callApiHttps("DELETE", "/event/timedeal-normal-list", {"evt_code": this.schedule.evt_code, "schedule_idx": this.now_schedule.schedule_idx, "itemid": itemid}, function(data){
                    alert("삭제되었습니다.");
                    _this.$store.dispatch("GET_NORMAL_LIST", _this.now_schedule.schedule_idx);
                });
            }
        }
        , go_normal_one_save(){
            this.show_normal_write_modal = true;
        }
        , getItemdivNameLink(itemdiv){
            return getItemdivName(itemdiv);
        }
        , format_price(price){
            if(price){
                return price.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            }
        }
    }
    , watch:{
        schedule(schedule) {
            if(schedule){
                this.now_schedule = schedule;

                this.start_day = this.schedule.startDate.split(" ")[0];
                this.start_time = this.schedule.startDate.split(" ")[1].substring(0, 5);

                this.end_day = this.schedule.endDate.split(" ")[0];
                this.end_time = this.schedule.endDate.split(" ")[1].substring(0, 5);

                this.is_write = true;
            }else{
                let query_param = new URLSearchParams(window.location.search);
                this.now_schedule = {
                    evt_code : query_param.get("evt_code")
                    , schedule_idx : query_param.get("schedule_idx")
                    , startDate : ""
                    , endDate : ""
                };

                this.is_write = false;
            }
        }
    }
});
