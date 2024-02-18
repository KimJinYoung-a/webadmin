var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div>
            <p class="p-table">
                <span>검색결과 : <strong></strong></span>
                
                <button @click="popup_content('', 'write')" type="button" class="button dark">신규 등록</button>
            </p>
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:5%"/>
                    <col style="width:5%"/>
                    <col style="width:20%"/>
                    <col style="width:20%"/>
                </colgroup>
                <thead>
                    <tr>
                        <th><input type="checkbox" /></th>
                        <th>idx</th>
                        <th>이벤트 코드</th>
                        <th>참여대상자 기간 시작일</th>
                        <th>참여대상자 기간 종료일</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in list" @click="popup_content(item.evt_code, 'edit')">
                        <td></td>
                        <td></td>
                        <td>{{item.evt_code}}</td>
                        <td>{{item.startdate_of_bill}}</td>
                        <td>{{item.enddate_of_bill}}</td>
                    </tr>
                </tbody>
            </table>
            
            <Modal v-show="show_write_modal" @save="go_write_modal_save" @close="show_write_modal = false" modal_width="830px" header_title="등록/수정">
                <Content-Write slot="body" ref="write"                    
                    :write_mode="write_mode" :content="content"
                />
            </Modal>
        </div>
    `
    , data() {
        return{
            show_write_modal : false
            , write_mode : ""
            , is_saving : false
            , validate_check_ok : false
        }
    }
    , created() {
        this.$store.dispatch("GET_LIST");
    }
    , mounted(){

    }
    , computed: {
        list(){
            return this.$store.getters.list;
        }
        , content(){
            return this.$store.getters.content;
        }
    }
    , methods: {
        popup_content(evt_code, mode) {
            if(mode == "edit"){
                this.$store.dispatch("GET_CONTENT", evt_code);
            }

            this.write_mode = mode;
            this.show_write_modal = true;
        }
        , go_write_modal_save(){
            if(this.is_saving){
                return false;
            }

            this.validate_content_data();
            if(this.validate_check_ok) {
                const _this = this;
                this.is_saving = true;

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = '//fapi.10x10.co.kr/api/admin/v1';
                }else if(location.hostname.startsWith('localwebadmin')) {
                    api_url = '//localhost:8080/api/admin/v1';
                }else{
                    api_url = '//testfapi.10x10.co.kr/admin/web/v1';
                }

                let form_data = $("#content").serialize();
                $.ajax({
                    type: "POST"
                    , url: api_url + "/automatic-event/fortunebill"
                    , data : form_data
                    , crossDomain: true
                    , xhrFields: {
                        withCredentials: true
                    }
                    , success: function(data){
                        alert("저장 되었습니다.");
                        _this.show_write_modal = false;
                    }
                    , error: function(xhr){
                        console.log("ajax error", xhr)
                    }
                    , complete : function(){
                        _this.is_saving = false;
                        _this.write_mode = "wait";
                    }
                });
            }
        }
        , validate_content_data() {
            const _this = this;

            this.validate_check_ok = true;
            $(".must").each(function(){
                if($(this).val().trim() == ""){
                    _this.validate_check_ok = false;
                    let th_name = $(this).parent().parent().find("th")[0].innerText;
                    alert("필수항목 " + th_name + "를 입력하지 않으셨습니다.");
                    $(this).focus();

                    return false;
                }
            });
        }
    }
    , watch:{

    }
});
