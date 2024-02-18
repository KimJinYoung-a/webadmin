var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div>
            <p class="p-table">
                <button @click="go_reg('')" type="button" class="button dark">�ű� ���</button>
            </p>
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:10%;">
                    <col style="width:50%;">
                    <col style="width:10%;">
                    <col style="width:20%;">
                    <col style="width:10%;">
                </colgroup>
                <thead>
                    <tr>
                        <th>��ȣ</th>
                        <th>key</th>
                        <th>��뿩��</th>
                        <th>����</th>
                        <th></th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in key_list" @click="go_reg(item.idx)">
                        <td>{{item.idx}}</td>
                        <td>{{item.validationKey}}</td>
                        <td>{{item.isusing}}</td>
                        <td>{{item.description}}</td>
                        <td><button @click.stop="go_delete(item.idx)" type="button" class="button dark">����</button></td>
                    </tr>
                </tbody>
            </table>
            
            <Modal v-show="show_write_modal"
                @save="save_content" @close="show_write_modal = false"
                modal_width="830px" header_title="�� Ű ���/����"
            >
                <APP-KEY-WRITE slot="body" :content="content" />
            </Modal>
        </div>
    `
    , data() {
        return{
            show_write_modal : false
        }
    }
    , created() {
        this.$store.dispatch("GET_KEY_LIST");
    }
    , mounted(){
    }
    , computed: {
        key_list(){
            return this.$store.getters.key_list;
        }
        , content(){
            return this.$store.getters.content;
        }
    }
    , methods: {
        save_content(){
            const _this = this;
            let form_data = $("#content_form").serialize();
            if(this.content.idx){
                // ���� callApiHttps�� /v1/ ��θ� �������̶� v2�� ����Ҽ� ���� ������ ����.
                callApiHttpsV2("PUT", "/v2/app/appkey", form_data, function (data){
                    alert("�����Ǿ����ϴ�.");
                    _this.$store.dispatch("GET_KEY_LIST");
                    _this.show_write_modal = false;
                }, function(xhr) {
                    let errorJson = JSON.parse(xhr.responseText);
                    if (errorJson.message) {
                        alert(errorJson.message);
                    } else {
                        alert("������ ������ �߻��߽��ϴ�.");
                    }
                });
            }else{
                // ���� callApiHttps�� /v1/ ��θ� �������̶� v2�� ����Ҽ� ���� ������ ����.
                callApiHttpsV2("POST", "/v2/app/appkey", form_data, function (data){
                    alert("����Ǿ����ϴ�.");
                    _this.$store.dispatch("GET_KEY_LIST");
                    _this.show_write_modal = false;
                }, function(xhr) {
                    let errorJson = JSON.parse(xhr.responseText);
                    if (errorJson.message) {
                        alert(errorJson.message);
                    } else {
                        alert("������ ������ �߻��߽��ϴ�.");
                    }
                });
            }
        }
        , go_reg(idx){
            if(idx){
                this.$store.dispatch("GET_KEY", idx);
                this.show_write_modal = true;
            }else{
                this.$store.dispatch("GET_KEY");
                this.show_write_modal = true;
            }
        }
        , go_delete(idx){
            const _this = this;
            // ���� callApiHttps�� /v1/ ��θ� �������̶� v2�� ����Ҽ� ���� ������ ����.
            callApiHttpsV2("DELETE", "/v2/app/appkey", {"idx" : idx}, function (data){
                alert("�����Ǿ����ϴ�.");
                _this.$store.dispatch("GET_KEY_LIST");
            }, function(xhr) {
                let errorJson = JSON.parse(xhr.responseText);
                if (errorJson.message) {
                    alert(errorJson.message);
                } else {
                    alert("������ ������ �߻��߽��ϴ�.");
                }
            });
        }
    }
    , watch:{
    }
});
