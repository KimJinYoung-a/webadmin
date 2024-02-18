Vue.component('Timesale-Detail',{
    template: `
        <div>
            <form id="timesale_detail">                
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col style="width:30%">
                        <col style="width:120px;">
                        <col style="width:30%">
                    </colgroup>
                    <tbody>
                        <h2>�⺻ ����</h2>
                        
                        <tr>
                            <th>�̺�Ʈ �ڵ�</th>
                            <td>
                                <template v-if="content.evt_code">
                                    <input v-model="now_content.evt_code" type="text" name="evt_code" readonly/>
                                </template>
                                <template v-if="!content.evt_code">
                                    <input v-model="now_content.evt_code" type="text" name="evt_code" />
                                </template>
                                
                            </td>
                            
                            <th>Ƽ�� �̺�Ʈ �ڵ�</th>
                            <td>
                                <input v-model="now_content.tz_evt_code" type="text" name="tz_evt_code" />
                            </td>
                        </tr>
                        <tr>
                            <th>īī���� ���ø� ��ȣ</th>
                            <td>
                                <input v-model="now_content.katalkTemplateNo" type="text" name="katalkTemplateNo" />
                            </td>
                            
                            <th>īī���� ����</th>
                            <td>
                                <input v-model="now_content.katalkTitle" type="text" name="katalkTitle" />
                            </td>
                        </tr>
                        <tr>
                            <th>īī���� ����</th>
                            <td colspan="3">
                                <textarea v-model="now_content.katalkContent" rows="4" name="katalkContent" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>īī���� �̹���</th>
                            <td colspan="3">
                                <p style="color:red;">- jpg, gif, png ������ ���ϸ� ��� �����մϴ�.</p>
                                <p style="color:red;">- �ִ� ũ�� : 1024 KB</p>
                                <div v-show="now_content.katalkImage" class="thumbnail-area">
                                    <img id="showPrizeImage" :src="now_content.katalkImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag($event)" name="katalkImage" id="katalkImage"/>    
                            </td>
                        </tr>
                        <tr>
                            <th>īī���� ��ũ ��ư��</th>
                            <td colspan="3">
                                <input v-model="now_content.katalkLinkButtonName" type="text" name="katalkLinkButtonName" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>īī���� ��ũ URL</th>
                            <td colspan="3">
                                <input v-model="now_content.katalkLinkUrl" type="text" name="katalkLinkUrl" style="width: 100%" />
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <div style="margin: 30px 0px  30px 740px;">                
                <button @click="$emit('save')" class="button dark">{{timesale_is_write ? '����' : '����'}}</button>
                <button @click="$emit('close')" class="button secondary">���</button>
            </div>
            
            <table class="table table-write table-dark">
                <tbody>
                    <h2>������ ����</h2>
                    <p style="color: red;"> * ������������ �����Ŀ� ������ּ���.</p>
                    
                    <input type="button" value="�߰�" @click="add_schedule()" />
                    
                    <SCHEDULE v-for="(item, index) in content_schedule" :key="index" :content_schedule="item" 
                        @go_schedule="go_schedule" @go_schedule_delete="go_schedule_delete"
                    ></SCHEDULE>
                </tbody>
            </table>
        </div>
    `
    , data(){
        return{
            now_content : {
                evt_code : ""
                , tz_evt_code : ""
                , katalkTemplateNo : ""
                , katalkTitle : ""
                , katalkContent : ""
                , katalkLinkButtonName : ""
                , katalkLinkUrl : ""
                , katalkImage : ""
            }
            , child_save_flag : false
        }
    }
    , props : {
        content : {
            evt_code : {type:String, default:""}
            , tz_evt_code : {type:String, default:""}
            , katalkTemplateNo : {type:String, default:""}
            , katalkTitle : {type:String, default:""}
            , katalkContent : {type:String, default:""}
            , katalkLinkButtonName : {type:String, default:""}
            , katalkLinkUrl : {type:String, default:""}
            , katalkImage : {type:String, default:""}
            , next_schedule_idx : {type:Number, default:1}
        }
        , content_schedule : {type:Array, default:[]}
        , timesale_is_write : {type:Boolean, default : false}
    }
    , methods : {
        add_schedule(){
            const _this = this;
            if(this.content.evt_code){
                callApiHttps("GET", "/event/timedeal-count", {"evt_code": this.content.evt_code}, function(data){
                    if(data > 0){
                        window.open(
                            "/admin/eventmanage/timesale/writeDetail.asp?evt_code=" + _this.content.evt_code + "&schedule_idx=" + (parseInt(_this.content.next_schedule_idx))
                            , "schedule"
                            , "width=1000, height=600"
                        );
                    }else{
                        alert("Ÿ�ӵ� �̺�Ʈ�� ���� ������ּ���. ����� �����ư�� ������ ��ϵ˴ϴ�.");
                    }
                });
            }else{
                alert("Ÿ�ӵ� �̺�Ʈ�� ���� ������ּ���. ����� �����ư�� ������ ��ϵ˴ϴ�.");
            }
        }
        , change_image_flag(event){
            const _this = this;
            const file = event.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            this.$emit("change_image_flag");
            reader.onload = function(e){
                _this.now_content.katalkImage = e.target.result;
            }
        }
        , go_schedule(schedule_idx){
            window.open(
                "/admin/eventmanage/timesale/writeDetail.asp?evt_code=" + this.content.evt_code + "&schedule_idx=" + schedule_idx
                , "schedule"
                , "width=800, height=900"
            );
        }
        , go_schedule_delete(schedule_idx){
            const _this = this;
            if(confirm("�����Ͻðڽ��ϱ�?")){
                callApiHttps("DELETE", "/event/timedeal-schedule", {"evt_code": this.content.evt_code, "schedule_idx":schedule_idx}, function(data){
                    _this.child_save_flag = true;
                    alert("�����Ǿ����ϴ�.");
                });
            }
        }
    }
    , watch : {
        timesale_is_write(timesale_is_write){
            if(timesale_is_write) {
                this.now_content = this.content;
            } else {
                this.now_content = {
                    evt_code : ""
                    , tz_evt_code : ""
                    , katalkTemplateNo : ""
                    , katalkTitle : ""
                    , katalkContent : ""
                    , katalkLinkButtonName : ""
                    , katalkLinkUrl : ""
                    , katalkImage : ""
                };
            }
        }
        , child_save_flag(){
            if(this.child_save_flag){
                this.$emit("reload", this.content.evt_code);
                this.child_save_flag = false;
            }
        }
    }
});