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
                                <template v-if="timesale_is_write">
                                    <input v-model="now_content.evt_code" type="text" name="evt_code" readonly style="background-color:lightgrey" />
                                </template>
                                <template v-else>
                                    <input v-model="now_content.evt_code" type="text" name="evt_code" />
                                </template>                                
                            </td>
                            
                            <th>���� ����</th>
                            <td><input type="text" name="raffleFlag" value="Y"  readonly style="background-color:lightgrey" /></td>
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
                
                <table class="table table-write table-dark" style="margin-top: 20px;">
                    <colgroup>
                        <col style="width:120px;">
                        <col style="width:30%">
                        <col style="width:120px;">
                        <col style="width:30%">
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>��÷�� �˸��� ���ø� ��ȣ</th>
                            <td>
                                <input v-model="now_content.winnerTemplateNo" type="text" name="winnerTemplateNo" />
                            </td>
                            
                            <th>��÷�� �˸��� ����</th>
                            <td>
                                <input v-model="now_content.winnerTitle" type="text" name="winnerTitle" />
                            </td>
                        </tr>
                        <tr>
                            <th>��÷�� �˸��� ����</th>
                            <td colspan="3">
                                <textarea v-model="now_content.winnerContent" rows="4" name="winnerContent" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>��÷�� �˸��� ��ũ ��ư��</th>
                            <td colspan="3">
                                <input v-model="now_content.winnerLinkButtonName" type="text" name="winnerLinkButtonName" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>��÷�� �˸��� ��ũ URL</th>
                            <td colspan="3">
                                <input v-model="now_content.winnerLinkUrl" type="text" name="winnerLinkUrl" style="width: 100%" />
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
                        :raffle_flag="'Y'"
                        @go_schedule="go_schedule" @go_schedule_delete="go_schedule_delete" @go_winner_popup="go_winner_popup"
                    ></SCHEDULE>
                </tbody>
            </table>
        </div>
    `
    , data(){
        return{
            now_content : {
                evt_code : null
                , katalkTemplateNo : ""
                , katalkTitle : ""
                , katalkContent : ""
                , katalkLinkButtonName : ""
                , katalkLinkUrl : ""
                , winnerTemplateNo : ""
                , winnerTitle : ""
                , winnerContent : ""
                , winnerLinkButtonName : ""
                , winnerLinkUrl : ""
            }
            , child_save_flag : false
        }
    }
    , props : {
        content : {
            evt_code : {type:String, default:null}
            , katalkTemplateNo : {type:String, default:""}
            , katalkTitle : {type:String, default:""}
            , katalkContent : {type:String, default:""}
            , katalkLinkButtonName : {type:String, default:""}
            , katalkLinkUrl : {type:String, default:""}
            , winnerTemplateNo : {type:String, default:""}
            , winnerTitle : {type:String, default:""}
            , winnerContent : {type:String, default:""}
            , winnerLinkButtonName : {type:String, default:""}
            , winnerLinkUrl : {type:String, default:""}
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
        , go_winner_popup(schedule_idx){
            window.open(
                "/admin/eventmanage/timesale/raffle/winnerPop.asp?evt_code=" + this.content.evt_code + "&schedule_idx=" + schedule_idx
                , "winner"
                , "width=800, height=900"
            );
        }
    }
    , watch : {
        timesale_is_write(timesale_is_write){
            if(timesale_is_write) {
                this.now_content = this.content;
            } else {
                this.now_content = {
                    evt_code : null
                    , katalkTemplateNo : ""
                    , katalkTitle : ""
                    , katalkContent : ""
                    , katalkLinkButtonName : ""
                    , katalkLinkUrl : ""
                    , winnerTemplateNo : ""
                    , winnerTitle : ""
                    , winnerContent : ""
                    , winnerLinkButtonName : ""
                    , winnerLinkUrl : ""
                };
            }

            console.log(this.now_content);
        }
        , child_save_flag(){
            if(this.child_save_flag){
                this.$emit("reload", this.content.evt_code);
                this.child_save_flag = false;
            }
        }
    }
});