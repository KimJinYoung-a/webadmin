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
                        <h2>기본 정보</h2>
                        
                        <tr>
                            <th>이벤트 코드</th>
                            <td>
                                <template v-if="timesale_is_write">
                                    <input v-model="now_content.evt_code" type="text" name="evt_code" readonly style="background-color:lightgrey" />
                                </template>
                                <template v-else>
                                    <input v-model="now_content.evt_code" type="text" name="evt_code" />
                                </template>                                
                            </td>
                            
                            <th>레플 여부</th>
                            <td><input type="text" name="raffleFlag" value="Y"  readonly style="background-color:lightgrey" /></td>
                        </tr>
                        <tr>
                            <th>카카오톡 템플릿 번호</th>
                            <td>
                                <input v-model="now_content.katalkTemplateNo" type="text" name="katalkTemplateNo" />
                            </td>
                            
                            <th>카카오톡 제목</th>
                            <td>
                                <input v-model="now_content.katalkTitle" type="text" name="katalkTitle" />
                            </td>
                        </tr>
                        <tr>
                            <th>카카오톡 내용</th>
                            <td colspan="3">
                                <textarea v-model="now_content.katalkContent" rows="4" name="katalkContent" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>카카오톡 링크 버튼명</th>
                            <td colspan="3">
                                <input v-model="now_content.katalkLinkButtonName" type="text" name="katalkLinkButtonName" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>카카오톡 링크 URL</th>
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
                            <th>당첨자 알림톡 템플릿 번호</th>
                            <td>
                                <input v-model="now_content.winnerTemplateNo" type="text" name="winnerTemplateNo" />
                            </td>
                            
                            <th>당첨자 알림톡 제목</th>
                            <td>
                                <input v-model="now_content.winnerTitle" type="text" name="winnerTitle" />
                            </td>
                        </tr>
                        <tr>
                            <th>당첨자 알림톡 내용</th>
                            <td colspan="3">
                                <textarea v-model="now_content.winnerContent" rows="4" name="winnerContent" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>당첨자 알림톡 링크 버튼명</th>
                            <td colspan="3">
                                <input v-model="now_content.winnerLinkButtonName" type="text" name="winnerLinkButtonName" style="width: 100%" />
                            </td>
                        </tr>
                        <tr>
                            <th>당첨자 알림톡 링크 URL</th>
                            <td colspan="3">
                                <input v-model="now_content.winnerLinkUrl" type="text" name="winnerLinkUrl" style="width: 100%" />
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <div style="margin: 30px 0px  30px 740px;">                
                <button @click="$emit('save')" class="button dark">{{timesale_is_write ? '수정' : '저장'}}</button>
                <button @click="$emit('close')" class="button secondary">취소</button>
            </div>
            
            <table class="table table-write table-dark">
                <tbody>
                    <h2>스케쥴 정보</h2>
                    <p style="color: red;"> * 스케쥴정보는 저장후에 등록해주세요.</p>
                    
                    <input type="button" value="추가" @click="add_schedule()" />
                    
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
                        alert("타임딜 이벤트를 먼저 등록해주세요. 상단의 저장버튼을 누르면 등록됩니다.");
                    }
                });
            }else{
                alert("타임딜 이벤트를 먼저 등록해주세요. 상단의 저장버튼을 누르면 등록됩니다.");
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
            if(confirm("삭제하시겠습니까?")){
                callApiHttps("DELETE", "/event/timedeal-schedule", {"evt_code": this.content.evt_code, "schedule_idx":schedule_idx}, function(data){
                    _this.child_save_flag = true;
                    alert("삭제되었습니다.");
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