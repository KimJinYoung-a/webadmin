Vue.component("NORMAL", {
    template : `
        <form id="normal_form">
            <table class="table table-write table-dark" style="margin-top: 15px;">
                <input type="text" :value="evt_code" name="evt_code" style="display: none"/>
                <input type="text" :value="schedule_idx" name="schedule_idx" style="display: none"/>
                <tbody>
                    <tr>
                        <th>상품 코드</th>
                        <td>
                            <input type="text" v-model="itemid" name="itemid" class="must"/>                                
                        </td>
                        
                        <th>정렬 순서</th>
                        <td>
                            <input v-model="sortNo" type="text" name="sortNo" class="must"/>                                
                        </td>
                    </tr>     
                    <!--<tr>
                        <th>커스텀 상품명</th>
                        <td>
                            <input type="text" name="custom_name"/>                                
                        </td>
                        
                        <th>커스텀 이미지</th>
                        <td>
                            <input type="file" name="custom_image"/>                                
                        </td>
                    </tr>   -->
                </tbody>
            </table>
            
            <div style="margin: 15px 0 0 450px;">
                <button @click="go_save_normal" type="button" class="button dark">저장</button>
                <button @click="$emit('close')" type="button" class="button secondary">취소</button>
            </div>
        </form>
    `
    , props : {
        evt_code : {type:String, default:""}
        , schedule_idx : {type:Number, default:1}
        , normal_list : {type:Array, default:[]}
    }
    , data(){
        return{
            itemid : ""
            , sortNo : ""
        }
    }
    , methods : {
        go_save_normal(){
            const _this = this;
            let form_data = $("#normal_form").serialize();
            const result = this.check_normal_dup(this.itemid);
            console.log("cc", result);
            if(result){
                callApiHttps("POST", "/event/timedeal-normal-one", form_data, function(data){
                    alert("저장되었습니다.");
                    _this.$emit("close");
                    _this.$emit("reload");
                });
            }else{
                alert("이미 해당 상품이 등록되어있습니다.");
            }
        }
        , check_normal_dup(itemid){
            let check_ok = true;
            this.normal_list.forEach(function (item){
                if(item.itemid == itemid){
                    check_ok = false;
                }
            });

            return check_ok;
        }
    }
});