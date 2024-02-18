Vue.component('Tenten-Exclusive-Main',{
    template: `
        <div>
            <form id="tenten_exclusive_main">                
                <input :value="now_content.exclusive_idx" type="text" name="exclusive_idx" style="display: none"/>
                <input :value="now_content.itemid" type="text" name="itemid" style="display: none"/>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:30%;">
                        <col style="width:70%">
                    </colgroup>
                    <tbody>
                        <h2>상품정보</h2>
                        
                        <tr>
                            <th>상품 썸네일</th>
                            <td>
                                <div v-show="now_content.item_img" class="thumbnail-area">
                                    <img id="show_item_img" :src="now_content.item_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image_flag($event, 'item_img')" type="file" name="item_img" id="item_img" />                   
                            </td>
                        </tr>
                        <tr>
                            <th>메인 카피</th>
                            <td>
                                <textarea v-model="now_content.head_title" name="head_title" style="width: 80%" rows="2"/>
                            </td>
                        </tr>
                        <tr>
                            <th>설명</th>
                            <td>
                                <textarea v-model="now_content.head_contents" name="head_contents" style="width: 80%" rows="5"></textarea>
                            </td>
                        </tr>
                        
                        <input :value="now_content.question_idx" name="question_idx" type="text" style="display: none" />
                        <h2>투표 및 코멘트</h2>
                        <tr>
                            <th>투표 질문</th>
                            <td>
                                <textarea v-model="now_content.question_contents" name="question_contents"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>선택지1</th>
                            <td>
                                <input v-model="now_content.choice_contents_1" type="text" name="choice_contents_1" />
                            </td>
                        </tr>
                        <tr>
                            <th>선택지2</th>
                            <td>
                                <input v-model="now_content.choice_contents_2" type="text" name="choice_contents_2" />
                            </td>
                        </tr>
                        <tr>
                            <th>선택지3</th>
                            <td>
                                <input v-model="now_content.choice_contents_3" type="text" name="choice_contents_3" />
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <div style="margin: 30px 0px  30px 740px;">                
                <button @click="$emit('save')" class="button dark">{{is_written ? '수정' : '저장'}}</button>
                <button @click="$emit('close')" class="button secondary">취소</button>
            </div>
        </div>
    `
    , data(){
        return{
            now_content : {
                exclusive_idx : ""
                , itemid : ""
                , item_img : ""
                , head_title : ""
                , head_contents : ""
                , question_idx : ""
                , question_contents : ""
                , choice_contents_1 : ""
                , choice_contents_2 : ""
                , choice_contents_3 : ""
            }
            , additemid : ""
        }
    }
    , props : {
        content : {
            exclusive_idx : {type:String, default:""}
            , itemid : {type:String, default:""}
            , item_img : {type:String, default:""}
            , head_title : {type:String, default:""}
            , head_contents : {type:String, default:""}
            , question_idx : {type:String, default:""}
            , question_contents : {type:String, default:""}
        }
        , is_written : {type:Boolean, default : false}
    }
    , mounted(){

    }
    , methods : {
        change_image_flag(event, type){
            const _this = this;
            const file = event.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            this.$emit("change_image_flag", type);
            reader.onload = function(e){
                if(type == "item_img"){
                    _this.now_content.item_img = e.target.result;
                }
            }
        }
    }
    , watch : {
        content(content){
            this.now_content = {
                itemid : ""
                , item_img : ""
                , head_title : ""
                , head_contents : ""
                , question_idx : ""
                , question_contents : ""
            }

            this.now_content = content;
        }
    }
});