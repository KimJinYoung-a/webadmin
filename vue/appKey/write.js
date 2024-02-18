Vue.component('APP-KEY-WRITE',{
    template: `
        <div>
            <form id="content_form">
                <input v-if="now_content.idx != 0" name="idx" type="hidden" v-model="content.idx">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>키(생성만 가능)</th>
                            <td>
                                <input v-model="now_content.validationKey" type="text" name="validationKey" class="form-control" />
                            </td>
                        </tr>
                        <tr>
                            <th>설명</th>
                            <td>
                                <input v-model="now_content.description" type="text" name="description" class="form-control" />
                            </td>
                        </tr>                        
                        <tr>
                            <th>사용여부(Y/N)</th>
                            <td>
                                <input v-model="now_content.isusing" type="text" name="isusing" class="form-control" />
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
        </div>
    `
    , data() {return { // 현재 컨텐츠
        now_content : {
            idx : 0
            , validationKey: ''
            , isusing: ''
            , description: ''
        }
    }},
    props: {
        content : {
            idx : {type:Number, default:0} // idx
            , validationKey: {type:String, default:''} // key
            , isusing: {type:String, default:''}
            , description: {type:String, default:''}
        }
    }
    , methods : {
    }
    , watch : {
        content(content){
            console.log("watch", content);
            if(content.idx > 0){
                this.now_content = content;
            }else{
                this.now_content = {
                    idx : 0
                    , validationKey: ''
                    , isusing: ''
                    , description: ''
                }
            }
        }
    }
});