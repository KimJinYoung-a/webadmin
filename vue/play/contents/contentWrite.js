Vue.component('Content-Write',{
    template: `
        <div style="height: 600px; overflow-y:auto;">
            <form name="play_content" id="play_content" enctype="multipart/form-data">
                <input v-if="current_content.cidx != 0" name="cidx" type="hidden" v-model="current_content.cidx">
                <table class="table table-write table-dark">
                    <colgroup>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>�������� *</th>
                            <td>
                                <input type="text" name="titlename" class="must" v-model="current_content.titlename" />
                                - �ѱ� ���� �ִ� 40�ڱ��� �Է� �����մϴ�.
                            </td>
                        </tr>
                        <tr>
                            <th>��� �̹���</th>
                            <td>
                                <label for="addMainimage">���</label> <br/> - 750*920 ũ���� jpg, gif ������ ���ϸ� ��� �����մϴ�.
                                <div class="thumbnail-area" @click="delete_main_image" v-show="current_content.mainimage">
                                    <img id="showMainImage" :src="current_content.mainimage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">����</div>
                                </div>
                                <input type="text" name="mainimage" v-model="current_content.mainimage" style="display: none;"/>                                
                                <input type="text" name="mainimageChangeF" value="N" style="display: none" />
                            </td>     
                        </tr> 
                        <tr>
                            <th id="tt">������ ����</th>
                            <td>
                                <textarea name="contents" style="width: 100%;" rows="3" v-model="current_content.contents"></textarea>
                                - �ѱ۱��� ġ�� 800�ڱ��� �Է� �����մϴ�.
                            </td>
                        </tr>                
                    </tbody>
                </table>
                
                <br/>   
                <h2>�Խ� ����</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>���� *</th>
                            <td>
                                <input type="radio" name="isview" value="0" />�������
                                <input type="radio" name="isview" value="1" />������
                            </td>
                        </tr>
                        <tr>
                            <th>���� ����</th>
                            <td>
                                <input type="text" name="sortnum" v-model="current_content.sortnum" />
                            </td>
                        </tr>
                    </tbody>
                </table>
                
                <br/> 
                <h2>� ����</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>�޸�</th>
                            <td>
                                <textarea name="ordertext" style="width: 100%;" rows="5" v-model="current_content.ordertext"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>����� *</th>
                            <td>
                                <select name="isusing" class="form-control inline small must" v-model="current_content.isusing">
                                    <option value="">����</option>
                                    <option value="1">���</option>
                                    <option value="0">�����</option>
                                </select>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <input name="addMainimage" id="addMainimage" @change="change_main_image" type="file" class="form-control inline middle" style="display: none;" />
        </div>
    `,
    data() {
        return {
            current_content : {
                cidx : 0
                , isusing : 0
                , changeflag : false
            }
            , additemimage: ""
        }
    },
    props: {
        pop_content : {
            cidx : {type:Number, default:0} // ������ idx
        }
    },
    methods : {
        change_main_image(image){
            const _this = this;
            const file = image.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = function(e){
                _this.current_content.mainimage = e.target.result;
            }

            $("input[name=mainimageChangeF]").val("Y");
        }
        , delete_main_image(){
            if(confirm("�����Ͻðڽ��ϱ�?")){
                this.current_content.mainimage = "";
                $("#addMainimage").val("");
            }
        }
    }
    , watch:{
        pop_content(popcontent) { // ������ ���� �� ���籸�а� set(�˾��Ǿ��� ��)
            $("input[name=mainimageChangeF]").val("N");

            if(popcontent.cidx != null){
                console.log("popcontent", popcontent);
                this.current_content = popcontent;
                this.current_content.changeflag = false;

                const isview = this.current_content.isview;

                $("input[name=isview]").each(function(){
                    if(this.value == isview){
                        this.checked = true;
                    }
                });
            }else{
                this.current_content = {
                    cidx : 0
                    , isusing : 0
                };

                $("input[name=isview]:first").prop("checked", true);
                $("#thumnailDiv").empty();
            }
        }
    }
});