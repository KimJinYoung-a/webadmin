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
                            <th>컨텐츠명 *</th>
                            <td>
                                <input type="text" name="titlename" class="must" v-model="current_content.titlename" />
                                - 한글 기준 최대 40자까지 입력 가능합니다.
                            </td>
                        </tr>
                        <tr>
                            <th>배경 이미지</th>
                            <td>
                                <label for="addMainimage">등록</label> <br/> - 750*920 크기의 jpg, gif 포맷의 파일만 등록 가능합니다.
                                <div class="thumbnail-area" @click="delete_main_image" v-show="current_content.mainimage">
                                    <img id="showMainImage" :src="current_content.mainimage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input type="text" name="mainimage" v-model="current_content.mainimage" style="display: none;"/>                                
                                <input type="text" name="mainimageChangeF" value="N" style="display: none" />
                            </td>     
                        </tr> 
                        <tr>
                            <th id="tt">컨텐츠 설명</th>
                            <td>
                                <textarea name="contents" style="width: 100%;" rows="3" v-model="current_content.contents"></textarea>
                                - 한글기준 치대 800자까지 입력 가능합니다.
                            </td>
                        </tr>                
                    </tbody>
                </table>
                
                <br/>   
                <h2>게시 설정</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>노출 *</th>
                            <td>
                                <input type="radio" name="isview" value="0" />노출안함
                                <input type="radio" name="isview" value="1" />노출함
                            </td>
                        </tr>
                        <tr>
                            <th>노출 순서</th>
                            <td>
                                <input type="text" name="sortnum" v-model="current_content.sortnum" />
                            </td>
                        </tr>
                    </tbody>
                </table>
                
                <br/> 
                <h2>운영 정보</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>메모</th>
                            <td>
                                <textarea name="ordertext" style="width: 100%;" rows="5" v-model="current_content.ordertext"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>운영설정 *</th>
                            <td>
                                <select name="isusing" class="form-control inline small must" v-model="current_content.isusing">
                                    <option value="">선택</option>
                                    <option value="1">운영중</option>
                                    <option value="0">운영안함</option>
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
            cidx : {type:Number, default:0} // 컨텐츠 idx
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
            if(confirm("제거하시겠습니까?")){
                this.current_content.mainimage = "";
                $("#addMainimage").val("");
            }
        }
    }
    , watch:{
        pop_content(popcontent) { // 컨텐츠 변경 시 현재구분값 set(팝업되었을 때)
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