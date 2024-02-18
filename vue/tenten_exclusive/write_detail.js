Vue.component('Tenten-Exclusive-Detail',{
    template: `
        <div>
            <form id="tenten_exclusive_detail">                
                <input :value="now_content.exclusive_idx" type="text" name="exclusive_idx" style="display: none"/>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:30%;">
                        <col style="width:70%">
                    </colgroup>
                    <tbody>
                        <h2>상품정보</h2>                        
                        <tr>
                            <th>상단 메인 이미지</th>
                            <td>
                                <div v-show="now_content.detail_top_img" class="thumbnail-area">
                                    <img id="show_detail_top_img" :src="now_content.detail_top_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image_flag($event, 'detail_top_img')" type="file" name="detail_top_img" id="detail_top_img" />                   
                            </td>
                        </tr>
                        <tr>
                            <th>브랜드명</th>
                            <td>
                                <input v-model="now_content.brand_name" type="text" name="brand_name"/>
                            </td>
                        </tr>
                        <tr>
                            <th>상품명</th>
                            <td>
                                <input v-model="now_content.item_name" type="text" name="item_name" />
                            </td>
                        </tr>
                        
                        <h2>브랜드 정보</h2>
                        <tr>
                            <th>브랜드 이미지</th>
                            <td>
                                <div v-show="now_content.brand_image" class="thumbnail-area">
                                    <img id="show_brand_image" :src="now_content.brand_image" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image_flag($event, 'brand_image')" type="file" name="brand_image" id="brand_image" />
                            </td>
                        </tr>
                        <tr>
                            <th>브랜드 설명</th>
                            <td>
                                <textarea v-model="now_content.brand_content" name="brand_content" style="width: 80%" rows="5"></textarea>
                            </td>
                        </tr>
                        
                        <h2>상품설명</h2>
                        <tr>
                            <th>이미지1</th>
                            <td>
                                <div v-show="now_content.image1" class="thumbnail-area">
                                    <img id="show_image1" :src="now_content.image1" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image_flag($event, 'image1')" type="file" name="image1" id="image1" />
                            </td>
                        </tr>
                        <tr>
                            <th>메인 카피1</th>
                            <td>
                                <textarea v-model="now_content.main_copy1" name="main_copy1" style="width: 80%" rows="2"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>설명1</th>
                            <td>
                                <textarea v-model="now_content.content1" name="content1" style="width: 80%" rows="5"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>롤링 이미지1</th>
                            <td>
                                <div v-show="now_content.rolling_image1" @click="delete_image('rolling_image1')" class="thumbnail-area">
                                    <img id="show_rolling_image1" :src="now_content.rolling_image1" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input @change="change_image_flag($event, 'rolling_image1')" type="file" name="rolling_image1" id="rolling_image1" />
                            </td>
                        </tr>
                        <tr>
                            <th>롤링 이미지2</th>
                            <td>
                                <div v-show="now_content.rolling_image2" @click="delete_image('rolling_image2')" class="thumbnail-area">
                                    <img id="show_rolling_image2" :src="now_content.rolling_image2" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input @change="change_image_flag($event, 'rolling_image2')" type="file" name="rolling_image2" id="rolling_image2" />
                            </td>
                        </tr>
                        <tr>
                            <th>롤링 이미지3</th>
                            <td>
                                <div v-show="now_content.rolling_image3" @click="delete_image('rolling_image3')" class="thumbnail-area">
                                    <img id="show_rolling_image3" :src="now_content.rolling_image3" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input @change="change_image_flag($event, 'rolling_image3')" type="file" name="rolling_image3" id="rolling_image3" />
                            </td>
                        </tr>
                        <tr>
                            <th>롤링 이미지4</th>
                            <td>
                                <div v-show="now_content.rolling_image4" @click="delete_image('rolling_image4')" class="thumbnail-area">
                                    <img id="show_rolling_image4" :src="now_content.rolling_image4" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input @change="change_image_flag($event, 'rolling_image4')" type="file" name="rolling_image4" id="rolling_image4" />
                            </td>
                        </tr>
                        <tr>
                            <th>롤링 이미지5</th>
                            <td>
                                <div v-show="now_content.rolling_image5" @click="delete_image('rolling_image5')" class="thumbnail-area">
                                    <img id="show_rolling_image5" :src="now_content.rolling_image5" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input @change="change_image_flag($event, 'rolling_image5')" type="file" name="rolling_image5" id="rolling_image5" />
                            </td>
                        </tr>
                        <tr>
                            <th>메인 카피2</th>
                            <td>
                                <textarea v-model="now_content.main_copy2" name="main_copy2" style="width: 80%" rows="5"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>설명2</th>
                            <td>
                                <textarea v-model="now_content.content2" name="content2" style="width: 80%" rows="5"></textarea>
                            </td>
                        </tr>
                        
                        <h2>브랜드측 말풍선 (제작 비하인드)</h2>                        
                        <tr>
                            <th>브랜드 프로필 이미지</th>
                            <td>
                                <div v-show="now_content.brand_profile_img" class="thumbnail-area">
                                    <img id="show_brand_profile_img" :src="now_content.brand_profile_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image_flag($event, 'brand_profile_img')" type="file" name="brand_profile_img" id="brand_profile_img" />                   
                            </td>
                        </tr>
                        <tr>
                            <th>닉네임</th>
                            <td>
                                <input v-model="now_content.brand_nickname" type="text" name="brand_nickname" style="width: 80%"/>
                            </td>
                        </tr>
                        <tr>
                            <th>말풍선 내용</th>
                            <td>
                                <textarea v-model="now_content.brand_bubble" name="brand_bubble" style="width: 80%" rows="5"/>
                            </td>
                        </tr>
                        
                        <h2>텐텐측 말풍선 (추천사)</h2>                        
                        <tr>
                            <th>프로필 이미지</th>
                            <td>
                                <div v-show="now_content.tenten_profile_img" class="thumbnail-area">
                                    <img id="show_tenten_profile_img" :src="now_content.tenten_profile_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image_flag($event, 'tenten_profile_img')" type="file" name="tenten_profile_img" id="tenten_profile_img" />                   
                            </td>
                        </tr>
                        <tr>
                            <th>닉네임</th>
                            <td>
                                <input v-model="now_content.tenten_nickname" type="text" name="tenten_nickname" style="width: 80%"/>
                            </td>
                        </tr>
                        <tr>
                            <th>말풍선 내용</th>
                            <td>
                                <textarea v-model="now_content.tenten_bubble" name="tenten_bubble" style="width: 80%" rows="5"/>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <div style="margin: 30px 0px  30px 740px;">
                <button @click="go_detail_save" class="button dark">{{is_written ? '수정' : '저장'}}</button>
                <button @click="$emit('close')" class="button secondary">취소</button>
            </div>
        </div>
    `
    , data(){
        return{
            now_content : {
                detail_top_img : ""
                , brand_name : ""
                , item_name : ""
                , brand_image : ""
                , brand_content : ""
                , image1 : ""
                , main_copy1 : ""
                , content1 : ""
                , main_copy2 : ""
                , content2 : ""
                , rolling_image1 : ""
                , rolling_image2 : ""
                , rolling_image3 : ""
                , rolling_image4 : ""
                , rolling_image5 : ""
                , brand_profile_img : ""
                , brand_nickname : ""
                , brand_bubble : ""
                , tenten_profile_img : ""
                , tenten_nickname : ""
                , tenten_bubble : ""
            }
            , additemid : ""
            , is_saving : false

            , detail_top_image_flag : false
            , brand_image_flag : false
            , image1_image_flag : false
            , rolling_image1_image_flag : false
            , rolling_image2_image_flag : false
            , rolling_image3_image_flag : false
            , rolling_image4_image_flag : false
            , rolling_image5_image_flag : false
            , brand_profile_image_flag : false
            , tenten_profile_image_flag : false
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
            reader.onload = function(e){
                if(type == "detail_top_img"){
                    _this.now_content.detail_top_img = e.target.result;
                    _this.detail_top_image_flag = true;
                }else if(type == "brand_image"){
                    _this.now_content.brand_image = e.target.result;
                    _this.brand_image_flag = true;
                }else if(type == "image1"){
                    _this.now_content.image1 = e.target.result;
                    _this.image1_image_flag = true;
                }else if(type == "rolling_image1"){
                    _this.now_content.rolling_image1 = e.target.result;
                    _this.rolling_image1_image_flag = true;
                }else if(type == "rolling_image2"){
                    _this.now_content.rolling_image2 = e.target.result;
                    _this.rolling_image2_image_flag = true;
                }else if(type == "rolling_image3"){
                    _this.now_content.rolling_image3 = e.target.result;
                    _this.rolling_image3_image_flag = true;
                }else if(type == "rolling_image4"){
                    _this.now_content.rolling_image4 = e.target.result;
                    _this.rolling_image4_image_flag = true;
                }else if(type == "rolling_image5"){
                    _this.now_content.rolling_image5 = e.target.result;
                    _this.rolling_image5_image_flag = true;
                }else if(type == "brand_profile_img"){
                    _this.now_content.brand_profile_img = e.target.result;
                    _this.brand_profile_image_flag = true;
                }else if(type == "tenten_profile_img"){
                    _this.now_content.tenten_profile_img = e.target.result;
                    _this.tenten_profile_image_flag = true;
                }
            }
        }
        , go_detail_save(){
            const _this = this;

            if(this.is_saving){
                return false;
            }

            this.is_saving = true;

            let form_data = $("#tenten_exclusive_detail").serialize();
            let file1 = document.getElementById("detail_top_img").files[0];
            let file2 = document.getElementById("brand_image").files[0];
            let file3 = document.getElementById("image1").files[0];
            let file4 = document.getElementById("rolling_image1").files[0];
            let file5 = document.getElementById("rolling_image2").files[0];
            let file6 = document.getElementById("rolling_image3").files[0];
            let file7 = document.getElementById("rolling_image4").files[0];
            let file8 = document.getElementById("rolling_image5").files[0];
            let file9 = document.getElementById("brand_profile_img").files[0];
            let file10 = document.getElementById("tenten_profile_img").files[0];

            this.save_detail_image(file1, file2, file3, file4, file5, file6, file7, file8, file9, file10).then(function(data){
                form_data += "&detail_top_img=" + data.detail_top_img;
                form_data += "&brand_image=" + data.brand_image;
                form_data += "&image1=" + data.image1;
                form_data += "&rolling_image1=" + data.rolling_image1;
                form_data += "&rolling_image2=" + data.rolling_image2;
                form_data += "&rolling_image3=" + data.rolling_image3;
                form_data += "&rolling_image4=" + data.rolling_image4;
                form_data += "&rolling_image5=" + data.rolling_image5;
                form_data += "&brand_profile_img=" + data.brand_profile_img;
                form_data += "&tenten_profile_img=" + data.tenten_profile_img;
                console.log("form_data", form_data);

                callApiHttps("post", "/tenten-exclusive/item-detail", form_data, function(data){
                    alert("저장되었습니다.");
                    _this.$emit('close');
                    _this.is_saving = false;
                });
            });
        }
        , save_detail_image(file1, file2, file3, file4, file5, file6, file7, file8, file9, file10){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                imgData.append("folderName", "detail");
                if(_this.detail_top_image_flag){
                    imgData.append('photo1', file1);
                }
                if(_this.brand_image_flag){
                    imgData.append('photo2', file2);
                }
                if(_this.image1_image_flag){
                    imgData.append('photo3', file3);
                }
                if(_this.rolling_image1_image_flag){
                    imgData.append('photo4', file4);
                }
                if(_this.rolling_image2_image_flag){
                    imgData.append('photo5', file5);
                }
                if(_this.rolling_image3_image_flag){
                    imgData.append('photo6', file6);
                }
                if(_this.rolling_image4_image_flag){
                    imgData.append('photo7', file7);
                }
                if(_this.rolling_image5_image_flag){
                    imgData.append('photo8', file8);
                }
                if(_this.brand_profile_image_flag){
                    imgData.append('photo9', file9);
                }
                if(_this.tenten_profile_image_flag){
                    imgData.append('photo10', file10);
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/tenten_exclusive/tenten_exclusive_detail_reg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);

                        if (response.response === 'ok') {
                            let image_url = {
                                "detail_top_img"  : response.photo1 ? response.photo1 : _this.now_content.detail_top_img
                                , "brand_image" : response.photo2 ? response.photo2 : _this.now_content.brand_image
                                , "image1" : response.photo3 ? response.photo3 : _this.now_content.image1
                                , "rolling_image1" : response.photo4 ? response.photo4 : _this.now_content.rolling_image1
                                , "rolling_image2" : response.photo5 ? response.photo5 : _this.now_content.rolling_image2
                                , "rolling_image3" : response.photo6 ? response.photo6 : _this.now_content.rolling_image3
                                , "rolling_image4" : response.photo7 ? response.photo7 : _this.now_content.rolling_image4
                                , "rolling_image5" : response.photo8 ? response.photo8 : _this.now_content.rolling_image5
                                , "brand_profile_img" : response.photo9 ? response.photo9 : _this.now_content.brand_profile_img
                                , "tenten_profile_img" : response.photo10 ? response.photo10 : _this.now_content.tenten_profile_img
                            };

                            console.log("image_url", image_url);
                            return resolve(image_url);
                        } else if(response.response === 'none'){
                            let image_url = {
                                "detail_top_img"  : _this.now_content.detail_top_img
                                , "brand_image" : _this.now_content.brand_image
                                , "image1" : _this.now_content.image1
                                , "rolling_image1" : _this.now_content.rolling_image1
                                , "rolling_image2" : _this.now_content.rolling_image2
                                , "rolling_image3" : _this.now_content.rolling_image3
                                , "rolling_image4" : _this.now_content.rolling_image4
                                , "rolling_image5" : _this.now_content.rolling_image5
                                , "brand_profile_img" : _this.now_content.brand_profile_img
                                , "tenten_profile_img" : _this.now_content.tenten_profile_img
                            };

                            return resolve(image_url);
                        }else {
                            alert('이미지 저장 중 오류가 발생했습니다. (Err: 001)');
                            return reject();
                        }
                    }
                    , error : function (e){
                        console.log(e);

                        return reject();
                    }
                });
            });
        }
        , delete_image(type){
            if(confirm("제거하시겠습니까?")){
                this.$emit("change_image_flag", type);
                if(type == "rolling_image1"){
                    this.now_content.rolling_image1 = "";
                    $("#rolling_image1").val("");
                }else if(type == "rolling_image2"){
                    this.now_content.rolling_image2 = "";
                    $("#rolling_image2").val("");
                }else if(type == "rolling_image3"){
                    this.now_content.rolling_image3 = "";
                    $("#rolling_image3").val("");
                }else if(type == "rolling_image4"){
                    this.now_content.rolling_image4 = "";
                    $("#rolling_image2").val("");
                }else if(type == "rolling_image5"){
                    this.now_content.rolling_image5 = "";
                    $("#rolling_image2").val("");
                }
            }
        }
    }
    , watch : {
        content(content){
            this.now_content = {
                detail_top_img : ""
                , brand_name : ""
                , item_name : ""
                , brand_image : ""
                , brand_content : ""
                , image1 : ""
                , main_copy1 : ""
                , content1 : ""
                , main_copy2 : ""
                , content2 : ""
                , rolling_image1 : ""
                , rolling_image2 : ""
                , rolling_image3 : ""
                , rolling_image4 : ""
                , rolling_image5 : ""
                , brand_profile_img : ""
                , brand_nickname : ""
                , brand_bubble : ""
                , tenten_profile_img : ""
                , tenten_nickname : ""
                , tenten_bubble : ""
            }

            this.now_content = content;
        }
    }
});