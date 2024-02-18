Vue.component('Tenten-Exclusive-Write',{
    template: `
        <div>
            <form id="tenten_exclusive_write">                
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:30%;">
                        <col style="width:70%">
                    </colgroup>
                    <tbody>
                        <input :value="now_content.exclusive_idx" type="text" name="exclusive_idx" style="display: none" />
                        
                        <h2>상품등록</h2>                        
                        <tr>
                            <th>상품정보</th>
                            <td>
                                <input v-model="now_content.itemid" type="text" name="itemid" id="itemid" class="must" />
                                <input v-if="!content.itemid" @click="show_item_search" type="button" value="상품 검색" />                                
                            </td>
                        </tr>
                        <tr>
                            <th>상품 오픈일</th>
                            <td>
                                시작일
                                <input v-model="open_day" type="text" name="open_day" id="open_day" style="width: 90px;" class="must" />
                                <input v-model="open_time" type="text" name="open_time" id="open_time" required size="8" class="must" />
                                <input v-model="now_content.open_date" type="text" name="open_date" id="open_date" style="display: none" />
                                <br/><p style="color: red">단독페이지에 오픈되는 날짜를 지정해주세요</p>
                            </td>
                        </tr>
                        <tr>
                            <th>프론트 노출여부</th>
                            <td>
                                <input v-model="now_content.display_yn" type="radio" name="display_yn" value="Y" />Y
                                <input v-model="now_content.display_yn" type="radio" name="display_yn" value="N" />N
                            </td>
                        </tr>
                        <tr>
                            <th>비고</th>
                            <td>
                               <textarea name="etc" style="width: 80%" rows="5"></textarea>  
                            </td>
                        </tr>
                        
                        <h2>판매예고(판매예정 탭)</h2>
                        <tr>
                            <th>노출여부</th>
                            <td>
                                <input v-model="now_content.pre_display_yn" type="radio" name="pre_display_yn" value="Y" />Y
                                <input v-model="now_content.pre_display_yn" type="radio" name="pre_display_yn" value="N" />N
                            </td>
                        </tr>
                        <tr>
                            <th>노출일자</th>
                            <td>
                                시작일
                                <input v-model="pre_open_day" type="text" name="pre_open_day" id="pre_open_day" style="width: 90px;" />
                                <input v-model="pre_open_time" type="text" name="pre_open_time" id="pre_open_time" required size="8" />
                                <input v-model="now_content.pre_open_date" type="text" name="pre_open_date" id="pre_open_date" style="display: none"/>
                                <br/><p style="color: red">오픈일이 되면 자동으로 판매예정 탭에서 제거되고, 판매 중인 탭에서 노출됩니다.</p>
                            </td>
                        </tr>
                        <tr>
                            <th>상품 이미지</th>
                            <td>
                                <div v-show="now_content.pre_img" class="thumbnail-area">
                                    <img id="show_pre_img" :src="now_content.pre_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />                                    
                                </div>
                                <input @change="change_image_flag($event, 'pre_img')" type="file" name="pre_img" id="pre_img" />                                
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <div style="margin: 30px 0px  30px 740px;">                
                <button @click="go_save" class="button dark">{{is_written ? '수정' : '저장'}}</button>
                <button @click="$emit('close')" class="button secondary">취소</button>
            </div>
        </div>
    `
    , data(){
        return{
            now_content : {
                itemid : ""
                , open_date : ""
                , display_yn : "Y"
                , etc : ""
                , pre_display_yn : "Y"
                , pre_open_date : ""
                , pre_img : ""
                , exclusive_idx : null
            }
            , open_day : ""
            , open_time : ""
            , pre_open_day : ""
            , pre_open_time : ""

            , is_saving : false
            , pre_image_flag : false
        }
    }
    , props : {
        content : {
            itemid : {type:String, default:""}
            , open_date : {type:String, default:""}
            , display_yn : {type:String, default:"Y"}
            , etc : {type:String, default:""}
            , pre_display_yn : {type:String, default:"Y"}
            , pre_open_date : {type:String, default:""}
            , pre_img : {type:String, default:""}
            , exclusive_idx : {type:String, default: null}
        }
        , is_written : {type:Boolean, default : false}
    }
    , mounted(){
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#tenten_exclusive_write #open_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const max_date = $(this).datepicker("getDate");
                $("#pre_open_day").datepicker('option', "maxDate", max_date);

                _this.open_day = document.getElementById("open_day").value;
                _this.now_content.open_date = _this.open_day + " " + _this.open_time;
            }
        });
        $("#open_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.open_time = document.getElementById("open_time").value;

                _this.now_content.open_date = _this.open_day + " " + _this.open_time;
            }
        });

        $("#tenten_exclusive_write #pre_open_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.pre_open_day = document.getElementById("pre_open_day").value;
                _this.now_content.pre_open_date = _this.pre_open_day + " " + _this.pre_open_time;
            }
        });
        $("#pre_open_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.pre_open_time = document.getElementById("pre_open_time").value;
                _this.now_content.pre_open_date = _this.pre_open_day + " " + _this.pre_open_time;
            }
        });
    }
    , methods : {
        show_item_search(){
            window.open("/admin/eventmanage/tenten_exclusive/pop_singleItemSelect.asp?target=tenten_exclusive_write&ptype=piece&itemarr=", "연관상품", "width:300px, height:200px");
        }
        , change_image_flag(event, type){
            const _this = this;
            const file = event.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            //this.$emit("change_image_flag", type);
            this.pre_image_flag = true;
            reader.onload = function(e){
                _this.now_content.pre_img = e.target.result;
            }
        }
        , go_save(){
            const _this = this;

            let checkOk = true;
            $("#tenten_exclusive_write .must").each(function(){
                if($(this).val().trim() == ""){
                    checkOk = false;
                    let th_name = $(this).parent().parent().find("th")[0].innerText;
                    alert("필수항목 " + th_name + "를 입력하지 않으셨습니다.");
                    $(this).focus();

                    return false;
                }
            });

            if(checkOk){
                //this.$emit('save');
                if(this.is_saving){
                    return false;
                }

                this.is_saving = true;
                let form_data = $("#tenten_exclusive_write").serialize();
                let file1 = document.getElementById("pre_img").files[0];

                this.save_image(file1).then(function(data){
                    if(data){
                        form_data += "&pre_img=" + data.photo1;
                    }
                    console.log("form_data", form_data);

                    callApiHttps("post", "/tenten-exclusive/item", form_data, function(data){
                        alert("저장되었습니다.");

                        _this.$emit('close');
                        _this.$emit('reload');

                        _this.is_saving = false;
                        _this.reset_write_data();
                    });
                });
            }
        }
        , save_image(file1){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                if(_this.pre_image_flag){
                    imgData.append('photo1', file1);
                    imgData.append("folderName", "pre_img");
                }

                if(_this.item_image_flag){
                    imgData.append('photo1', file1);
                    imgData.append("folderName", "item_img");
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/tenten_exclusive/tenten_exclusive_reg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);

                        if (response.response === 'ok') {
                            return resolve(response);
                        } else if(response.response === 'none'){
                            return resolve();
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
        , reset_write_data(){
            this.now_content = {
                itemid : ""
                , open_date : ""
                , display_yn : "Y"
                , etc : ""
                , pre_display_yn : "Y"
                , pre_open_date : ""
                , pre_img : ""
                , exclusive_idx : null
            };
            this.open_day = "";
            this.open_time = "";
            this.pre_open_day = "";
            this.pre_open_time = "";

            $("#pre_img").val("");
        }
    }
    , watch : {
        is_written(is_written){
            if(is_written) {
                this.now_content = this.content;

                this.open_day = this.content.open_date.split(" ")[0];
                this.open_time = this.content.open_date.split(" ")[1].substring(0, 5);
                this.pre_open_day = this.content.pre_open_date.split(" ")[0];
                this.pre_open_time = this.content.pre_open_date.split(" ")[1].substring(0, 5);
            } else {
                this.reset_write_data();
            }
        }
        , content(content){
            this.reset_write_data();

            this.now_content = this.content;
            this.open_day = this.content.open_date.split(" ")[0];
            this.open_time = this.content.open_date.split(" ")[1].substring(0, 5);
            this.pre_open_day = this.content.pre_open_date.split(" ")[0];
            this.pre_open_time = this.content.pre_open_date.split(" ")[1].substring(0, 5);
        }
    }
});