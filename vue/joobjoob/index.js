var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div>
            <form id="my_form">
                <table class="table table-dark table-search table-write">
                    <colgroup>
                        <col style="width:30%">
                        <col style="width:70%">
                    </colgroup>
                    <tbody>
                        <p>기본설정</p>
                        <tr>
                            <th>APP 이벤트 코드</th>
                            <td>
                                <input type="text" name="evt_code" v-model="evt_code" readonly style="background: lightgrey;" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>줍줍기간설정</th>
                            <td>
                                <input type="text" id="open_day" :value="open_day" style="float: left; width: 90px;" class="must"/>
                                <input type="text" id="open_time" v-model="open_time" required size="8" style="float: left;" class="must"/>
                                <input type="text" name="open_date" v-model="now_content.open_date" class="must" style="display: none"/>
                                <p style="float: left;">&nbsp ~ &nbsp</p>
                                <input type="text" id="end_day" v-model="end_day" style="float: left; width: 90px;" class="must"/>
                                <input type="text" id="end_time" v-model="end_time" required size="8" style="float: left;" class="must"/>
                                <input type="text" name="end_date" v-model="now_content.end_date" class="must" style="display: none"/>
                            </td>
                        </tr>
                        <tr>
                            <th>리스트 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.listimg" class="thumbnail-area">
                                    <img id="showListImage" :src="now_content.listimg" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                
                                <input type="file" @change="change_image_flag('listimg', $event)" name="listimg" id="listimg"/>
                            </td>
                        </tr>
                        <tr>
                            <th>메인 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.mainimg" class="thumbnail-area">
                                    <img id="showMainImage" :src="now_content.mainimg" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('mainimg', $event)" name="mainimg" id="mainimg"/>
                            </td>
                        </tr>
                        <tr>
                            <th>모바일웹메인 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.mainimg2" class="thumbnail-area">
                                    <img id="showMainImage2" :src="now_content.mainimg2" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('mainimg2', $event)" name="mainimg2" id="mainimg2"/>
                            </td>
                        </tr>
                        <tr>
                            <th>PC 이벤트 코드</th>
                            <td>
                                <input type="text" name="link_code" v-model="now_content.link_code" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>APP 설치 링크</th>
                            <td>
                                <input type="text" name="applink" v-model="now_content.applink" style="width: 100%" class="must" />
                            </td>
                        </tr>
                        
                        <br/>
                        <p>카카오톡 공유하기 설정</p>
                        
                        <tr>
                            <th>카카오톡 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.snsimg" class="thumbnail-area">
                                    <img id="showSnsImage" :src="now_content.snsimg" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('snsimg', $event)" name="snsimg" id="snsimg"/>
                            </td>
                        </tr>
                        <tr>
                            <th>카카오톡 제목</th>
                            <td>
                                <input type="text" name="snstitle" v-model="now_content.snstitle" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>카카오톡 내용</th>
                            <td>
                                <textarea name="snstext" v-model="now_content.snstext" @input="change_textarea('sns')" rows="4" style="width: 100%" class="must"/>
                                {{snstext_length}} / 128
                            </td>
                        </tr>
                        
                        <br/>
                        <p>푸시 설정</p>
                        
                        <tr>
                            <th>푸시 제목</th>
                            <td>
                                <input type="text"  name="pushTitle" v-model="now_content.pushTitle" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>푸시 내용</th>
                            <td>
                                <textarea name="pushText" v-model="now_content.pushText" @input="change_textarea('push')" rows="4" style="width: 100%" class="must"/>
                                {{pushtext_length}} / 128
                            </td>
                        </tr>
                        
                        <br/>
                        <p>당첨 설정</p>
                        
                        <tr>
                            <th>상품명</th>
                            <td>
                                <input type="text" name="option1" v-model="now_content.option1" style="width:100%;" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>상품코드</th>
                            <td>
                                <input type="text" name="option2" v-model="now_content.option2" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>당첨 수량(고정값)</th>
                            <td>
                                <input type="number" name="itemea" v-model="now_content.itemea" @input="change_itemea" min="0" class="must" />
                            </td>
                        </tr>
                        <tr v-if="is_write > 0">
                            <th>재고 수량(유동값)</th>
                            <td>
                                <input type="number" name="option4" v-model="now_content.option4" min="0" class="must" />
                            </td>
                        </tr>
                        <tr>
                            <th>판매 가격</th>
                            <td>
                                <input type="text" name="option6" v-model="now_content.option6" @input="change_price('option6')" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>당첨 시 구매 가격</th>
                            <td>
                                <input type="text" name="option7" v-model="now_content.option7" @input="change_price('option7')" class="must"/>
                            </td>
                        </tr>
                        
                        <tr>
                            <th>상품상세명</th>
                            <td>
                                <input type="text" name="option8" v-model="now_content.option8" style="width:100%;" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>추가 유의사항</th>
                            <td>
                                <input type="text" name="etcText" v-model="now_content.etcText" style="width:100%;"/>
                            </td>
                        </tr>
                        <tr>
                            <th>당첨 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.prizeimg" class="thumbnail-area">
                                    <img id="showPrizeImage" :src="now_content.prizeimg" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('prizeimg', $event)" name="prizeimg" id="prizeimg"/>
                            </td>
                        </tr>
                        <tr>
                            <th>당첨 시간</th>
                            <td>
                                <table class="table table-dark">
                                    <thead>
                                        <tr>
                                            <th>순번</th>
                                            <th>당첨일시</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr v-for="item in (1, parseInt(now_content.itemea))">
                                            <td>{{item}}</td>
                                            <td>
                                                <input type="text" :id="'prizedate' + item" style="width: 90px;" class="prizedate must" /> 
                                                <input type="text" :id="'prizetime' + item" required size="8" class="prizetime must" />
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </tbody>                    
                </table>
            </form>
            <div style="float: right; padding-top: 10px;">
                <button v-if="check_resetable()" @click="go_reset()" class="button dark">초기화</button>
                <button @click="go_save()" class="button dark">저장</button>
                <button onclick="self.close()" class="button secondary">취소</button>
            </div>
        </div>
    `
    , data() {
        return{
            evt_code : ""
            , itemea_count : 0
            , now_content : {
                listimg : ""
                , mainimg : ""
                , mainimg2 : ""
                , snsimg : ""
                , prizeimg: ""
                , open_date : ""
                , end_date : ""
                , itemea : 0
            }
            , snstext_length : 0
            , pushtext_length : 0
            , open_day : ""
            , open_time : ""
            , end_day : ""
            , end_time : ""
            , image_flag : {
                listimg_f : false
                , mainimg_f : false
                , snsimg_f : false
                , prizeimg_f : false
            }
            , check_ok : true
        }
    }
    , created() {
        let query_param = new URLSearchParams(window.location.search);
        this.evt_code = query_param.get("evt_code");

        this.now_content.open_date = query_param.get("open_date");
        this.open_day = query_param.get("open_date").split(" ")[0];
        this.open_time = query_param.get("open_date").split(" ")[1].substr(0,5);
        this.now_content.end_date = query_param.get("end_date");
        this.end_day = query_param.get("end_date").split(" ")[0];
        this.end_time = query_param.get("end_date").split(" ")[1].substr(0,5);

        this.$store.dispatch("GET_IS_WRITE", this.evt_code);

        this.check_resetable();
    }
    , mounted(){
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#open_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#end_date").datepicker('setDate', min_date);
                $("#end_date").datepicker('option', "minDate", min_date);

                _this.open_day = document.getElementById("open_day").value;
            }
        });
        $("#end_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.end_day = document.getElementById("end_day").value;
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
            }
        });
        $("#end_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.end_time = document.getElementById("end_time").value;
            }
        });
    }
    , computed: {
        is_write(){
            return this.$store.getters.is_write;
        }
        , content(){
            return this.$store.getters.content;
        }
        , post_snstext_length(){
            return this.$store.getters.snstext_length;
        }
        , post_pushtext_length(){
            return this.$store.getters.pushtext_length;
        }
    }
    , methods: {
        go_save(){
            const _this = this;
            this.form_validate();
            if(this.check_ok){
                return new Promise(function (resolve, reject) {
                    _this.now_content.open_date = _this.open_day + " " + _this.open_time;
                    _this.now_content.end_date = _this.end_day + " " + _this.end_time;

                    _this.now_content.option6 = _this.now_content.option6.toString().replace(/\D/g, "");
                    _this.now_content.option7 = _this.now_content.option7.toString().replace(/\D/g, "");

                    return resolve();
                }).then(function(){
                    let form_data = $("#my_form").serializeArray();
                    let json_data = {};
                    $.each(form_data, function(i, e){
                        if (json_data[e.name]) {
                            if (!json_data[e.name].push) {
                                json_data[e.name] = [json_data[e.name]];
                            }
                            json_data[e.name].push(e.value || '');
                        } else {
                            json_data[e.name] = e.value || '';
                        }
                    });

                    let joobjoobTime = new Array();
                    for(let i = 1; i <= parseInt(_this.itemea_count); i++){
                        let joobjoobObj = {};
                        joobjoobObj["prizedate"] = $("#prizedate" + i).val();
                        joobjoobObj["prizetime"] = $("#prizetime" + i).val();
                        joobjoobTime.push(joobjoobObj)
                    }
                    json_data["joobjoobTime"] = joobjoobTime;

                    let file1 = document.getElementById("listimg").files[0];
                    let file2 = document.getElementById("mainimg").files[0];
                    let file3 = document.getElementById("mainimg2").files[0];
                    let file4 = document.getElementById("snsimg").files[0];
                    let file5 = document.getElementById("prizeimg").files[0];

                    _this.save_image(file1, file2, file3, file4, file5).then(function () {
                        json_data["listimg"] = _this.now_content.listmig;
                        json_data["mainimg"] = _this.now_content.mainimg;
                        json_data["mainimg2"] = _this.now_content.mainimg2;
                        json_data["snsimg"] = _this.now_content.snsimg;
                        json_data["prizeimg"] = _this.now_content.prizeimg;

                        //console.log("form_data", form_data);
                        if (_this.is_write > 0) {
                            callApiHttps("PUT", "/event/joobjoob", json_data, function (data) {
                                alert("수정되었습니다.");
                                self.close();
                            });
                        } else {
                            callApiHttps("POST", "/event/joobjoob", json_data, function (data) {
                                alert("등록되었습니다.");
                                self.close();
                            });
                        }
                    });
                });
            }
        }
        , save_image(file1, file2, file3, file4, file5){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                if(_this.image_flag.listimg_f){
                    imgData.append('imgFile1', file1);
                }
                if(_this.image_flag.mainimg_f){
                    imgData.append('imgFile2', file2);
                }
                if(_this.image_flag.mainimg2_f){
                    imgData.append('imgFile3', file3);
                }
                if(_this.image_flag.snsimg_f){
                    imgData.append('imgFile4', file4);
                }
                if(_this.image_flag.prizeimg_f){
                    imgData.append('imgFile5', file5);
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/event_admin/joobjoob_admin_multi_imgreg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);

                        _this.now_content.listmig = response.imgurl1 ? response.imgurl1 : _this.now_content.listimg;
                        _this.now_content.mainimg = response.imgurl2 ? response.imgurl2 : _this.now_content.mainimg;
                        _this.now_content.mainimg2 = response.imgurl3 ? response.imgurl3 : _this.now_content.mainimg2;
                        _this.now_content.snsimg = response.imgurl4 ? response.imgurl4 : _this.now_content.snsimg;
                        _this.now_content.prizeimg = response.imgurl5 ? response.imgurl5 : _this.now_content.prizeimg;

                        return resolve();
                    }
                    , error : function (request,status,error){
                        console.log("code", request.status);
                        console.log("message", request.responseText);
                        console.log("error", error);

                        return reject();
                    }
                });
            });
        }
        , change_image_flag(image_name, event){
            const _this = this;

            var file = event.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            switch (image_name){
                case "listimg" : this.image_flag.listimg_f = true;
                    reader.onload = function(e){
                        _this.now_content.listimg = e.target.result;
                    }
                    break;
                case "mainimg" : this.image_flag.mainimg_f = true;
                    reader.onload = function(e){
                        _this.now_content.mainimg = e.target.result;
                    }
                    break;
                case "mainimg2" : this.image_flag.mainimg2_f = true;
                    reader.onload = function(e){
                        _this.now_content.mainimg2 = e.target.result;
                    }
                    break;
                case "snsimg" : this.image_flag.snsimg_f = true;
                    reader.onload = function(e){
                        _this.now_content.snsimg = e.target.result;
                    }
                    break;
                case "prizeimg" : this.image_flag.prizeimg_f = true;
                    reader.onload = function(e){
                        _this.now_content.prizeimg = e.target.result;
                    }
                    break;
            }
        }
        , form_validate() {
            const _this = this;
            _this.check_ok = true;

            $(".must").each(function () {
                if ($(this).val().trim() == "") {
                    _this.check_ok = false;
                    alert("필수사항을 입력해주세요.");
                    $(this).focus();

                    return false;
                }
            });

            if (_this.check_ok) {
                if (!_this.now_content.listimg || _this.now_content.listimg == "") {
                    _this.check_ok = false;
                    alert("리스트 이미지를 등록해주세요.");
                    $("#listimg").focus();
                } else if (!_this.now_content.mainimg || _this.now_content.mainimg == "") {
                    _this.check_ok = false;
                    alert("메인 이미지를 등록해주세요.");
                    $("#mainimg").focus();
                } else if (!_this.now_content.mainimg2 || _this.now_content.mainimg2 == "") {
                    _this.check_ok = false;
                    alert("모바일웹메인 이미지를 등록해주세요.");
                    $("#mainimg2").focus();
                } else if (!_this.now_content.snsimg || _this.now_content.snsimg == "") {
                    _this.check_ok = false;
                    alert("카카오톡 이미지를 등록해주세요.");
                    $("#snsimg").focus();
                } else if (!_this.now_content.prizeimg || _this.now_content.prizeimg == "") {
                    _this.check_ok = false;
                    alert("당첨 이미지를 등록해주세요.");
                    $("#prizeimg").focus();
                }
            }
        }
        , change_price(price_name){
            if(price_name == "option6"){
                this.now_content.option6 = this.now_content.option6.toString().replace(/\D/g, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            }else{
                this.now_content.option7 = this.now_content.option7.toString().replace(/\D/g, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            }
        }
        , change_textarea(textarea_name){
            if(textarea_name == 'sns'){
                this.snstext_length = this.now_content.snstext.length;
            }else if(textarea_name == 'push'){
                this.pushtext_length = this.now_content.pushText.length;
            }
        }
        , go_reset(){
            if(confirm("당첨자는 즉시 삭제됩니다. 초기화 하시겠습니까?")){
                let query_param = new URLSearchParams(window.location.search);

                this.open_day = query_param.get("open_date").split(" ")[0];
                this.open_time = query_param.get("open_date").split(" ")[1].substr(0,5);
                this.end_day = query_param.get("end_date").split(" ")[0];
                this.end_time = query_param.get("end_date").split(" ")[1].substr(0,5);

                this.now_content.option4 = this.now_content.itemea;

                callApiHttps("DELETE", "/event/joobjoob", {"evt_code" : this.evt_code}, function (data) {
                    alert("당첨자가 삭제되었습니다.");
                });
            }
        }
        , check_resetable(){
            let query_param = new URLSearchParams(window.location.search);

            let open_date = new Date(query_param.get("open_date"));
            let now = new Date;

            if( now < open_date){
                return true;
            }else{
                return false;
            }
        }
        , change_itemea(){
            this.itemea_count = this.now_content.itemea;
        }
    }
    , watch:{
        itemea_count(){
            const _this = this;

            const arrDayMin = ["일","월","화","수","목","금","토"];
            const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];

            this.$nextTick(function() {
                $(".prizedate").datepicker({
                    dateFormat: "yy-mm-dd",
                    prevText: '이전달', nextText: '다음달', yearSuffix: '년',
                    dayNamesMin: arrDayMin,
                    monthNames: arrMonth,
                    showMonthAfterYear: true
                });

                $(".prizetime").timepicker({
                    timeFormat: "HH:mm"
                    , dropdown: true
                    , scrollbar: true
                    , dynamic: false
                    , interval: 1
                });

                if(this.content.prize){
                    let count = 1;
                    this.content.prize.forEach(function (){
                        $("#prizedate" + count).val(_this.content.prize[count-1].prizedate);
                        $("#prizetime" + count).val(_this.content.prize[count-1].prizetime);
                        count++;
                    });
                }
            });
        }
        , content() {
            this.itemea_count = this.content.itemea;

            this.now_content = this.content;

            this.open_day = this.now_content.open_date.split(" ")[0];
            this.open_time = this.now_content.open_date.split(" ")[1].substring(0, 5);

            this.end_day = this.now_content.end_date.split(" ")[0];
            this.end_time = this.now_content.end_date.split(" ")[1].substring(0, 5);

            this.snstext_length = this.post_snstext_length;
            this.pushtext_length = this.post_pushtext_length;
        }
    }
});
