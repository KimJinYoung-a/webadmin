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
                        <tr>
                            <th>APP 이벤트 코드</th>
                            <td>
                                <input type="text" name="evt_code" v-model="evt_code" readonly style="background: lightgrey;" class="must"/>
                            </td>
                        </tr>
                        <tr>
                            <th>이벤트 기간</th>
                            <td>
                                <input type="text" id="open_day" v-model="open_day" style="float: left; width: 90px;" class="must"/>
                                <input type="text" id="open_time" v-model="open_time" required size="8" style="float: left;" class="must"/>
                                <input type="text" name="open_date" v-model="now_content.open_date" class="must" style="display: none"/>
                                <p style="float: left;">&nbsp ~ &nbsp</p>
                                <input type="text" id="end_day" v-model="end_day" style="float: left; width: 90px;" class="must"/>
                                <input type="text" id="end_time" v-model="end_time" required size="8" style="float: left;" class="must"/>
                                <input type="text" name="end_date" v-model="now_content.end_date" class="must" style="display: none"/>
                            </td>
                        </tr>
                        <tr>
                            <th>서브카피</th>
                            <td>
                                <textarea v-model="now_content.subcopy" @input="change_textarea" name="subcopy" class="must" rows="5" style="width: 80%"/>
                                {{subcopy_length}} / 200
                            </td>
                        </tr>
                        <tr>
                            <th>마일리지 지급명</th>
                            <td>
                                <input type="text" name="mileage_name" v-model="now_content.mileage_name" class="must" style="width: 80%"/>
                            </td>
                        </tr>
                        <tr>
                            <th>마일리지 사용기간 (까지)</th>
                            <td>
                                <input type="text" id="mileage_expire_day" v-model="mileage_expire_day" style="float: left; width: 90px;" class="must"/>
                                <input type="text" id="mileage_expire_time" v-model="mileage_expire_time" required size="8" readonly style="float: left; background: lightgrey;" class="must"/>
                                <input type="text" name="mileage_expire_date" v-model="now_content.mileage_expire_date" style="display: none"/>
                            </td>
                        </tr>
                        <tr>
                            <th>상단 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.top_img" class="thumbnail-area">
                                    <img id="showTopImage" :src="now_content.top_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('top_img', $event)" name="top_img" id="top_img"/>
                            </td>
                        </tr>   
                        <tr>
                            <th>출석체크 버튼 위 구름문구 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.cloud_img" class="thumbnail-area">
                                    <img id="showCloudImage" :src="now_content.cloud_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('cloud_img', $event)" name="cloud_img" id="cloud_img"/>
                            </td>
                        </tr>
                        <tr>
                            <th>동전 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.coin_img" class="thumbnail-area">
                                    <img id="showCoinImage" :src="now_content.coin_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('coin_img', $event)" name="coin_img" id="coin_img"/>
                            </td>
                        </tr>
                        <tr>
                            <th>체크 이미지</th>
                            <td>
                                <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                                <p style="color:red;">- 최대 크기 : 1024 KB</p>
                                <div v-show="now_content.check_img" class="thumbnail-area">
                                    <img id="showCheckImage" :src="now_content.check_img" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input type="file" @change="change_image_flag('check_img', $event)" name="check_img" id="check_img"/>
                            </td>
                        </tr>
                        
                        <tr>
                            <th>색상</th>
                            <td>
                                <p style="color: red;">- 숫자만 입력하세요</p>
                                <p>출석체크 버튼 배경 : #<input v-model="now_content.attendance_btn_color" type="text" name="attendance_btn_color" size="8" /></p>
                                <p>출석체크 버튼 폰트 : #<input v-model="now_content.attendance_btn_font_color" type="text" name="attendance_btn_font_color" size="8" /></p>
                                <p>마일리지 배경 : #<input v-model="now_content.mileage_color" type="text" name="mileage_color" size="8" /></p>
                                <p>마일리지 폰트 : #<input v-model="now_content.mileage_font_color" type="text" name="mileage_font_color" size="8" /></p>
                                <p>마일리지 리스트 배경 : #<input v-model="now_content.mileage_list_color" type="text" name="mileage_list_color" size="8" /></p>
                                <p>마일리지 리스트 폰트 - 날짜 : #<input v-model="now_content.mileage_list_font_day_color" type="text" name="mileage_list_font_day_color" size="8" /></p>
                                <p>마일리지 리스트 폰트 - 포인트 : #<input v-model="now_content.mileage_list_font_point_color" type="text" name="mileage_list_font_point_color" size="8" /></p>
                                <p>마일리지 투데이 보더 : #<input v-model="now_content.today_border_color" type="text" name="today_border_color" size="8" /></p>                                
                            </td>
                        </tr>
                    </tbody>                    
                </table>
            </form>
            <div style="float: right; padding-top: 10px;">
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
                open_date : ""
                , end_date : ""
                , subcopy : ""
                , mileage_name : ""
                , mileage_expire_date : ""
                , top_img : ""
                , cloud_img : ""
                , coin_img : ""
                , check_img : ""
                , attendance_btn_color : ""
                , attendance_btn_font_color : ""
                , mileage_color : ""
                , mileage_font_color : ""
                , mileage_list_color : ""
                , mileage_list_font_day_color : ""
                , mileage_list_font_point_color : ""
                , today_border_color : ""
            }
            , subcopy_length : 0
            , open_day : ""
            , open_time : ""
            , end_day : ""
            , end_time : ""
            , mileage_expire_day : ""
            , mileage_expire_time : "23:59"
            , image_flag : {
                top_img_f : false
                , cloud_img_f : false
                , coin_img_f : false
                , check_img_f : false
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

        $("#mileage_expire_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.mileage_expire_day = document.getElementById("mileage_expire_day").value;
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
    }
    , methods: {
        go_save(){
            const _this = this;
            this.form_validate();
            if(this.check_ok){
                return new Promise(function (resolve, reject) {
                    _this.now_content.open_date = _this.open_day + " " + _this.open_time;
                    _this.now_content.end_date = _this.end_day + " " + _this.end_time;
                    _this.now_content.mileage_expire_date = _this.mileage_expire_day + " " + _this.mileage_expire_time;

                    return resolve();
                }).then(function(){
                    let form_data = $("#my_form").serialize();

                    let file1 = document.getElementById("top_img").files[0];
                    let file2 = document.getElementById("cloud_img").files[0];
                    let file3 = document.getElementById("coin_img").files[0];
                    let file4 = document.getElementById("check_img").files[0];

                    _this.save_image(file1, file2, file3, file4).then(function () {
                        form_data += "&top_img=" + _this.now_content.top_img;
                        form_data += "&cloud_img=" + _this.now_content.cloud_img;
                        form_data += "&coin_img=" + _this.now_content.coin_img;
                        form_data += "&check_img=" + _this.now_content.check_img;

                        //console.log("form_data", form_data);
                        if (_this.is_write > 0) {
                            callApiHttps("POST", "/event/update-everyday-mileage", form_data, function (data) {
                                alert("수정되었습니다.");
                                self.close();
                            });
                        } else {
                            callApiHttps("POST", "/event/everyday-mileage", form_data, function (data) {
                                alert("등록되었습니다.");
                                self.close();
                            });
                        }
                    });
                });
            }
        }
        , save_image(file1, file2, file3, file4){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                imgData.append("folderName", "everyday_mileage");
                if(_this.image_flag.top_img_f){
                    imgData.append('imgFile1', file1);
                }
                if(_this.image_flag.cloud_img_f){
                    imgData.append('imgFile2', file2);
                }
                if(_this.image_flag.coin_img_f){
                    imgData.append('imgFile3', file3);
                }
                if(_this.image_flag.check_img_f){
                    imgData.append('imgFile4', file4);
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/event_admin/etc_event_admin_multi_reg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);

                        _this.now_content.top_img = response.imgurl1 ? response.imgurl1 : _this.now_content.top_img;
                        _this.now_content.cloud_img = response.imgurl2 ? response.imgurl2 : _this.now_content.cloud_img;
                        _this.now_content.coin_img = response.imgurl3 ? response.imgurl3 : _this.now_content.coin_img;
                        _this.now_content.check_img = response.imgurl4 ? response.imgurl4 : _this.now_content.check_img;

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

            let file = event.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            switch (image_name){
                case "top_img" : this.image_flag.top_img_f = true;
                    reader.onload = function(e){
                        _this.now_content.top_img = e.target.result;
                    }
                    break;
                case "cloud_img" : this.image_flag.cloud_img_f = true;
                    reader.onload = function(e){
                        _this.now_content.cloud_img = e.target.result;
                    }
                    break;
                case "coin_img" : this.image_flag.coin_img_f = true;
                    reader.onload = function(e){
                        _this.now_content.coin_img = e.target.result;
                    }
                    break;
                case "check_img" : this.image_flag.check_img_f = true;
                    reader.onload = function(e){
                        _this.now_content.check_img = e.target.result;
                    }
                    break;
            }
        }
        , form_validate() {
            const _this = this;
            _this.check_ok = true;

            $("#my_form .must").each(function(){
                if($(this).val().trim() == ""){
                    _this.check_ok = false;
                    let th_name = $(this).parent().parent().find("th")[0].innerText;
                    alert("필수항목 " + th_name + "를 입력하지 않으셨습니다.");
                    $(this).focus();

                    return false;
                }
            });

            if (_this.check_ok) {
                if (!_this.now_content.top_img || _this.now_content.top_img == "") {
                    _this.check_ok = false;
                    alert("상단 이미지를 등록해주세요.");
                    $("#top_img").focus();
                }else if (!_this.now_content.cloud_img || _this.now_content.cloud_img == "") {
                    _this.check_ok = false;
                    alert("구름문구 이미지를 등록해주세요.");
                    $("#cloud_img").focus();
                }else if (!_this.now_content.coin_img || _this.now_content.coin_img == "") {
                    _this.check_ok = false;
                    alert("동전 이미지를 등록해주세요.");
                    $("#coin_img").focus();
                }else if (!_this.now_content.check_img || _this.now_content.check_img == "") {
                    _this.check_ok = false;
                    alert("출석체크 완료 이미지를 등록해주세요.");
                    $("#check_img").focus();
                }
            }
        }
        , change_textarea(textarea_name){
            this.subcopy_length = this.now_content.subcopy.length;
        }
    }
    , watch:{
        content() {
            this.now_content = this.content;

            this.open_day = this.now_content.open_date.split(" ")[0];
            this.open_time = this.now_content.open_date.split(" ")[1].substring(0, 5);

            this.end_day = this.now_content.end_date.split(" ")[0];
            this.end_time = this.now_content.end_date.split(" ")[1].substring(0, 5);

            this.mileage_expire_day = this.now_content.mileage_expire_date.split(" ")[0];
            this.mileage_expire_time = this.now_content.mileage_expire_date.split(" ")[1].substring(0, 5);

            this.subcopy_length = this.now_content.subcopy.length;
        }
    }
});
