Vue.component("MIKKI", {
    template : `
        <form id="mikki_form">
            <table class="table table-write table-dark" style="margin-top: 15px;">
                <input type="text" :value="evt_code" name="evt_code" style="display: none"/>
                <input type="text" :value="schedule_idx" name="schedule_idx" style="display: none"/>
                
                <input :value="origin_startDate" type="text" name="origin_startDate" style="display: none"/>
                <input :value="origin_endDate" type="text" name="origin_endDate" style="display: none"/>
                <tbody>
                    <tr>
                        <th>기간</th>
                        <td colspan="3">
                            <input type="text" id="mikki_start_day" v-model="start_day" style="float: left; width: 90px;" class="must" autocomplete="false"/>
                            <input type="text" id="mikki_start_time" v-model="start_time" required size="8" style="float: left;" class="must" autocomplete="false"/>
                            <input type="text" name="startDate" v-model="now_mikki.startDate" class="must" style="display: none"/>
                            <p style="float: left;">&nbsp ~ &nbsp</p>
                            <input type="text" id="mikki_end_day" v-model="end_day" style="float: left; width: 90px;" class="must" autocomplete="false"/>
                            <input type="text" id="mikki_end_time" v-model="end_time" required size="8" style="float: left;" class="must" autocomplete="false"/>
                            <input type="text" name="endDate" v-model="now_mikki.endDate" class="must" style="display: none"/>                  
                        </td>
                    </tr>
                    <tr>
                        <th>상품 코드</th>
                        <td>
                            <input type="text" v-model="now_mikki.itemid" name="itemid" class="must"/>                                
                        </td>
                        
                        <th>상품명</th>
                        <td>
                            <input v-model="now_mikki.itemName" type="text" name="itemName" class="must"/>                                
                        </td>
                    </tr>
                    <tr>
                        <th>PC 상품 이미지</th>
                        <td colspan="3">
                            <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                            <p style="color:red;">- 최대 크기 : 1024 KB</p>
                            <div v-show="now_mikki.itemImage" class="thumbnail-area">
                                <img id="showPrizeImage" :src="now_mikki.itemImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                            </div>
                            <input type="file" @change="change_image_flag($event, 'item')" id="mikkiImage"/>               
                        </td>
                    </tr>
                    <tr>
                        <th>M 상품 이미지</th>
                        <td colspan="3">
                            <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                            <p style="color:red;">- 최대 크기 : 1024 KB</p>
                            <div v-show="now_mikki.moItemImage" class="thumbnail-area">
                                <img :src="now_mikki.moItemImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                            </div>
                            <input type="file" @change="change_image_flag($event, 'm_item')" id="moMikkiImage"/>               
                        </td>
                    </tr>
                    <tr>
                        <th>상품 개수</th>
                        <td>
                            <input v-model="now_mikki.itemCnt" type="number" name="itemCnt" class="must"/>                                
                        </td>
                    </tr>
                    <tr>
                        <th>상품 판매가</th>
                        <td>
                            <input v-model="now_mikki.orgPrice" @input="change_price('orgPrice')" type="text" name="orgPrice" class="must"/>                                
                        </td>
                        
                        <th>상품 할인가</th>
                        <td>
                            <input v-model="now_mikki.sellCash" @input="change_price('sellCash')" type="text" name="sellCash" class="must"/>                                
                        </td>
                    </tr>
                    <tr>
                        <th>할인율</th>
                        <td>
                            <input v-model="now_mikki.saleValue" type="text" name="saleValue" />                                
                        </td>
                        
                        <th>할인구분</th>
                        <td>
                            <select v-model="now_mikki.saleType" name="saleType">
                                <option value="1">비율</option>
                                <option value="2">금액</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <th>PC 매진 이미지</th>
                        <td colspan="3">
                            <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                            <p style="color:red;">- 최대 크기 : 1024 KB</p>
                            <div v-show="now_mikki.soldoutImage" class="thumbnail-area">
                                <img :src="now_mikki.soldoutImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                            </div>
                            <input type="file" @change="change_image_flag($event, 'soldout')" id="soldoutImage"/>               
                        </td>
                    </tr>
                    <tr>
                        <th>M 매진 이미지</th>
                        <td colspan="3">
                            <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                            <p style="color:red;">- 최대 크기 : 1024 KB</p>
                            <div v-show="now_mikki.moSoldoutImage" class="thumbnail-area">
                                <img :src="now_mikki.moSoldoutImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                            </div>
                            <input type="file" @change="change_image_flag($event, 'm_soldout')" id="moSoldoutImage"/>               
                        </td>
                    </tr>
                    <tr>
                        <th>PC 티저 이미지</th>
                        <td colspan="3">
                            <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                            <p style="color:red;">- 최대 크기 : 1024 KB</p>
                            <div v-show="now_mikki.tzImage" class="thumbnail-area">
                                <img :src="now_mikki.tzImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                            </div>
                            <input type="file" @change="change_image_flag($event, 'tz')" id="tzImage"/>               
                        </td>
                    </tr>
                    <tr>
                        <th>M 티저 이미지</th>
                        <td colspan="3">
                            <p style="color:red;">- jpg, gif, png 포맷의 파일만 등록 가능합니다.</p>
                            <p style="color:red;">- 최대 크기 : 1024 KB</p>
                            <div v-show="now_mikki.moTzImage" class="thumbnail-area">
                                <img :src="now_mikki.moTzImage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                            </div>
                            <input type="file" @change="change_image_flag($event, 'm_tz')" id="moTzImage"/>               
                        </td>
                    </tr>
                </tbody>
            </table>
            
            <div style="margin: 15px 0 0 450px;">
                <button @click="go_save_mikki" type="button" class="button dark">저장</button>
                <button @click="$emit('close')" type="button" class="button secondary">취소</button>
            </div>
        </form>
    `
    , props : {
        content_mikki : {
            startDate : {type:String, default:""}
            , endDate : {type:String, default:""}
            , itemid : {type:String, default:""}
            , itemName : {type:String, default:""}
            , itemImage : {type:String, default:""}
            , itemCnt : {type:Number, default:0}
            , orgPrice : {type:Number, default:0}
            , sellCash : {type:Number, default:0}
            , saleValue : {type:Number, default:0}
            , saleType : {type:Number, default:1}
            , soldoutImage : {type:String, default:""}
            , tzImage : {type:String, default:""}
            , moItemImage : {type:String, default:""}
            , moSoldoutImage : {type:String, default:""}
            , moTzImage : {type:String, default:""}
        }
        , mikki_is_write : {type:Boolean, default:false}
        , evt_code : {type:String, default:""}
        , schedule_idx : {type:Number, default:1}
        , mikki_list : {type:Array, default:[]}
        , schedule_start : {type:String, default:""}
        , schedule_end : {type:String, default:""}
    }
    , data(){
        return{
            now_mikki : {
                startDate : ""
                , endDate : ""
                , itemid : ""
                , itemName:""
                , itemImage: ""
                , itemCnt : 0
                , orgPrice : 0
                , sellCash : 0
                , saleValue : 0
                , saleType : 1
                , soldoutImage : ""
                , tzImage : ""
                , moItemImage : ""
                , moSoldoutImage : ""
                , moTzImage : ""
            }
            , image_flag : {
                item : false
                , soldout : false
                , tz : false
                , m_item : false
                , m_soldout : false
                , m_tz : false
            }
            , origin_startDate : ""
            , origin_endDate : ''

            , start_day : ""
            , start_time : ""
            , end_day : ""
            , end_time : ""
        }
    }
    , mounted(){
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#mikki_start_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#mikki_end_day").datepicker('setDate', min_date);
                $("#mikki_end_day").datepicker('option', "minDate", min_date);

                _this.start_day = document.getElementById("mikki_start_day").value;
            }
        });
        $("#mikki_end_day").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.end_day = document.getElementById("mikki_end_day").value;
            }
        });
        $("#mikki_start_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.start_time = document.getElementById("mikki_start_time").value;
            }
        });
        $("#mikki_end_time").timepicker({
            timeFormat: "HH:mm"
            , dropdown: true
            , scrollbar: true
            , dynamic: false
            , interval: 1
            , change : function (time){
                _this.end_time = document.getElementById("mikki_end_time").value;
            }
        });
    }
    , methods : {
        change_image_flag(event, image_name){
            const _this = this;
            const file = event.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            reader.onload = function(e){
                switch (image_name){
                    case "item" : _this.now_mikki.itemImage = e.target.result;
                        _this.image_flag.item = true;
                        break;
                    case "soldout" : _this.now_mikki.soldoutImage = e.target.result;
                        _this.image_flag.soldout = true;
                        break;
                    case "tz" : _this.now_mikki.tzImage = e.target.result;
                        _this.image_flag.tz = true;
                        break;
                    case "m_item" : _this.now_mikki.moItemImage = e.target.result;
                        _this.image_flag.m_item = true;
                        break;
                    case "m_soldout" : _this.now_mikki.moSoldoutImage = e.target.result;
                        _this.image_flag.m_soldout = true;
                        break;
                    case "m_tz" : _this.now_mikki.moTzImage = e.target.result;
                        _this.image_flag.m_tz = true;
                        break;
                }

            }
        }
        , go_save_mikki(){
            const _this = this;

            return new Promise(function (resolve, reject) {
                if(!_this.check_mikki_date()){
                    alert("스케쥴 기간이 아닙니다.");
                    reject();
                }else{
                    _this.now_mikki.startDate = _this.start_day + " " + _this.start_time;
                    _this.now_mikki.endDate = _this.end_day + " " + _this.end_time;

                    _this.now_mikki.orgPrice = _this.now_mikki.orgPrice.toString().replace(/\D/g, "");
                    _this.now_mikki.sellCash = _this.now_mikki.sellCash.toString().replace(/\D/g, "");

                    resolve();
                }
            }).then(function () {
                let check = _this.check_mikki_list_dup(_this.now_mikki.startDate, _this.now_mikki.endDate);

                if(check) {
                    let form_data = $("#mikki_form").serialize();
                    let file1 = document.getElementById("mikkiImage").files[0];
                    let file2 = document.getElementById("soldoutImage").files[0];
                    let file3 = document.getElementById("tzImage").files[0];
                    let file4 = document.getElementById("moMikkiImage").files[0];
                    let file5 = document.getElementById("moSoldoutImage").files[0];
                    let file6 = document.getElementById("moTzImage").files[0];

                    _this.save_image(file1, file2, file3, file4, file5, file6).then(function () {
                        form_data += "&itemImage=" + _this.now_mikki.itemImage;
                        form_data += "&soldoutImage=" + _this.now_mikki.soldoutImage;
                        form_data += "&tzImage=" + _this.now_mikki.tzImage;
                        form_data += "&moItemImage=" + _this.now_mikki.moItemImage;
                        form_data += "&moSoldoutImage=" + _this.now_mikki.moSoldoutImage;
                        form_data += "&moTzImage=" + _this.now_mikki.moTzImage;
                        //console.log("form_data", form_data);

                        if (_this.mikki_is_write) {
                            callApiHttps("PUT", "/event/timedeal-mikki", form_data, function (data) {
                                alert("수정되었습니다.");
                                location.reload();
                            });
                        } else {
                            callApiHttps("POST", "/event/timedeal-mikki", form_data, function (data) {
                                alert("저장되었습니다.");
                                location.reload();
                            });
                        }
                    });
                }else{
                    alert("미끼상품 기간이 중복됩니다.");
                }
            });
        }
        , save_image(file1, file2, file3, file4, file5, file6){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();

                if(_this.image_flag.item){
                    imgData.append('imgFile1', file1);
                }
                if(_this.image_flag.soldout){
                    imgData.append('imgFile2', file2);
                }
                if(_this.image_flag.tz){
                    imgData.append('imgFile3', file3);
                }
                if(_this.image_flag.m_item){
                    imgData.append('imgFile4', file4);
                }
                if(_this.image_flag.m_soldout){
                    imgData.append('imgFile5', file5);
                }
                if(_this.image_flag.m_tz){
                    imgData.append('imgFile6', file6);
                }

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/event_admin/timedeal_admin_imgreg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        //console.log(data);
                        const response = JSON.parse(data);

                        _this.now_mikki.itemImage = response.imgurl1 ? response.imgurl1 : _this.now_mikki.itemImage;
                        _this.now_mikki.soldoutImage = response.imgurl2 ? response.imgurl2 : _this.now_mikki.soldoutImage;
                        _this.now_mikki.tzImage = response.imgurl3 ? response.imgurl3 : _this.now_mikki.tzImage;
                        _this.now_mikki.moItemImage = response.imgurl4 ? response.imgurl4 : _this.now_mikki.moItemImage;
                        _this.now_mikki.moSoldoutImage = response.imgurl5 ? response.imgurl5 : _this.now_mikki.moSoldoutImage;
                        _this.now_mikki.moTzImage = response.imgurl6 ? response.imgurl6 : _this.now_mikki.moTzImage;

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
        , check_mikki_list_dup(startDate, endDate){
            let startDateConvert = new Date(startDate);
            let endDateConvert = new Date(endDate);

            let check_ok = true;
            this.mikki_list.forEach(function (item){
                let mikkiStartDate = new Date(item.startDate);
                let mikkiEndDate = new Date(item.endDate);

                if((startDateConvert == mikkiStartDate && endDateConvert == mikkiEndDate) && ((startDateConvert <= mikkiStartDate && endDateConvert >= mikkiStartDate) || (startDateConvert >= mikkiStartDate && startDateConvert <= mikkiEndDate))){
                    check_ok = false;
                }
            });

            return check_ok;
        }
        , change_price(price_name){
            if(price_name == "orgPrice"){
                this.now_mikki.orgPrice = this.now_mikki.orgPrice.toString().replace(/\D/g, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            }else{
                this.now_mikki.sellCash = this.now_mikki.sellCash.toString().replace(/\D/g, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            }
        }
        , check_mikki_date(){
            let mikki_start = new Date(this.start_day + " " + this.start_time + ":00");
            let mikki_end = new Date(this.end_day + " " + this.end_time + ":00");

            let schedule_start = new Date(this.schedule_start);
            let schedule_end = new Date(this.schedule_end)

            console.log(mikki_start, mikki_end, schedule_start, schedule_end);

            if(mikki_start < schedule_start || mikki_end > schedule_end){
                return false;
            }else{
                return true;
            }
        }
    }
    , watch: {
        mikki_is_write(data){
            if(!data) {
                this.now_mikki = {
                    startDate : ""
                    , endDate : ""
                    , itemid : ""
                    , itemName:""
                    , itemImage: ""
                    , itemCnt : 0
                    , orgPrice : 0
                    , sellCash : 0
                    , saleValue : 0
                    , saleType : 1
                    , soldoutImage : ""
                    , tzImage : ""
                    , moItemImage : ""
                    , moSoldoutImage : ""
                    , moTzImage : ""
                };

                this.start_day = "";
                this.start_time = "";
                this.end_day = "";
                this.end_time = "";
            }
        }
        , content_mikki(content_mikki){
            if(this.mikki_is_write) {
                this.now_mikki = this.content_mikki;

                this.start_day = this.content_mikki.startDate.split(" ")[0];
                this.start_time = this.content_mikki.startDate.split(" ")[1].substring(0, 5);

                this.end_day = this.content_mikki.endDate.split(" ")[0];
                this.end_time = this.content_mikki.endDate.split(" ")[1].substring(0, 5);

                this.origin_startDate = this.content_mikki.startDate;
                this.origin_endDate = this.content_mikki.endDate;
            }
        }
    }
});