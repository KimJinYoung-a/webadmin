Vue.component('Snack-Write',{
    template: `
        <div style="height: 600px; overflow-y:auto;">
            <p style="color:red;">
                - : 빈값일시 작성중 상태가 됩니다. <br/>
                * : 필수값입니다. 빈값일시 저장이 불가능합니다.
            </p>
            <form id="snack_content" enctype="multipart/form-data">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>                        
                        <tr>
                            <th>영상</th>
                            <td>
                                <div v-show="current_content.video_url">
                                    <input @click="go_video(current_content.video_url)" type="button" value="영상 보기" style="margin-bottom: 5px; border: dotted;"/>
                                    <input @click="$emit('go_delete_video', current_content.video_idx, current_content.media_id)" type="button" value="영상 제거" style="margin-bottom: 5px;"/>
                                </div>
                                <input @change="$emit('change_video_flag')" type="file" name="video" id="video" />
                            </td>
                        </tr>
                        <tr>
                            <th>영상 썸네일</th>
                            <td>
                                <div v-show="current_content.video_thumbnail_url" class="thumbnail-area">
                                    <img id="video_thumbnail_url" :src="current_content.video_thumbnail_url" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image('video_thumbnail_url', $event)" type="file" name="video_thumbnail" id="video_thumbnail" />
                            </td>
                        </tr>
                        <tr>
                            <th>항목 구분 <p style="display: inline; color: red;">-</p></th>
                            <td>
                                <select v-model="current_content.entry_type" name="entry_type">
                                    <option value="item">상품</option>
                                    <option value="brand">브랜드</option>
                                    <option value="event">이벤트/기획전</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>항목 ID <p style="display: inline; color: red;">-</p></th>
                            <td>
                                <input v-model="current_content.entry_id" type="text" name="entry_id" />
                            </td>
                        </tr>
                        <tr>
                            <th>항목 썸네일</th>
                            <td>
                                <div v-show="current_content.entry_thumbnail_url" class="thumbnail-area">
                                    <img id="entry_thumbnail_url" :src="current_content.entry_thumbnail_url" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image('entry_thumbnail_url', $event)" type="file" name="entry_thumbnail" id="entry_thumbnail" />
                            </td>
                        </tr>
                        <tr>
                            <th>항목URL <p style="display: inline; color: red;">-</p></th>
                            <td>
                                <input v-model="current_content.entry_url" type="text" name="entry_url" style="width: 80%"/>
                            </td>
                        </tr>
                        <tr>
                            <th>항목명</th>
                            <td>
                                <input v-model="current_content.entry_name" type="text" name="entry_name" />
                            </td>
                        </tr>
                        <tr>
                            <th>항목설명</th>
                            <td>
                                <textarea v-model="current_content.entry_desc" rows="4" style="width: 80%" name="entry_desc" ></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>작성자 이름</th>
                            <td>
                                <p style="color: red">-작성자 이름을 작성하지 않으시면 작성자 관련 정보는 전부 프로필 정보로 노출됩니다.</p>
                                <input v-model="current_content.writer_name" type="text" name="writer_name" />
                            </td>
                        </tr>
                        <tr>
                            <th>작성자 서브타이틀</th>
                            <td>
                                <input v-model="current_content.writer_subtitle" type="text" name="writer_subtitle" />
                            </td>
                        </tr>
                        <tr>
                            <th>작성자 이미지</th>
                            <td>
                                <div v-show="current_content.writer_thumbnail_url" class="thumbnail-area">
                                    <img id="writer_thumbnail_url" :src="current_content.writer_thumbnail_url" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image('writer_thumbnail_url', $event)" type="file" name="writer_thumbnail" id="writer_thumbnail" />
                            </td>
                        </tr>
                        <tr>
                            <th>아이폰용 mp4 경로</th>
                            <td>
                                <p style="color: red">-영상 업로드후 자동 생성 됩니다.</p>
                                <input v-model="current_content.video_url_mp4" type="text" name="video_url_mp4" readonly style="background: lightgrey; width: 80%;"/>
                            </td>
                        </tr>
                        <tr>
                            <th>영상 기간 <p style="display: inline; color: red;">*</p></th>
                            <td>
                                <input v-model="current_content.start_dt" type="text" name="start_dt" id="start_dt" style="width: 90px;" class="must" /> ~ <input v-model="current_content.end_dt" type="text" name="end_dt" id="end_dt" style="width: 90px;" class="must" />
                            </td>
                        </tr>
                    </tbody>
                </table>                
            </form>
        </div>
    `
    , mounted() {
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#start_dt").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#end_dt").datepicker('setDate', min_date);
                $("#end_dt").datepicker('option', "minDate", min_date);

                _this.current_content.start_dt = document.getElementById("start_dt").value;
            }
        });

        $("#end_dt").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.current_content.end_dt = document.getElementById("end_dt").value;
            }
        });
    }
    , data() {
        return {
            current_content : {
                video_idx : null
                , video_url : null
                , video_thumbnail_url : ""
                , entry_type : "item"
                , entry_id : null
                , entry_thumbnail_url : ""
                , entry_url : ""
                , entry_name : ""
                , entry_desc : ""
                , writer_name : ""
                , writer_subtitle : ""
                , writer_thumbnail_url : ""
                , video_url_mp4 : ""
                , start_dt : ""
                , end_dt : ""
                , media_id: ""
            }
        }
    }
    , props: {
        content : {
            video_idx : { type:String, default: null }
            , video_url : { type:String, default: null }
            , video_thumbnail_url : { type:String, default: null }
            , entry_type : { type:String, default: "item" }
            , entry_id : { type:String, default: null }
            , entry_thumbnail_url : { type:String, default: null }
            , entry_url : { type:String, default: null }
            , entry_name : { type:String, default: null }
            , entry_desc : { type:String, default: null }
            , writer_name : { type:String, default: null }
            , writer_subtitle : { type:String, default: null }
            , writer_thumbnail_url : { type:String, default: null }
            , video_url_mp4 :  { type:String, default: null }
            , start_dt : { type:String, default: null }
            , end_dt : { type:String, default: null }
            , media_id : { type:String, default: null }
        }
        , write_mode : {type:String, default:"wait"}
    }
    , methods : {
        change_image(image_name, event){
            const _this = this;

            let file = event.target.files[0];

            if (!file.type.match("image.*")) {
                alert("only image");
                return false;
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            switch (image_name){
                case "video_thumbnail_url" :
                    _this.$emit('change_image_flag', 'video_thumbnail_url');
                    reader.onload = function(e){
                        _this.current_content.video_thumbnail_url = e.target.result;
                    }
                    break;
                case "entry_thumbnail_url" :
                    _this.$emit('change_image_flag', 'entry_thumbnail_url');
                    reader.onload = function(e){
                        _this.current_content.entry_thumbnail_url = e.target.result;
                    }
                    break;
                case "writer_thumbnail_url" :
                    _this.$emit('change_image_flag', 'writer_thumbnail_url');
                    reader.onload = function(e){
                        _this.current_content.writer_thumbnail_url = e.target.result;
                    }
                    break;
            }
        }
        , go_video(video_url){
            window.open(video_url);
        }
        , init_write_data(){
            this.current_content = {
                video_idx : null
                , video_url : null
                , video_thumbnail_url : ""
                , entry_type : "item"
                , entry_id : null
                , entry_thumbnail_url : ""
                , entry_url : ""
                , entry_name : ""
                , entry_desc : ""
                , writer_name : ""
                , writer_subtitle : ""
                , writer_thumbnail_url : ""
                , video_url_mp4 : ""
                , start_dt : ""
                , end_dt : ""
            }
        }
    }
    , watch:{
        content(content){
            this.init_write_data();
            this.current_content = this.content;
        }
        , write_mode(write_mode){
            console.log(write_mode);
            const _this = this;
            switch (write_mode){
                case "write" : _this.init_write_data(); break;
            }
        }
    }
});