Vue.component('Snack-Write',{
    template: `
        <div style="height: 600px; overflow-y:auto;">
            <p style="color:red;">
                - : ���Ͻ� �ۼ��� ���°� �˴ϴ�. <br/>
                * : �ʼ����Դϴ�. ���Ͻ� ������ �Ұ����մϴ�.
            </p>
            <form id="snack_content" enctype="multipart/form-data">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>                        
                        <tr>
                            <th>����</th>
                            <td>
                                <div v-show="current_content.video_url">
                                    <input @click="go_video(current_content.video_url)" type="button" value="���� ����" style="margin-bottom: 5px; border: dotted;"/>
                                    <input @click="$emit('go_delete_video', current_content.video_idx, current_content.media_id)" type="button" value="���� ����" style="margin-bottom: 5px;"/>
                                </div>
                                <input @change="$emit('change_video_flag')" type="file" name="video" id="video" />
                            </td>
                        </tr>
                        <tr>
                            <th>���� �����</th>
                            <td>
                                <div v-show="current_content.video_thumbnail_url" class="thumbnail-area">
                                    <img id="video_thumbnail_url" :src="current_content.video_thumbnail_url" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image('video_thumbnail_url', $event)" type="file" name="video_thumbnail" id="video_thumbnail" />
                            </td>
                        </tr>
                        <tr>
                            <th>�׸� ���� <p style="display: inline; color: red;">-</p></th>
                            <td>
                                <select v-model="current_content.entry_type" name="entry_type">
                                    <option value="item">��ǰ</option>
                                    <option value="brand">�귣��</option>
                                    <option value="event">�̺�Ʈ/��ȹ��</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>�׸� ID <p style="display: inline; color: red;">-</p></th>
                            <td>
                                <input v-model="current_content.entry_id" type="text" name="entry_id" />
                            </td>
                        </tr>
                        <tr>
                            <th>�׸� �����</th>
                            <td>
                                <div v-show="current_content.entry_thumbnail_url" class="thumbnail-area">
                                    <img id="entry_thumbnail_url" :src="current_content.entry_thumbnail_url" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image('entry_thumbnail_url', $event)" type="file" name="entry_thumbnail" id="entry_thumbnail" />
                            </td>
                        </tr>
                        <tr>
                            <th>�׸�URL <p style="display: inline; color: red;">-</p></th>
                            <td>
                                <input v-model="current_content.entry_url" type="text" name="entry_url" style="width: 80%"/>
                            </td>
                        </tr>
                        <tr>
                            <th>�׸��</th>
                            <td>
                                <input v-model="current_content.entry_name" type="text" name="entry_name" />
                            </td>
                        </tr>
                        <tr>
                            <th>�׸񼳸�</th>
                            <td>
                                <textarea v-model="current_content.entry_desc" rows="4" style="width: 80%" name="entry_desc" ></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>�ۼ��� �̸�</th>
                            <td>
                                <p style="color: red">-�ۼ��� �̸��� �ۼ����� �����ø� �ۼ��� ���� ������ ���� ������ ������ ����˴ϴ�.</p>
                                <input v-model="current_content.writer_name" type="text" name="writer_name" />
                            </td>
                        </tr>
                        <tr>
                            <th>�ۼ��� ����Ÿ��Ʋ</th>
                            <td>
                                <input v-model="current_content.writer_subtitle" type="text" name="writer_subtitle" />
                            </td>
                        </tr>
                        <tr>
                            <th>�ۼ��� �̹���</th>
                            <td>
                                <div v-show="current_content.writer_thumbnail_url" class="thumbnail-area">
                                    <img id="writer_thumbnail_url" :src="current_content.writer_thumbnail_url" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                </div>
                                <input @change="change_image('writer_thumbnail_url', $event)" type="file" name="writer_thumbnail" id="writer_thumbnail" />
                            </td>
                        </tr>
                        <tr>
                            <th>�������� mp4 ���</th>
                            <td>
                                <p style="color: red">-���� ���ε��� �ڵ� ���� �˴ϴ�.</p>
                                <input v-model="current_content.video_url_mp4" type="text" name="video_url_mp4" readonly style="background: lightgrey; width: 80%;"/>
                            </td>
                        </tr>
                        <tr>
                            <th>���� �Ⱓ <p style="display: inline; color: red;">*</p></th>
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

        const arrDayMin = ["��","��","ȭ","��","��","��","��"];
        const arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
        $("#start_dt").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '������', nextText: '������', yearSuffix: '��',
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
            prevText: '������', nextText: '������', yearSuffix: '��',
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