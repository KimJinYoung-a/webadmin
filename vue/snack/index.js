const app = new Vue({
    el: "#app"
    , store: store
    , mixins: [api_mixin]
    , template: `
        <div>
            <!-- �˻� ���̺� -->
            <table class="table table-dark table-search">
                <colgroup>
                    <col style="width:10%"/>
                    <col style="width:30%"/>
                    <col style="width:10%"/>
                    <col style="width:30%"/>
                    <col style="width:20%"/>
                </colgroup>
                <thead class="thead-tenbyten">
                    <tr>
                        <th>����</th>
                        <td style="text-align:left;">
                            <select v-model="search_param.state" class="form-control inline small">
                                <option value="all">��ü</option>
                                <option value="encoding">���ڵ�</option>
                                <option value="writing">�ۼ���</option>    
                                <option value="pre">����</option>
                                <option value="ing">������</option>                                
                                <option value="end">����</option>  
                            </select>
                        </td>
                        
                        <th>�׸� ����</th>
                        <td style="text-align:left;">
                            <select v-model="search_param.search_entry_type" class="form-control inline small">
                                <option value="all">��ü</option>
                                <option value="item">��ǰ</option>
                                <option value="brand">�귣��</option>
                                <option value="event">�̺�Ʈ/��ȹ��</option>
                            </select>
                        </td>
                        
                        <td rowspan="2">
                            <button @click="go_search" type="button" class="button dark">�˻�</button>
                            <button @click="reload" type="button" class="button secondary">�˻�����Reset</button>
                        </td>
                    </tr>
                    <tr>                            
                        <th>Ű���� �Է�</th>
                        <td style="text-align:left;">
                            <select v-model="search_param.search_keyword_option" class="form-control inline small">
                                <option value="entry_id">�׸� ID</option>
                                <option value="entry_name">�׸��</option>
                                <option value="entry_desc">�׸� ����</option>
                                <option value="writer_name">�ۼ��� �̸�</option>
                            </select>
                            <input v-model="search_param.search_keyword_text" type="text" class="form-control inline" style="width: 70%"/>
                        </td>
                        
                        <th>���ڵ� ����</th>
                        <td style="text-align:left;">
                            <select v-model="search_param.search_encoding_state" class="form-control inline small">
                                <option value="A">��ü</option>
                                <option value="Y">����</option>
                                <option value="N">���</option>
                                <option value="E">����</option>
                            </select>
                        </td>
                    </tr>                   
                </thead>                
            </table>
            
            <p class="p-table">
                <span>�˻���� : <strong>{{total_count}}</strong></span>
                <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                    <option v-for="n in 5" :value="n*10">{{n*10}}���� ����</option>
                </select>
                <button @click="popup_sort()" type="button" class="button dark">��������</button>
                <button @click="popup_content('', 'write')" type="button" class="button dark">�ű� ���</button>
            </p>

            <!-- ����Ʈ ���̺� -->
            <table class="table table-dark">
                <colgroup>
                    <col style="width:5%"/>
                    <col style="width:5%"/>
                    <col style="width:20%"/>
                    <col style="width:20%"/>
                    <col style="width:05%"/>
                    <col style="width:5%"/>
                    <col style="width:5%"/>
                    <col style="width:20%"/>
                </colgroup>
                <thead>
                    <tr>
                        <th><input type="checkbox" /></th>
                        <th>idx</th>
                        <th>�ۼ��� ����(�̸�/����)</th>
                        <th>�׸� �����</th>
                        <th>�׸񼳸�</th>
                        <th>���ڵ� ����</th>
                        <th>����</th>
                        <th>����Ⱓ</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in list" @click="popup_content(item.video_idx, 'edit')">
                        <td><input type="checkbox" /></td>
                        <td>{{item.video_idx}}</td>
                        <td>{{item.writer_name}} / {{item.writer_subtitle}}</td>
                        <td><img :src="item.entry_thumbnail_url" style="max-height: 80px;"></td>
                        <td>{{item.entry_desc}}</td>
                        <td>{{item.encoding_state}}</td>
                        <td>{{item.state}}</td>
                        <td>{{item.start_dt}} ~ {{item.end_dt}}</td>
                    </tr>
                </tbody>
            </table>

            <!-- ������ -->
            <Pagination
                @click_page="click_page"
                :current_page="current_page"
                :last_page="last_page"
            ></Pagination>

            <!-- ���/���� ��� -->
            <Modal v-show="show_write_modal" @save="save_content" @close="show_write_modal = false" modal_width="830px" header_title="Snack ���/����">
                <Snack-Write slot="body" ref="write"
                    @change_image_flag="change_image_flag" @change_video_flag="change_video_flag"
                    @go_delete_video="go_delete_video"
                    :write_mode="write_mode" :content="content"
                />
            </Modal>
            
            <Modal v-show="show_sort_modal" @save="save_sort" @close="show_sort_modal = false" modal_width="830px" header_title="Snack ���/����">
                <Snack-Sort slot="body" ref="sort"
                    :content="sort_list"
                />
            </Modal>
            
            <div v-show="is_saving">
                <div class="bg_dim" style="position: fixed;top:0;left:0;right:0;bottom:0;z-index: 9999;background-color: rgba(0,0,0,0.5);"></div>
                <div id="lyLoading" style="z-index: 9999;position: fixed;left: 50%;transform: translateX(-50%);">
                    <img src="http://fiximage.10x10.co.kr/icons/loading16.gif" style="width: 35px;height: 35px;">
                </div>
            </div>            
        </div>
    `
    , data() {
        return {
            show_write_modal : false // ��� ��� ���⿩��
            , show_sort_modal : false //���� ��� ���⿩��
            , check_ok : false // ��ȿ���� �÷��װ�
            , is_saving : false //�񵿱���� �ߺ�ó�� ���� �÷��װ�

            , image_flag : {
                video_thumbnail_flag : false
                , entry_thumbnail_flag : false
                , writer_thumbnail_flag : false
            }
            , video_flag : false
            , write_mode : "wait"
            , search_param : {
                state : "all"
                , search_entry_type : "all"
                , search_keyword_option : "entry_id"
                , search_keyword_text : ""
                , search_encoding_state : "A"
            }
        };
    }
    , created() {
        this.$store.dispatch("GET_LIST", this.search_param);
    }
    , computed: {
        list(){
            return this.$store.getters.list;
        }
        , content(){
            return this.$store.getters.content;
        }
        , total_count(){
            return this.$store.getters.total_count;
        }
        , current_page(){
            return this.$store.getters.current_page;
        }
        , last_page(){
            return this.$store.getters.last_page;
        }
        , sort_list(){
            return this.$store.getters.sort_list;
        }
    }
    , methods: {
        click_page(page) {
            // ������ Ŭ�� �̺�Ʈ
            this.$store.commit("SET_CURRENT_PAGE", page);
            this.$store.dispatch("GET_LIST", this.search_param);
            window.scrollTo(0, 0);
        },
        go_search() {
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_LIST", this.search_param);
        }
        ,change_page_size() {
            // �������� ������ ���� �� ����
            this.$store.commit("SET_PAGE_SIZE", document.getElementById("page_size").value);
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_LIST", this.search_param);
        }
        , popup_content(video_id, mode) {
            if(mode == "edit"){
                this.$store.dispatch("GET_CONTENT", video_id);
            }

            this.write_mode = mode;
            this.show_write_modal = true;
        }
        , save_content() {
            if(this.is_saving){
                return false;
            }

            this.validate_content_data();
            if(this.check_ok){
                const _this = this;
                _this.is_saving = true;

                this.save_image().then(function(data){
                    let form_data = new FormData();
                    let video = $("#video")[0].files[0];
                    if(_this.video_flag){
                        form_data.append("video", video);
                    }else{
                        form_data.append("video_url", app.$refs.write.current_content.video_url);
                    }
                    form_data.append("video_thumbnail_url", app.$refs.write.current_content.video_thumbnail_url);
                    form_data.append("entry_type", app.$refs.write.current_content.entry_type);
                    form_data.append("entry_id", app.$refs.write.current_content.entry_id);
                    form_data.append("entry_thumbnail_url", app.$refs.write.current_content.entry_thumbnail_url);
                    form_data.append("entry_url", app.$refs.write.current_content.entry_url);
                    form_data.append("entry_name", app.$refs.write.current_content.entry_name);
                    form_data.append("entry_desc", app.$refs.write.current_content.entry_desc);
                    form_data.append("writer_name", app.$refs.write.current_content.writer_name);
                    form_data.append("writer_subtitle", app.$refs.write.current_content.writer_subtitle);
                    form_data.append("writer_thumbnail_url", app.$refs.write.current_content.writer_thumbnail_url);
                    form_data.append("start_dt", app.$refs.write.current_content.start_dt);
                    form_data.append("end_dt", app.$refs.write.current_content.end_dt);
                    form_data.append("mediaId", app.$refs.write.current_content.media_id);

                    let api_url, api_type;
                    if (location.hostname.startsWith('webadmin')) {
                        api_url = '//fapi.10x10.co.kr/api/web/v1';
                    }else if(location.hostname.startsWith('localwebadmin')) {
                        api_url = '//localhost:8080/api/web/v1';
                    }else{
                        api_url = '//testfapi.10x10.co.kr/api/web/v1';
                    }

                    if(_this.write_mode == "write"){
                        api_type = "POST";
                    }else if(_this.write_mode == "edit"){
                        api_type = "PUT";
                        form_data.append("video_idx", app.$refs.write.current_content.video_idx);
                        form_data.append("video_url_mp4", app.$refs.write.current_content.video_url_mp4);
                    }

                    $.ajax({
                        type: api_type
                        , url: api_url + "/snack/upload"
                        , data : form_data
                        , processData : false
                        , contentType: false
                        , crossDomain: true
                        , xhrFields: {
                            withCredentials: true
                        }
                        , success: function(data){
                            alert("���� �Ǿ����ϴ�.");
                            _this.$store.dispatch("GET_LIST", _this.search_param);
                            _this.show_write_modal = false;
                        }
                        , error: function(xhr){
                            console.log("ajax error", xhr)
                        }
                        , complete : function(){
                            _this.is_saving = false;
                            _this.write_mode = "wait";
                            _this.video_flag = false;
                        }
                    });
                });
            }
        }
        , save_image(){
            const _this = this;

            return new Promise(function (resolve, reject) {
                const imgData = new FormData();
                if(_this.image_flag.video_thumbnail_flag){
                    imgData.append('photo1', document.getElementById("video_thumbnail").files[0]);
                }
                if(_this.image_flag.entry_thumbnail_flag){
                    imgData.append('photo2', document.getElementById("entry_thumbnail").files[0]);
                }
                if(_this.image_flag.writer_thumbnail_flag){
                    imgData.append('photo3', document.getElementById("writer_thumbnail").files[0]);
                }


                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = '//upload.10x10.co.kr';
                } else {
                    api_url = '//testupload.10x10.co.kr';
                }
                $.ajax({
                    url: api_url + "/linkweb/snack/snack_reg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        const response = JSON.parse(data);
                        console.log(response);

                        if (response.response === 'ok') {
                            if(response.photo1){
                                app.$refs.write.current_content.video_thumbnail_url = response.photo1;
                            }
                            if(response.photo2){
                                app.$refs.write.current_content.entry_thumbnail_url = response.photo2;
                            }
                            if(response.photo3){
                                app.$refs.write.current_content.writer_thumbnail_url = response.photo3;
                            }

                            return resolve();
                        } else {
                            alert('�̹��� ���� �� ������ �߻��߽��ϴ�. (Err: 001)');
                            return reject();
                        }
                    }
                    , complete : function(){
                        _this.image_flag.video_thumbnail_flag = false;
                        _this.image_flag.entry_thumbnail_flag = false;
                        _this.image_flag.writer_thumbnail_flag = false;
                    }
                });
            });
        }
        , validate_content_data() {
            const _this = this;

            this.check_ok = true;
            $(".must").each(function(){
                if($(this).val().trim() == ""){
                    _this.check_ok = false;
                    let th_name = $(this).parent().parent().find("th")[0].innerText;
                    alert("�ʼ��׸� " + th_name + "�� �Է����� �����̽��ϴ�.");
                    $(this).focus();

                    return false;
                }
            });
        }
        , reload() {
            window.location.reload(true);
        }
        , change_image_flag(type){
            switch (type){
                case "video_thumbnail_url" :
                    this.image_flag.video_thumbnail_flag = true;
                    break;
                case "entry_thumbnail_url" :
                    this.image_flag.entry_thumbnail_flag = true;
                    break;
                case "writer_thumbnail_url" :
                    this.image_flag.writer_thumbnail_flag = true;
                    break;
            }
        }
        , change_write_mode(mode){
            this.write_mode = mode;
        }
        , change_video_flag(){
            this.video_flag = true;
        }
        , go_delete_video(video_idx, media_id){
            const _this = this;

            if(this.is_saving){
                return false;
            }

            if(confirm("������ �����Ͻðڽ��ϱ�?? \n������ �����Ͻ� �� �����ϴ�.")){
                this.is_saving = true;

                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = '//fapi.10x10.co.kr/api/web/v1';
                }else if(location.hostname.startsWith('localwebadmin')) {
                    api_url = '//localhost:8080/api/web/v1';
                }else{
                    api_url = '//testfapi.10x10.co.kr/api/web/v1';
                }
                $.ajax({
                    url: api_url + "/snack/upload"
                    , type: "DELETE"
                    , data: {"video_idx" : video_idx, "mediaId" : media_id}
                    , crossDomain: true
                    , success: function (data) {
                        _this.$store.dispatch("GET_CONTENT", video_idx);
                        alert("������ �����Ǿ����ϴ�.");
                    }
                    , complete : function(){
                        _this.is_saving = false;
                    }
                });
            }
        }
        , popup_sort(){
            this.$store.dispatch("GET_SORT_LIST");
            this.show_sort_modal = true;
        }
        , save_sort(){
            const _this = this;

            if(_this.is_saving){
                return false;
            }

            _this.is_saving = true;

            let api_url;
            if (location.hostname.startsWith('webadmin')) {
                api_url = '//fapi.10x10.co.kr/api/web/v1';
            }else if(location.hostname.startsWith('localwebadmin')) {
                api_url = '//localhost:8080/api/web/v1';
            }else{
                api_url = '//testfapi.10x10.co.kr/api/web/v1';
            }
            $.ajax({
                url: api_url + "/snack/sort-list"
                , type: "PUT"
                , data: {"sort_idx" : app.$refs.sort.sorted_arr}
                , crossDomain: true
                , success: function (data) {
                    alert("���������� �Ϸ�Ǿ����ϴ�.");
                    location.reload(); //����ȸ���Ʈ ������ �ٲ� �Ⱥ���
                    //_this.show_sort_modal = false;
                }
                , complete : function(){
                    _this.is_saving = false;
                }
            });
        }
    }
    , mounted() {
        const arrDayMin = ["��","��","ȭ","��","��","��","��"];
        const arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
        $(".thead-tenbyten #startDate").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '������', nextText: '������', yearSuffix: '��',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
        });
        $(".thead-tenbyten #endDate").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '������', nextText: '������', yearSuffix: '��',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
        });
    }
});
