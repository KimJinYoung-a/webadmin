var app = new Vue({
    el: "#app",
    store: store,
    template: `
        <div>
            <div style="margin: 5px 30px 30px 30px;">
                <input type="button" @click="change_show_type('list')" value="����Ʈ" class="btn"/>
                <input type="button" @click="change_show_type('contents')" value="������" class="btn"/>
                <div style="float: right">
                    <p style="display: contents; font-size: 15px;"><b>{{nickname.occupation}}</b> {{nickname.nickname}}</p>
                    <input type="button" @click="popup_nickname" value="�г��� ����" class="btn" />
                </div>
            </div>
            
            <div v-if="show_type == 'list'">
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
                            <th>�Ⱓ</th>
                            <td style="text-align:left;display: flex;" colspan="4">
                                <select id="period" class="form-control inline small" style="margin-right: 5px;">
                                    <option value="1">������ ����</option>
                                    <option value="2">������ ����</option>
                                    <option value="3">����� ����</option>
                                </select>
                                
                                <input id="startDate" class="form-control small" size="10" maxlength="10" style="margin-right: 5px;" />
                                 <p>~</p>
                                <input id="endDate" class="form-control small" size="10" maxlength="10" style="margin-left: 5px;" />
                            </td>
                            
                            <th>���̾ƿ�</th>
                            <td style="text-align:left;">
                                <select id="uiNumber" class="form-control inline small">
                                    <option value="0">��ü</option>
                                    <option value="1">����Ʈ��</option>
                                    <option value="2">����</option>
                                    <option value="3">��������</option>
                                    <option value="4">�̺�Ʈ��</option>
                                </select>
                            </td>
                            
                            <td rowspan="3">
                                <button @click="do_search" type="button" class="button dark">�˻�</button> <br/><br/>
                                <button @click="reload" type="button" class="button secondary">�˻�����Reset</button>
                            </td>
                        </tr>
                        <tr>                            
                            <th>������</th>
                            <td style="text-align:left;">
                                <select id="contentsNumber" class="form-control inline small">
                                    <option value="0">��ü</option>
                                    <option value="1">�������ǽ�</option>
                                    <option value="2">Ž����Ȱ</option>
                                    <option value="3">DAY.FILM</option>
                                    <option value="4">THING.����</option>
                                    <option value="5">PLAY.GOODS</option>
                                    <option value="7">WEEKLY WALLPAPER</option>
                                </select>
                            </td>
                            
                            <th>�������</th>
                            <td style="text-align:left;">
                                <select id="stateFlag" class="form-control inline small">
                                    <option value="0" selected="selected">����</option>
                                    <option value="1">��ϴ��</option>
                                    <option value="2">�����ο�û</option>
                                    <option value="3">�ۺ��̿�û</option>
                                    <option value="4">���߿�û</option>
                                    <option value="5">���¿�û</option>
                                    <option value="7">����</option>
                                    <option value="8">����</option>
                                    <option value="9">����</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>Ű���� �˻�</th>
                            <td style="text-align:left;">
                                <select id="searchKey" class="form-control inline small">
                                    <option value="1">��ȣ</option>
                                    <option value="2">��������</option>
                                    <option value="3">�ۼ���</option>
                                </select>
                            </td>
                            
                            <th></th>
                            <td style="text-align:left;">
                            </td>
                          </tr>                      
                    </thead>                
                </table>
                
                <p class="p-table">
                    <span>�˻���� : <strong>{{content_count}}</strong></span>
                    <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                        <option v-for="n in 5" :value="n*10">{{n*10}}���� ����</option>
                    </select>
                    <button id="reg_new_content" @click="popup_content('')" type="button" class="button dark">�ű� ���</button>
                </p>
                
                <p style="margin-bottom: 50px;">
                    <table class="table table-dark">
                        <colgroup>
                            <col style="width:33%"/>
                            <col style="width:33%"/>
                            <col style="width:33%"/>
                        </colgroup>
                        <thead>
                          <tr>
                              <th>������ 1</th>
                              <th>������ 2</th>
                              <th>������ 3</th>
                          </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td v-for="item in opening_list">
                                    <div v-if="item.pidx">
                                        {{item.pidx}} {{item.contentTitleName}} <br/>
                                      <b>{{item.titlename}}</b>
                                      <p>{{item.startdate}} ~ {{item.enddate}} ���� </p>
                                      <input type="button" @click="deleteOpening(item.pidx)" value="����"/>     
                                    </div>                                                           
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </p>
    
                <!-- ����Ʈ ���̺� -->
                <table class="table table-dark">
                    <colgroup>
                        <col style="width:20%"/>
                        <col style="width:40%"/>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                        <col style="width:20%"/>
                    </colgroup>
                    <thead>
                        <tr>
                            <th>�����</th>
                            <th>��</th>
                            <th>��ȸ��</th>
                            <th>���� ����</th>
                            <th>��Ÿ</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr v-for="content in contents" :key="content.content_idx">
                            <td>
                                <p v-if="content.openingflag !== 0">{{content.openingflag}}</p>
                                <img @click="popup_thumbnail" :src="content.listimage" class="img-thumbnail link" style="width:50px;height:50px;" onerror="this.src='/images/loading.gif'" />
                            </td>
                            <td @click="popup_content(content.pidx)" class="link">
                                {{content.pidx}} {{content.contentTitleName}} <br/>
                                <b>{{content.titlename}}</b>
                                <p>{{content.occupation}} {{content.nickname}} {{content.regdate}} ���</p>
                                <p v-if="content.lastupdate">{{content.lastOccupation}} {{content.lastNickname}} {{content.lastupdate}} ��������</p> 
                                <p>{{content.startdate}} ~ {{content.enddate}} ����</p>
                            </td>
                            <td>{{content.viewcount}}</td>
                            <td>{{content.stateflag_name}}</td>
                            <td><input type="button" @click="popup_content(content.pidx)" value="����"/> <input type="button" @click="delete_playlist(content.pidx)" value="����"/></td>
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
                <Modal v-show="show_write_modal" @save="save_list_content" @close="show_write_modal = false" modal_width="830px" header_title="PLAY ������ ���/����">
                    <List-Write slot="body" :pop_content="pop_content" :pop_content_items="pop_content_items" :pop_content_tag="pop_content_tag" 
                        @change_content_tag="change_content_tag"
                        ref="write"/>
                </Modal>
    
                <!-- ����� ��� -->
                <Modal v-show="show_thumbnail_modal" @close="show_thumbnail_modal = false"
                    modal_width="400px" :show_header_yn="false" :show_footer_yn="false"
                    :close_background_click_yn="true"
                >
                    <img width="100%" :src="popup_thumbnail_src" slot="body" />
                </Modal>
            </div>
            
            <div v-else-if="show_type == 'contents'">
                <table class="table table-dark table-search">
                    <colgroup>
                        <col style="width: 30%" />
                        <col style="width: 70%" />
                    </colgroup>
                    <thead class="thead-tenbyten">
                        <tr>
                            <th>�����</th>
                            <td style="text-align:left;display: flex;" colspan="4">
                                <select id="isUsing" class="form-control inline small">
                                    <option value="3">��ü</option>
                                    <option value="1">���</option>
                                    <option value="0">�����</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>���⿩��</th>
                            <td style="text-align:left;">
                                <select id="isView" class="form-control inline small">
                                    <option value="3">��ü</option>
                                    <option value="1">����</option>
                                    <option value="0">�����</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>Ű���� �˻�</th>
                            <td style="text-align:left;">
                                <select id="keywordSearchType" class="form-control inline small">
                                    <option value="number">��ȣ</option>
                                    <option value="contentName">��������</option>
                                    <option value="author">�ۼ���</option>
                                </select>
                                
                                <input type="text" id="keywordSearch" />
                            </td>
                          </tr>         
                          <tr>
                            <td class="td-button align-right">
                                <button @click="reload" type="button" class="button secondary">�˻�����Reset</button>
                                <button @click="do_contents_search" type="button" class="button dark">�˻�</button>
                            </td>
                          </tr>            
                    </thead>
                </table>
                <p class="p-table">
                    <span>�˻���� : <strong>{{content_count}}</strong></span>
                    <i class='fas fa-sync' @click="reload"></i>
                    <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                        <option v-for="n in 5" :value="n*10">{{n*10}}���� ����</option>
                    </select>
                    <button id="reg_new_content" @click="popup_content('')" type="button" class="button dark">�ű� ���</button>
                </p>            
    
                <!-- ����Ʈ ���̺� -->
                <table class="table table-dark">
                    <colgroup>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                        <col style="width:20%"/>
                        <col style="width:10%"/>
                        <col style="width:20%"/>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                        <col style="width:10%"/>
                    </colgroup>
                    <thead>
                        <tr>
                            <th>��ȣ</th>
                            <th>�����</th>
                            <th>������ ��</th>
                            <th>���� �������</th>
                            <th>���� ��������</th>
                            <th>���� ����</th>
                            <th>���� ����</th>
                            <th>���� / ����</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr v-for="item in contents">
                            <td>{{item.cidx}}</td>                        
                            <td>{{item.isusing == 1 ? '���' : '�����'}}</td>
                            <td>{{item.titlename}}</td>
                            <td>
                                <p>{{item.occupation}} {{item.nickname}}</p>
                                <p>{{item.regdate}}</p>
                            </td>
                            <td>
                                <p>{{item.lastOccupation}} {{item.lastNickname}}</p>
                                <p>{{item.lastupdate}}</p>
                            </td>
                            <td>{{item.sortnum}}</td>
                            <td>{{item.isview == 1 ? '������' : '�������'}}</td>
                            <td><input type="button" @click="popup_content(item.cidx)" value="����"/> <input type="button" @click="delete_content(item.cidx)" value="����"/></td>
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
                <Modal v-show="show_write_modal" @save="save_contents" @close="show_write_modal = false" modal_width="830px" header_title="PLAY ������ ���/����">
                    <Content-Write slot="body" :pop_content="pop_content" ref="write"/>
                </Modal>
            </div>
            
            <Modal v-show="show_nickname_modal" @save="save_nickname" @close="show_nickname_modal = false" modal_width="400px" header_title="�г��� ����">
                <Nickname-Write slot="body" :nickname="nickname" ref="nickname"/>
            </Modal>
        </div>
    `,
    data() {
        return {
            show_write_modal: false // ��� ��� ���⿩��
            , show_thumbnail_modal: false // ����� ��� ���⿩��
            , popup_thumbnail_src: "" // ����� ��� �̹��� src
            , show_nickname_modal : false //�г��� ��� ���⿩��
            , check_ok:false // ��ȿ���� �÷��װ�
            , is_saving : false //�񵿱���� �ߺ�ó�� ���� �÷��װ�
        };
    },
    created() {
        this.$store.dispatch("GET_CONTENTS"); // ������ ����Ʈ ��ȸ
        this.$store.dispatch("GET_OPENING_LIST");
        this.$store.dispatch("GET_NICKNAME");
    },
    computed: {
        content_count() {
          // ������ �� ����
          return this.$store.getters.content_count;
        },
        contents() {
          // ������ ����Ʈ
          return this.$store.getters.contents;
        },
        current_page() {
          // ���� ������
          return this.$store.getters.current_page;
        },
        last_page() {
          // ������ ������
          return this.$store.getters.last_page;
        },
        pop_content() {
          // ���� ������
          return this.$store.getters.pop_content;
        }
        , opening_list(){
            return this.$store.getters.opening_list;
        }
        , pop_content_items(){
            return this.$store.getters.pop_content_items;
          }
        , pop_content_tag(){
            return this.$store.getters.pop_content_tag;
        }
        , show_type(){
            return this.$store.getters.show_type;
        }
        , nickname(){
            return this.$store.getters.nickname;
        }
    },
    methods: {
        click_page(page) {
            // ������ Ŭ�� �̺�Ʈ
            this.$store.commit("SET_CURRENT_PAGE", page);
            this.$store.dispatch("GET_CONTENTS");
            window.scrollTo(0, 0);
        },
        do_search() {
            // �˻���ư Ŭ�� �̺�Ʈ
            this.$store.commit("SET_SEARCH_PARAMETER", {
              period: $("#period").val()
              , startdate: $("#startDate").val()
              , enddate: $("#endDate").val()
              , uinumber: $("#uiNumber").val()
              , contentsnumber: $("#contentsNumber").val()
              , stateflag: $("#stateFlag").val()
              , searchkey: $("#searchKey").val()
              , searchstring: $("#searchString").val()
            });
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_CONTENTS");
        }
        , do_contents_search() {
            this.$store.commit("SET_CONTENTS_SEARCH_PARAMETER", {
                isusing: $("#isUsing").val()
                , isview: $("#isView").val()
                , keywordsearchtype: $("#keywordSearchType").val()
                , keywordsearch: $("#keywordSearch").val()
            });
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_CONTENTS");
        }
        ,change_page_size() {
            // �������� ������ ���� �� ����
            this.$store.commit("SET_PAGE_SIZE", document.getElementById("page_size").value);
            this.$store.commit("SET_CURRENT_PAGE", 1);
            this.$store.dispatch("GET_CONTENTS");
        },
        popup_content(pidx) {
            if(pidx != ""){
              this.$store.dispatch("GET_POP_CONTENT", pidx);
            }else{
              this.$store.commit("SET_POP_CONTENT", {});
              this.$store.commit("SET_POP_CONTENT_ITEMS", []);
              this.$store.commit("SET_POP_CONTENT_TAG", []);
            }
            this.show_write_modal = true;
        },
        popup_thumbnail(e) {
            // popup ����� ���
            this.popup_thumbnail_src = e.target.src;
            this.show_thumbnail_modal = true;
        }
        , popup_nickname(){
            this.show_nickname_modal = true;
        }
        , save_list_content() {
            if(this.is_saving){
                return false;
            }

            this.validate_content_data();
            if(this.check_ok){
                const _this = this;
                _this.is_saving = true;

                _this.save_image().then(function (data){
                    // ������ ����
                    const form_data = new FormData(document.play_content);
                    const api_data = {};
                    form_data.forEach((value, key) => {
                        api_data[key] = value;
                    });

                    // let apiType = "Put"; //����
                    let apiType = "Post"; //����
                    let url = "/mobileSite/play/update/list-content";
                    if(!$("input[name=pidx]").val()){
                        apiType="Post"; //���
                        url = "/mobileSite/play/list-content";
                    }
                    callApiHttps(apiType, url, api_data, function (data) {
                        _this.is_saving = false;

                        alert("���� �Ǿ����ϴ�.");
                        _this.$store.dispatch("GET_CONTENTS");
                        _this.$store.dispatch("GET_OPENING_LIST");
                        _this.show_write_modal = false;
                    }, function (xhr){
                        _this.is_saving = false;
                        console.log("ajax error", xhr)
                    });
                });
            }
        }
        , save_contents(){
            if(this.is_saving){
                return false;
            }

            this.validate_content_data();
            if(this.check_ok){
                const _this = this;

                _this.save_image().then(function (data){
                    // ������ ����
                    const form_data = new FormData(document.play_content);
                    const api_data = {};
                    form_data.forEach((value, key) => {
                        api_data[key] = value;
                    });

                    let apiType = "Put"; //����
                    if(!$("input[name=cidx]").val()){
                        apiType="Post"; //���
                    }
                    callApiHttps(apiType, "/mobileSite/play/contents-content", api_data, function (data) {
                        _this.is_saving = false;

                        alert("���� �Ǿ����ϴ�.");
                        _this.$store.dispatch("GET_CONTENTS");
                        _this.show_write_modal = false;
                    }, function (xhr){
                        _this.is_saving = false;
                        console.log("ajax error", xhr)
                    });
                });
            }
        }
        , save_image(){
            const _this = this;
            return new Promise(function (resolve, reject) {
                console.log("this.show_type", _this.show_type);
                let imgChangeF = "";
                let file, filePath;
                if(_this.show_type == "list"){
                    imgChangeF = $("input[name=listimageChangeF]").val();
                    file = document.getElementById("addListimage").files[0];
                    filePath = "list_listimage";
                }else if(_this.show_type == "contents"){
                    imgChangeF = $("input[name=mainimageChangeF]").val();
                    file = document.getElementById("addMainimage").files[0];
                    filePath = "contents_mainimage";
                }

                if(imgChangeF == "Y"){
                    //����Ʈ �̹��� ����
                    const imgData = new FormData();
                    imgData.append('sfImg', file);
                    imgData.append('sName', filePath);

                    let api_url;
                    if (location.hostname.startsWith('webadmin')) {
                        api_url = '//upload.10x10.co.kr';
                    } else {
                        api_url = '//testupload.10x10.co.kr';
                    }
                    $.ajax({
                        url: api_url + "/linkweb/play/play_admin_imgreg_json.asp"
                        , type: "POST"
                        , processData: false
                        , contentType: false
                        , data: imgData
                        , crossDomain: true
                        , success: function (data) {
                            const response = JSON.parse(data);

                            if (response.response === 'ok') {
                                if(_this.show_type == "list"){
                                    app.$refs.write.current_content.listimage = response.imgurl;
                                    console.log(app.$refs.write.current_content.listimage);
                                }else if(_this.show_type == "contents"){
                                    app.$refs.write.current_content.mainimage = response.imgurl;
                                    console.log(app.$refs.write.current_content.mainimage);
                                }

                                return resolve();
                            } else {
                                alert('�̹��� ���� �� ������ �߻��߽��ϴ�. (Err: 001)');
                                return reject();
                            }
                        }
                    });
                }else{
                    return resolve();
                }
            });
        }
        , validate_content_data(content) {
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

            if(this.check_ok && $("select[name=stateflag]").val() == "0"){
                _this.check_ok = false;
                alert("�����Ȳ�� �������ּ���.");
                $("select[name=stateflag]").focus();

                return false;
            }
        },
        reload() {
            window.location.reload(true);
        }
        , deleteOpening(pidx){
            const _this = this;

            if(confirm("�����Ͻðڽ��ϱ�?")){
                const api_data = {
                    pidx : pidx
                    , lastadminid : sessionStorage.getItem("ssBctId")
                };

                callApiHttps("delete", "/mobileSite/play/openinglist", api_data, function (data) {
                    alert("���� �ƽ��ϴ�.");
                    _this.$store.dispatch("GET_OPENING_LIST");
                    _this.show_write_modal = false;
                });
            }
        }
        , delete_playlist(pidx) {
            const _this = this;

            if (confirm("�����Ͻðڽ��ϱ�?")) {
                const api_data = {
                    pidx : pidx
                    , lastadminid : sessionStorage.getItem("ssBctId")
                };

                callApiHttps("DELETE", "/mobileSite/play/list", api_data, function (data) {
                    alert("���� �ƽ��ϴ�.");
                    _this.$store.dispatch("GET_CONTENTS");
                });
            }
        }
        , change_show_type(show_type){
            this.$store.commit("SET_SHOW_TYPE", show_type);
            this.$store.dispatch("GET_CONTENTS");
        }
        , delete_content(cidx){
            const _this = this;

            if (confirm("�����Ͻðڽ��ϱ�?")) {
                const api_data = {
                    cidx : cidx
                };

                callApiHttps("DELETE", "/mobileSite/play/contents", api_data, function (data) {
                    alert("���� �ƽ��ϴ�.");
                    _this.$store.dispatch("GET_CONTENTS");
                });
            }
        }
        , save_nickname(){
            const _this = this;
            let api_data = $("#play_nickname").serialize();

            callApiHttps("PUT", "/mobileSite/play/nickname", api_data, function (data) {
                alert("���� �Ǿ����ϴ�.");
                _this.$store.dispatch("GET_NICKNAME");
                _this.show_nickname_modal = false;
            }, function (xhr){console.log("ajax error", xhr)});
        }
        , change_content_tag(data){
            this.$store.commit("SET_POP_CONTENT_TAG", data);
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
