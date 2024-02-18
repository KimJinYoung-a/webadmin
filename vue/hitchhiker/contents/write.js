/*
    ��ġ����Ŀ ������ ���� ��� Body
*/
Vue.component('Hitchhiker-Content-Write',{
    template: `
        <div>
            <form name="hitchhiker_content">
                <input v-if="current_content.content_idx != 0" name="content_idx" type="hidden" v-model="content.content_idx">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>����</th>
                            <td>
                                <select id="write_gubun" name="gubun" class="form-control inline small" v-model="current_content.gubun">
                                    <option value="">����</option>
                                    <option value="1">PC</option>
                                    <option value="2">MOBILE</option>
                                    <option value="3">MOVIE</option>
                                    <option value="4">MOBILE���</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>Ÿ��Ʋ</th>
                            <td>
                                <input id="write_title" name="title" v-model="current_content.title" type="text" class="form-control" placeholder="Ÿ��Ʋ �Է�">
                            </td>
                        </tr>
                        <tr>
                            <th>������</th>
                            <td>
                                <input id="write_start_date" name="start_date" v-model="current_content.start_date" type="text" class="form-control small inline" placeholder="������ �Է�">
                            </td>
                        </tr>
                        <tr>
                            <th>��� ����</th>
                            <td>
                                <div class="form-check">
                                    <input v-model="current_content.use_yn" value="true" name="use_yn" id="use_yn_true" class="form-check-input" type="radio">
                                    <label class="form-check-label" for="use_yn_true">�����</label>
                                    <input v-model="current_content.use_yn" value="false" name="use_yn" id="use_yn_false" class="form-check-input" type="radio">
                                    <label class="form-check-label" for="use_yn_false">������</label>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <th>�����</th>
                            <td>
                                <div @click="delete_thumbnail" class="thumbnail-area">
                                    <img v-show="current_content.thumbnail" :src="current_content.thumbnail" class="thumbnail">
                                    <div class="overlay">����</div>
                                </div>
                                <input id="write_thumbnail" @change="upload_thumbnail" type="file" class="form-control inline middle">
                                <input name="thumbnail" type="hidden" v-model="current_content.thumbnail">
                            </td>
                        </tr>
                        <tr v-if="current_content.gubun == 1 || current_content.gubun == 2">
                            <th>�̹��� ��ũ</th>
                            <td>
                                <ul v-if="current_content.gubun == 1" class="ul-write-hitchhiker">
                                    <li v-for="size in current_content.pc_wallpaper_sizes">
                                    <span class="small">������ : {{size.content_size}}</span>
                                        <input name="size_link_idxs" v-model="size.link_idx" type="hidden">
                                        <input name="size_device_idxs" v-model="size.device_idx" type="hidden">
                                        <span>��ũ : <input name="size_links" v-model="size.link" type="text" class="form-control inline middle"></span>
                                    </li>
                                </ul>
                                <ul v-if="current_content.gubun == 2" class="ul-write-hitchhiker">
                                    <li v-for="size in current_content.mobile_wallpaper_sizes">
                                        <input name="size_link_idxs" v-model="size.link_idx" type="hidden">
                                        <input name="size_device_idxs" v-model="size.device_idx" type="hidden">
                                        <span class="small">��ǥ���� : {{size.device_name}}</span>
                                        <span class="small">������ : {{size.content_size}}</span>
                                        <span>��ũ : <input name="size_links" v-model="size.link" type="text" class="form-control inline middle"></span>
                                    </li>
                                </ul>
                                <span class="alert alert-danger">
                                    * ���ϴٿ�ε� �Է½� : �ٿ�ε� ��ȣ�� �Է� (javascript ���� �ܾ� �Է� �Ұ�)
                                </span>
                            </td>
                        </tr>
                        <tr v-if="current_content.gubun == 3">
                            <th>�� ����</th>
                            <td>
                                <input id="write_detail_content" name="detail_content" v-model="current_content.detail_content" type="text" class="form-control">
                            </td>
                        </tr>
                        <tr v-if="current_content.gubun == 3">
                            <th>���� ��ũ</th>
                            <td>
                                <input id="write_movie_link" name="movie_link" v-model="current_content.movie_link" type="text" class="form-control">
                                <div class="alert alert-danger">
                                    * ��޿� ��밡��<br/>
                                    * ��޿� : copy embed code ���� (�� : //player.vimeo.com/video/102309330 ) http: ����<br/>
                                    * ������ : �ҽ��ڵ� ���� (�� : http://www.youtube.com/embed/qj4rn1I_dC8 )
                                </div>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
        </div>
    `,
    mounted() {
        //�޷´�ȭâ ����
        const arrDayMin = ["��","��","ȭ","��","��","��","��"];
        const arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
        $("#write_start_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '������', nextText: '������', yearSuffix: '��',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true,
            numberOfMonths: 2,
            showCurrentAtPos: 1,
            showOn: "button"
            //maxDate: "<%=eDt%>"
        });
    },
    data() {return { // ���� ������
        current_content : {
            content_idx : 0,
            detail_content: '',
            mobile_wallpaper_sizes : [],
            pc_wallpaper_sizes : [],
            reg_date: '',
            start_date: '',
            thumbnail: '',
            title: '',
            gubun: '',
            use_yn: true,
            movie_link: ''
        }, 
        mobile_wallpaper_sizes : [], // Mobile ���ȭ�� ������ ����Ʈ
        pc_wallpaper_sizes : [], // PC ���ȭ�� ������ ����Ʈ
    }},
    props: {
        content : {
            content_idx : {type:Number, default:0}, // ������ idx
            detail_content: {type:String, default:''}, // �� ����
            mobile_wallpaper_sizes : {type:Array, default:function(){return [];}}, // Mobile ���ȭ�� ������ ����Ʈ
            pc_wallpaper_sizes : {type:Array, default:function(){return [];}}, // PC ���ȭ�� ������ ����Ʈ
            reg_date: {type:String, default:''}, // �������
            start_date: {type:String, default:''}, // ��������
            thumbnail: {type:String, default:''}, // �����
            title: {type:String, default:''}, // Ÿ��Ʋ
            gubun: {type:String, default:''}, // ����(1:PC,2:Mobile,3:Movie,4:Mobile���)
            use_yn: {type:Boolean, default:true}, // ��뿩��
            movie_link: {type:String, default:''}, // ����ũ
        },
        wallpaper_sizes : { // Default ���ȭ�� ������ ����Ʈ
            mobile_wallpaper_sizes : {type:Array, default:function(){return [];}}, // Mobile
            pc_wallpaper_sizes : {type:Array, default:function(){return [];}} // PC
        }
    },
    watch : {
        content(content) { // ������ ���� �� ���籸�а� set(�˾��Ǿ��� ��)
            if( Object.keys(content).length > 0 ) {
                this.current_content = content;
            } else {
                this.current_content = {
                    content_idx : 0,
                    detail_content: '',
                    mobile_wallpaper_sizes : this.wallpaper_sizes.mobile_wallpaper_sizes,
                    pc_wallpaper_sizes : this.wallpaper_sizes.pc_wallpaper_sizes,
                    reg_date: '',
                    start_date: '',
                    thumbnail: '',
                    title: '',
                    gubun: '',
                    use_yn: true,
                    movie_link: ''
                };
            }
        }
    },
    methods : {
        upload_thumbnail(e) { // ����� ���
            const _this = this;
            if( e.target.value === '' )
                return false;

            const name_arr = e.target.value.split('.');
            const ext = name_arr[name_arr.length - 1];
            const allow_ext_arr = ['jpg', 'gif'];
            if( allow_ext_arr.indexOf(ext.toLowerCase()) < 0 ) {
                alert('jpg, gif Ȯ���� �̹����� ��� �����մϴ�.');
                e.target.value = '';
                return false;
            }

            let api_url;
            if( location.hostname.startsWith('webadmin') ) {
                api_url = 'http://upload.10x10.co.kr';
            } else {
                api_url = 'http://testupload.10x10.co.kr';
            }

            const form_data = new FormData();
            form_data.append('sfImg', e.target.files[0]);
            form_data.append('sName', 'con_viewthumbimg');
            $.ajax({
                type: 'POST',
                url: api_url + '/linkweb/hitchhiker/hitchhiker_imgreg_json.asp',
                processData: false,
                contentType: false,
                crossDomain: true,
                data: form_data,
                success: function(data) {
                    try {
                        const response = JSON.parse(data);
                        if( response.response === 'ok' ) {
                            _this.current_content.thumbnail = response.imgurl;
                        } else {
                            alert('�̹��� ���� �� ������ �߻��߽��ϴ�. (Err: 001)');
                        }
                    } catch(e) {
                        console.log(data, e);
                        alert('�̹��� ���� �� ������ �߻��߽��ϴ�. (Err: 002)');
                    }
                },
                error: function(xhr) {
                    alert('�̹��� ���� �� ������ �߻��߽��ϴ�. (Err: 003)');
                    console.log(xhr.responseText);
                }
            });
        },
        delete_thumbnail() { // ����� ����
            this.current_content.thumbnail = '';
        }
    }
});