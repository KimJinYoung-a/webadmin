/*
    히치하이커 컨텐츠 관리 모달 Body
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
                            <th>구분</th>
                            <td>
                                <select id="write_gubun" name="gubun" class="form-control inline small" v-model="current_content.gubun">
                                    <option value="">구분</option>
                                    <option value="1">PC</option>
                                    <option value="2">MOBILE</option>
                                    <option value="3">MOVIE</option>
                                    <option value="4">MOBILE배경</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>타이틀</th>
                            <td>
                                <input id="write_title" name="title" v-model="current_content.title" type="text" class="form-control" placeholder="타이틀 입력">
                            </td>
                        </tr>
                        <tr>
                            <th>시작일</th>
                            <td>
                                <input id="write_start_date" name="start_date" v-model="current_content.start_date" type="text" class="form-control small inline" placeholder="시작일 입력">
                            </td>
                        </tr>
                        <tr>
                            <th>사용 여부</th>
                            <td>
                                <div class="form-check">
                                    <input v-model="current_content.use_yn" value="true" name="use_yn" id="use_yn_true" class="form-check-input" type="radio">
                                    <label class="form-check-label" for="use_yn_true">사용함</label>
                                    <input v-model="current_content.use_yn" value="false" name="use_yn" id="use_yn_false" class="form-check-input" type="radio">
                                    <label class="form-check-label" for="use_yn_false">사용안함</label>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <th>썸네일</th>
                            <td>
                                <div @click="delete_thumbnail" class="thumbnail-area">
                                    <img v-show="current_content.thumbnail" :src="current_content.thumbnail" class="thumbnail">
                                    <div class="overlay">삭제</div>
                                </div>
                                <input id="write_thumbnail" @change="upload_thumbnail" type="file" class="form-control inline middle">
                                <input name="thumbnail" type="hidden" v-model="current_content.thumbnail">
                            </td>
                        </tr>
                        <tr v-if="current_content.gubun == 1 || current_content.gubun == 2">
                            <th>이미지 링크</th>
                            <td>
                                <ul v-if="current_content.gubun == 1" class="ul-write-hitchhiker">
                                    <li v-for="size in current_content.pc_wallpaper_sizes">
                                    <span class="small">사이즈 : {{size.content_size}}</span>
                                        <input name="size_link_idxs" v-model="size.link_idx" type="hidden">
                                        <input name="size_device_idxs" v-model="size.device_idx" type="hidden">
                                        <span>링크 : <input name="size_links" v-model="size.link" type="text" class="form-control inline middle"></span>
                                    </li>
                                </ul>
                                <ul v-if="current_content.gubun == 2" class="ul-write-hitchhiker">
                                    <li v-for="size in current_content.mobile_wallpaper_sizes">
                                        <input name="size_link_idxs" v-model="size.link_idx" type="hidden">
                                        <input name="size_device_idxs" v-model="size.device_idx" type="hidden">
                                        <span class="small">대표기종 : {{size.device_name}}</span>
                                        <span class="small">사이즈 : {{size.content_size}}</span>
                                        <span>링크 : <input name="size_links" v-model="size.link" type="text" class="form-control inline middle"></span>
                                    </li>
                                </ul>
                                <span class="alert alert-danger">
                                    * 파일다운로드 입력시 : 다운로드 번호만 입력 (javascript 등의 단어 입력 불가)
                                </span>
                            </td>
                        </tr>
                        <tr v-if="current_content.gubun == 3">
                            <th>상세 내용</th>
                            <td>
                                <input id="write_detail_content" name="detail_content" v-model="current_content.detail_content" type="text" class="form-control">
                            </td>
                        </tr>
                        <tr v-if="current_content.gubun == 3">
                            <th>영상 링크</th>
                            <td>
                                <input id="write_movie_link" name="movie_link" v-model="current_content.movie_link" type="text" class="form-control">
                                <div class="alert alert-danger">
                                    * 비메오 사용가능<br/>
                                    * 비메오 : copy embed code 복사 (예 : //player.vimeo.com/video/102309330 ) http: 제외<br/>
                                    * 유투브 : 소스코드 복사 (예 : http://www.youtube.com/embed/qj4rn1I_dC8 )
                                </div>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
        </div>
    `,
    mounted() {
        //달력대화창 설정
        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#write_start_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true,
            numberOfMonths: 2,
            showCurrentAtPos: 1,
            showOn: "button"
            //maxDate: "<%=eDt%>"
        });
    },
    data() {return { // 현재 컨텐츠
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
        mobile_wallpaper_sizes : [], // Mobile 배경화면 사이즈 리스트
        pc_wallpaper_sizes : [], // PC 배경화면 사이즈 리스트
    }},
    props: {
        content : {
            content_idx : {type:Number, default:0}, // 컨텐츠 idx
            detail_content: {type:String, default:''}, // 상세 내용
            mobile_wallpaper_sizes : {type:Array, default:function(){return [];}}, // Mobile 배경화면 사이즈 리스트
            pc_wallpaper_sizes : {type:Array, default:function(){return [];}}, // PC 배경화면 사이즈 리스트
            reg_date: {type:String, default:''}, // 등록일자
            start_date: {type:String, default:''}, // 시작일자
            thumbnail: {type:String, default:''}, // 썸네일
            title: {type:String, default:''}, // 타이틀
            gubun: {type:String, default:''}, // 구분(1:PC,2:Mobile,3:Movie,4:Mobile배경)
            use_yn: {type:Boolean, default:true}, // 사용여부
            movie_link: {type:String, default:''}, // 영상링크
        },
        wallpaper_sizes : { // Default 배경화면 사이즈 리스트
            mobile_wallpaper_sizes : {type:Array, default:function(){return [];}}, // Mobile
            pc_wallpaper_sizes : {type:Array, default:function(){return [];}} // PC
        }
    },
    watch : {
        content(content) { // 컨텐츠 변경 시 현재구분값 set(팝업되었을 때)
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
        upload_thumbnail(e) { // 썸네일 등록
            const _this = this;
            if( e.target.value === '' )
                return false;

            const name_arr = e.target.value.split('.');
            const ext = name_arr[name_arr.length - 1];
            const allow_ext_arr = ['jpg', 'gif'];
            if( allow_ext_arr.indexOf(ext.toLowerCase()) < 0 ) {
                alert('jpg, gif 확장자 이미지만 등록 가능합니다.');
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
                            alert('이미지 저장 중 오류가 발생했습니다. (Err: 001)');
                        }
                    } catch(e) {
                        console.log(data, e);
                        alert('이미지 저장 중 오류가 발생했습니다. (Err: 002)');
                    }
                },
                error: function(xhr) {
                    alert('이미지 저장 중 오류가 발생했습니다. (Err: 003)');
                    console.log(xhr.responseText);
                }
            });
        },
        delete_thumbnail() { // 썸네일 삭제
            this.current_content.thumbnail = '';
        }
    }
});