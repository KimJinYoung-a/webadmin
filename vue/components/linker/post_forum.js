Vue.component('POST-FORUM', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:100px;">
                    <col>
                </colgroup>
                <tbody>
                    <!--region 제목-->
                    <tr>
                        <th>제목</th>
                        <td><input v-model="title" type="text" placeholder="포럼 제목을 입력해주세요"></td>
                    </tr>
                    <!--endregion-->
                    <!--region 부제목-->
                    <tr>
                        <th>부제목</th>
                        <td><input v-model="subTitle" type="text" placeholder="포럼 부제목을 입력해주세요"></td>
                    </tr>
                    <!--endregion-->
                    <!--region 설명-->
                    <tr>
                        <th>설명</th>
                        <td><textarea v-model="description" placeholder="포럼 설명을 입력해주세요"></textarea></td>
                    </tr>
                    <!--endregion-->
                    <!--region 백그라운드 PC-->
                    <tr>
                        <th>백그라운드<br>PC</th>
                        <td :class="['backMedia', {'flex' : backPCType === 'I'}]">
                            <div>
                                <p class="radio-area">
                                    <input v-model="backPCType" value="I" id="backPcImage" type="radio" checked>
                                    <label for="backPcImage">이미지</label>
                                    <input v-model="backPCType" value="V" id="backPcVideo" type="radio">
                                    <label for="backPcVideo">동영상</label>
                                </p>
                                <template v-if="backPCType === 'I'">
                                    <button @click="clickPcImageButton" class="linker-btn">이미지 첨부</button>
                                    <input @change="changeFile($event, 'pc')" type="file" class="hiddenFile">
                                </template>
                                <input v-else v-model="pcVideo" type="text" placeholder="영상 URL을 입력해주세요">
                            </div>
                            <div v-if="backPCType === 'I'">
                                <img v-if="pcImage" :src="pcImage" class="preview">
                            </div>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region 백그라운드 M-->
                    <tr>
                        <th>백그라운드<br>M</th>
                        <td :class="['backMedia', {'flex' : backMobileType === 'I'}]">
                            <div>
                                <p class="radio-area">
                                    <input v-model="backMobileType" value="I" id="backMImage" type="radio" checked>
                                    <label for="backMImage">이미지</label>
                                    <input v-model="backMobileType" value="V" id="backMVideo" type="radio">
                                    <label for="backMVideo">동영상</label>
                                </p>
                                <template v-if="backMobileType === 'I'">
                                    <button @click="clickPcImageButton" class="linker-btn">이미지 첨부</button>
                                    <input @change="changeFile($event, 'm')" type="file" class="hiddenFile">
                                </template>
                                <input v-else v-model="mobileVideo" type="text" placeholder="영상 URL을 입력해주세요">
                            </div>
                            <div v-if="backMobileType === 'I'">
                                <img v-if="mobileImage" :src="mobileImage" class="preview">
                            </div>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region 운영기간-->
                    <tr>
                        <th>운영기간</th>
                        <td>
                            <span class="datepicker">
                                <label for="forumStartDate">
                                    <strong>시작일</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setStartDate" :date="startDate" id="forumStartDate"/>

                                <label for="forumEndDate">
                                    <strong>종료일</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setEndDate" :date="endDate" id="forumEndDate"/>
                            </span>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region 프론트 노출여부-->
                    <tr>
                        <th>프론트<br>노출여부</th>
                        <td>
                            <p class="radio-area">
                                <input v-model="frontShowYn" value="Y" id="showY" type="radio" checked>
                                <label for="showY">Y</label>
                                <input v-model="frontShowYn" value="N" id="showN" type="radio">
                                <label for="showN">N</label>
                            </p>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region 정렬순서-->
                    <tr>
                        <th>정렬순서</th>
                        <td><input v-model="sortNo" type="text" style="width: 100px;"></td>
                    </tr>
                    <!--endregion-->
                    <!--region 비고-->
                    <tr>
                        <th>비고</th>
                        <td><textarea v-model="note"></textarea></td>
                    </tr>
                    <!--endregion-->
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveForum" class="linker-btn">저장</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.modifyForum ) {
            this.setModifyForumData();
        }
    },
    data() {return {
        // region 입력 데이터
        forumIndex : null,
        title : '',
        subTitle : '',
        description : '',
        backPCType : 'I',
        backPCValue : '',
        backMobileType : 'I',
        backMobileValue : '',
        startDate : '',
        endDate : '',
        frontShowYn : 'Y',
        sortNo : '',
        note : '',
        // endregion
        uploadImageType : '', // 업로드할 이미지 유형(m, pc)
        pcImage : '', // pc 이미지
        mobileImage : '', // mobile 이미지
        pcVideo : '', // pc 동영상
        mobileVideo : '', // mobile 동영상
    }},
    props : {
        //region modifyForum 수정 포럼
        modifyForum : {
            forumIdx : { type:Number, default:0 },
            subTitle : { type:String, default:'' },
            description : { type:String, default:'' },
            startDate : { type:String, default:'' },
            endDate : { type:String, default:'' },
            useYn : { type:Boolean, default:false },
            sortNo : { type:Number, default:0 },
            note : { type:String, default:'' },
            backgroundMediaTypePc : { type:String, default:'I' },
            backgroundMediaValuePc : { type:String, default:'' },
            backgroundMediaTypeM : { type:String, default:'I' },
            backgroundMediaValueM : { type:String, default:'' },
        },
        //endregion
    },
    computed : {
        //region apiData 포럼 등록 API 전송 데이터
        apiData() {
            return {
                forumIndex : this.forumIndex,
                title : this.title,
                subTitle : this.subTitle,
                description : this.description,
                startDate : this.startDate,
                endDate : this.endDate,
                backgroundMediaTypePc : this.backPCType,
                backgroundMediaTypeM : this.backMobileType,
                backgroundMediaValuePc : this.backPCValue,
                backgroundMediaValueM : this.backMobileValue,
                useYn : this.frontShowYn === 'Y',
                sortNo : isNaN(this.sortNo) ? 0 : this.sortNo,
                note : this.note
            }
        },
        //endregion
        //region isModify 수정 중 여부
        isModify() {
            return this.modifyForum !== null;
        },
        //endregion
    },
    methods : {
        //region setStartDate Set 시작일자
        setStartDate(date) {
            this.startDate = date;
        },
        //endregion
        //region setEndDate Set 종료일자
        setEndDate(date) {
            this.endDate = date;
        },
        //endregion
        //region saveForum 포럼 저장
        saveForum() {
            if( !confirm('저장 하시겠습니까?') )
                return;

            const uri = this.isModify ? '/linker/forum/update' : '/linker/forum';
            this.callApi(2, 'POST', uri, this.apiData, this.successSaveForum);
        },
        successSaveForum(data) {
            if( isNaN(data) ) {
                alert('저장 중 에러가 발생했습니다.');
            } else {
                alert('저장 되었습니다.');
                this.$emit('saveForum', Number(data));
            }
        },
        //endregion
        //region clickPcImageButton PC 이미지 첨부 버튼 클릭
        clickPcImageButton(e) {
            e.target.nextElementSibling.click();
        },
        //endregion
        //region changeFile 파일 변경
        changeFile(e, type) {
            this.uploadImageType = type;

            const file = e.target.files[0];
            if (!file) {
                this.clearImageFile(e.target);
                return false;
            } else if (!file.type.match('image.*')) {
                this.clearImageFile(e.target);
                alert('이미지 파일만 등록하실 수 있습니다.');
                return false;
            }else if(file.size > 5*1024*1024){
                this.clearImageFile(e.target);
                alert('5MB 이하의 이미지를 등록해주세요');
                return false;
            }

            const imgData = this.createUploadImageData(e.target);
            this.uploadImage(imgData);
        },
        //endregion
        //region uploadImage 이미지 업로드
        uploadImage(imgData) {
            $.ajax({
                url: '//oimgstatic.10x10.co.kr/linkweb/linker/upload_json.asp'
                , type: 'POST'
                , processData: false
                , contentType: false
                , data: imgData
                , crossDomain: true
                , success : this.successUploadImage
                , error : e => {
                    alert('이미지 업로드 중 에러가 발생했습니다.\nCode: 002');
                    console.log(e);
                }
            });
        },
        successUploadImage(data) {
            try {
                const result = JSON.parse(data);
                if( result.response === 'ok' ) {
                    if( this.uploadImageType === 'pc' ) {
                        this.pcImage = result.filePath;
                        this.backPCValue = this.pcImage;
                    } else {
                        this.mobileImage = result.filePath;
                        this.backMobileValue = this.mobileImage;
                    }
                } else {
                    alert(result.message);
                }
            } catch(e) {
                alert('이미지 업로드 중 에러가 발생했습니다.\nCode: 001');
            }
        },
        createUploadImageData(input) {
            const imgData = new FormData();
            imgData.append('image', input.files[0]);
            imgData.append('ch', this.uploadImageType);
            return imgData;
        },
        //endregion
        //region clearImageFile 이미지 초기화
        clearImageFile(input) {
            if( this.uploadImageType === 'pc' )
                this.pcImage = '';
            else
                this.mobileImage = '';
            input.value = '';
            this.uploadImageType = '';
        },
        //endregion
        //region setModifyForumData Set 수정 포럼 데이터
        setModifyForumData() {
            this.forumIndex = this.modifyForum.forumIdx;
            this.title = this.modifyForum.title;
            this.subTitle = this.modifyForum.subTitle;
            this.description = this.modifyForum.description;
            this.startDate = this.getLocalDateTimeFormat(this.modifyForum.startDate, 'yyyy-MM-dd');
            this.endDate = this.getLocalDateTimeFormat(this.modifyForum.endDate, 'yyyy-MM-dd');
            this.frontShowYn = this.modifyForum.useYn ? 'Y' : 'N';
            this.sortNo = this.modifyForum.sortNo;
            this.note = this.modifyForum.note;

            this.setBackPCValues();
            this.setBackMobileValues();
        },
        setBackPCValues() {
            this.backPCType = this.modifyForum.backgroundMediaTypePc;
            if( this.backPCType === 'I' ) {
                this.pcImage = this.modifyForum.backgroundMediaValuePc;
                this.backPCValue = this.pcImage;
            } else {
                this.pcVideo = this.modifyForum.backgroundMediaValuePc;
                this.backPCValue = this.pcVideo;
            }
        },
        setBackMobileValues() {
            this.backMobileType = this.modifyForum.backgroundMediaTypeM;
            if( this.backMobileType === 'I' ) {
                this.mobileImage = this.modifyForum.backgroundMediaValueM;
                this.backMobileValue = this.mobileImage;
            } else {
                this.mobileVideo = this.modifyForum.backgroundMediaValueM;
                this.backMobileValue = this.mobileVideo;
            }
        },
        //endregion
    }
});