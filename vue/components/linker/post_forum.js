Vue.component('POST-FORUM', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:100px;">
                    <col>
                </colgroup>
                <tbody>
                    <!--region ����-->
                    <tr>
                        <th>����</th>
                        <td><input v-model="title" type="text" placeholder="���� ������ �Է����ּ���"></td>
                    </tr>
                    <!--endregion-->
                    <!--region ������-->
                    <tr>
                        <th>������</th>
                        <td><input v-model="subTitle" type="text" placeholder="���� �������� �Է����ּ���"></td>
                    </tr>
                    <!--endregion-->
                    <!--region ����-->
                    <tr>
                        <th>����</th>
                        <td><textarea v-model="description" placeholder="���� ������ �Է����ּ���"></textarea></td>
                    </tr>
                    <!--endregion-->
                    <!--region ��׶��� PC-->
                    <tr>
                        <th>��׶���<br>PC</th>
                        <td :class="['backMedia', {'flex' : backPCType === 'I'}]">
                            <div>
                                <p class="radio-area">
                                    <input v-model="backPCType" value="I" id="backPcImage" type="radio" checked>
                                    <label for="backPcImage">�̹���</label>
                                    <input v-model="backPCType" value="V" id="backPcVideo" type="radio">
                                    <label for="backPcVideo">������</label>
                                </p>
                                <template v-if="backPCType === 'I'">
                                    <button @click="clickPcImageButton" class="linker-btn">�̹��� ÷��</button>
                                    <input @change="changeFile($event, 'pc')" type="file" class="hiddenFile">
                                </template>
                                <input v-else v-model="pcVideo" type="text" placeholder="���� URL�� �Է����ּ���">
                            </div>
                            <div v-if="backPCType === 'I'">
                                <img v-if="pcImage" :src="pcImage" class="preview">
                            </div>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region ��׶��� M-->
                    <tr>
                        <th>��׶���<br>M</th>
                        <td :class="['backMedia', {'flex' : backMobileType === 'I'}]">
                            <div>
                                <p class="radio-area">
                                    <input v-model="backMobileType" value="I" id="backMImage" type="radio" checked>
                                    <label for="backMImage">�̹���</label>
                                    <input v-model="backMobileType" value="V" id="backMVideo" type="radio">
                                    <label for="backMVideo">������</label>
                                </p>
                                <template v-if="backMobileType === 'I'">
                                    <button @click="clickPcImageButton" class="linker-btn">�̹��� ÷��</button>
                                    <input @change="changeFile($event, 'm')" type="file" class="hiddenFile">
                                </template>
                                <input v-else v-model="mobileVideo" type="text" placeholder="���� URL�� �Է����ּ���">
                            </div>
                            <div v-if="backMobileType === 'I'">
                                <img v-if="mobileImage" :src="mobileImage" class="preview">
                            </div>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region ��Ⱓ-->
                    <tr>
                        <th>��Ⱓ</th>
                        <td>
                            <span class="datepicker">
                                <label for="forumStartDate">
                                    <strong>������</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setStartDate" :date="startDate" id="forumStartDate"/>

                                <label for="forumEndDate">
                                    <strong>������</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setEndDate" :date="endDate" id="forumEndDate"/>
                            </span>
                        </td>
                    </tr>
                    <!--endregion-->
                    <!--region ����Ʈ ���⿩��-->
                    <tr>
                        <th>����Ʈ<br>���⿩��</th>
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
                    <!--region ���ļ���-->
                    <tr>
                        <th>���ļ���</th>
                        <td><input v-model="sortNo" type="text" style="width: 100px;"></td>
                    </tr>
                    <!--endregion-->
                    <!--region ���-->
                    <tr>
                        <th>���</th>
                        <td><textarea v-model="note"></textarea></td>
                    </tr>
                    <!--endregion-->
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveForum" class="linker-btn">����</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.modifyForum ) {
            this.setModifyForumData();
        }
    },
    data() {return {
        // region �Է� ������
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
        uploadImageType : '', // ���ε��� �̹��� ����(m, pc)
        pcImage : '', // pc �̹���
        mobileImage : '', // mobile �̹���
        pcVideo : '', // pc ������
        mobileVideo : '', // mobile ������
    }},
    props : {
        //region modifyForum ���� ����
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
        //region apiData ���� ��� API ���� ������
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
        //region isModify ���� �� ����
        isModify() {
            return this.modifyForum !== null;
        },
        //endregion
    },
    methods : {
        //region setStartDate Set ��������
        setStartDate(date) {
            this.startDate = date;
        },
        //endregion
        //region setEndDate Set ��������
        setEndDate(date) {
            this.endDate = date;
        },
        //endregion
        //region saveForum ���� ����
        saveForum() {
            if( !confirm('���� �Ͻðڽ��ϱ�?') )
                return;

            const uri = this.isModify ? '/linker/forum/update' : '/linker/forum';
            this.callApi(2, 'POST', uri, this.apiData, this.successSaveForum);
        },
        successSaveForum(data) {
            if( isNaN(data) ) {
                alert('���� �� ������ �߻��߽��ϴ�.');
            } else {
                alert('���� �Ǿ����ϴ�.');
                this.$emit('saveForum', Number(data));
            }
        },
        //endregion
        //region clickPcImageButton PC �̹��� ÷�� ��ư Ŭ��
        clickPcImageButton(e) {
            e.target.nextElementSibling.click();
        },
        //endregion
        //region changeFile ���� ����
        changeFile(e, type) {
            this.uploadImageType = type;

            const file = e.target.files[0];
            if (!file) {
                this.clearImageFile(e.target);
                return false;
            } else if (!file.type.match('image.*')) {
                this.clearImageFile(e.target);
                alert('�̹��� ���ϸ� ����Ͻ� �� �ֽ��ϴ�.');
                return false;
            }else if(file.size > 5*1024*1024){
                this.clearImageFile(e.target);
                alert('5MB ������ �̹����� ������ּ���');
                return false;
            }

            const imgData = this.createUploadImageData(e.target);
            this.uploadImage(imgData);
        },
        //endregion
        //region uploadImage �̹��� ���ε�
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
                    alert('�̹��� ���ε� �� ������ �߻��߽��ϴ�.\nCode: 002');
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
                alert('�̹��� ���ε� �� ������ �߻��߽��ϴ�.\nCode: 001');
            }
        },
        createUploadImageData(input) {
            const imgData = new FormData();
            imgData.append('image', input.files[0]);
            imgData.append('ch', this.uploadImageType);
            return imgData;
        },
        //endregion
        //region clearImageFile �̹��� �ʱ�ȭ
        clearImageFile(input) {
            if( this.uploadImageType === 'pc' )
                this.pcImage = '';
            else
                this.mobileImage = '';
            input.value = '';
            this.uploadImageType = '';
        },
        //endregion
        //region setModifyForumData Set ���� ���� ������
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