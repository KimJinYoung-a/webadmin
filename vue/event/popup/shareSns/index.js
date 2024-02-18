const app = new Vue({
    el : '#app',
    mixins : [api_mixin],
    template : /*html*/`
        <div class="popV19">
            <div class="popHeadV19">
                <h1>SNS ���� ����</h1>
            </div>
            <div class="popContV19">
                <table class="tableV19A">
                    <colgroup>
                        <col style="width:150px;">
                        <col style="width:auto;">
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>īī����</th>
                            <td>
                                <div class="bMar05">
                                    <input v-model="kakaoTitle" type="text" class="formControl" placeholder="Ÿ��Ʋ"><br>
                                </div>
                                <div class="bMar05">
                                    <input v-model="kakaoDescription" type="text" class="formControl" placeholder="����"><br>
                                </div>
                                <button @click="clickImageButton('kakao')" class="btn4 btnBlue1">�̹��� ���</button>
                                <input id="kakaoImage" @change="changeImage($event, 'kakao')" type="file" style="display:none;">
                                <button v-show="kakaoImage" @click="kakaoImage = ''" class="btn4 btnGrey1 lMar05">����</button>
                                <img v-if="kakaoImage" :src="kakaoImage" style="display: block;max-height: 200px;margin-top: 10px;">
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div class="popBtnWrapV19">
                <button @click="close" class="btn4 btnWhite1">���</button>
                <button @click="saveShareSnsInfo" class="btn4 btnBlue1">����</button>
            </div>
        </div>
    `,
    data() {return {
        eventCode : 0,
        kakaoTitle : '',
        kakaoDescription : '',
        kakaoImage : '',
    }},
    mounted() {
        this.eventCode = Number(this.parameters.eC);
        this.getShareSnsInfo();
    },
    methods : {
        //region getShareSnsInfo ���� ���� ��ȸ
        getShareSnsInfo() {
            const url = `/event/share/sns/${this.eventCode}`;
            this.callApi(2, 'GET', url, null, this.successGetShareSnsInfo);
        },
        successGetShareSnsInfo(data) {
            this.kakaoTitle = data.kakaoTitle;
            this.kakaoDescription = data.kakaoDescription;
            this.kakaoImage = data.kakaoImage;
        },
        //endregion
        //region saveShareSnsInfo ���� ���� ����
        saveShareSnsInfo() {
            if( !confirm('���� �Ͻðڽ��ϱ�?') )
                return false;

            this.callApi(2, 'POST', '/event/share/sns', this.createSaveShareSnsData(),
                this.successSaveShareSnsData)
        },
        createSaveShareSnsData() {
            return {
                eventCode : this.eventCode,
                kakaoTitle : this.kakaoTitle,
                kakaoDescription : this.kakaoDescription,
                kakaoImage : this.kakaoImage
            };
        },
        successSaveShareSnsData() {
            alert('���� �Ǿ����ϴ�.');
            window.document.domain = '10x10.co.kr'
            opener.document.location.reload();
            self.close();
        },
        //endregion
        //region close �˾� �ݱ�
        close() {
            self.close();
        },
        //endregion

        // region �̹��� ����
        //region changeImage �̹��� ����
        changeImage(e, type) {
            const file = e.target.files[0];
            if( !file ) {
                this[type + 'Image'] = '';
                return false;
            }

            const _this = this;
            const imgData = this.createUploadImageData(type);
            this.callAjaxUploadImage(imgData, data => {
                const response = JSON.parse(data);

                if (response.response === 'ok') {
                    _this[type + 'Image'] = response.filePath;
                } else {
                    alert(response.message);
                }
            });
        },
        //endregion
        //region callAjaxUploadImage �̹��� ���ε� ���ε弭�� ajax ȣ��
        callAjaxUploadImage(imgData, success) {
            $.ajax({
                url: '//oimgstatic.10x10.co.kr/linkweb/event/tabbar_image_upload.asp'
                , type: 'POST'
                , processData: false
                , contentType: false
                , data: imgData
                , crossDomain: true
                , success : success
                , error : e => {
                    alert('�̹����� ���ε� �ϴ� �� ������ �߻��߽��ϴ�.');
                    console.log(e);
                }
            });
        },
        //endregion
        //region createUploadImageData �̹��� ���ε� Data ����
        createUploadImageData(type) {
            const imgData = new FormData();
            imgData.append('image', document.getElementById(type + 'Image').files[0]);
            return imgData;
        },
        //endregion
        //region clickBackImageButton �̹��� ��� ��ư Ŭ��
        clickImageButton(type) {
            document.getElementById(type + 'Image').click();
        },
        //endregion
        // endregion
    }
});