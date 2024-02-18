const app = new Vue({
    el : '#app',
    mixins : [api_mixin],
    template : /*html*/`
        <div class="popV19">
            <div class="popHeadV19">
                <h1>SNS 공유 설정</h1>
            </div>
            <div class="popContV19">
                <table class="tableV19A">
                    <colgroup>
                        <col style="width:150px;">
                        <col style="width:auto;">
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>카카오톡</th>
                            <td>
                                <div class="bMar05">
                                    <input v-model="kakaoTitle" type="text" class="formControl" placeholder="타이틀"><br>
                                </div>
                                <div class="bMar05">
                                    <input v-model="kakaoDescription" type="text" class="formControl" placeholder="설명"><br>
                                </div>
                                <button @click="clickImageButton('kakao')" class="btn4 btnBlue1">이미지 등록</button>
                                <input id="kakaoImage" @change="changeImage($event, 'kakao')" type="file" style="display:none;">
                                <button v-show="kakaoImage" @click="kakaoImage = ''" class="btn4 btnGrey1 lMar05">삭제</button>
                                <img v-if="kakaoImage" :src="kakaoImage" style="display: block;max-height: 200px;margin-top: 10px;">
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div class="popBtnWrapV19">
                <button @click="close" class="btn4 btnWhite1">취소</button>
                <button @click="saveShareSnsInfo" class="btn4 btnBlue1">저장</button>
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
        //region getShareSnsInfo 공유 정보 조회
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
        //region saveShareSnsInfo 공유 정보 저장
        saveShareSnsInfo() {
            if( !confirm('저장 하시겠습니까?') )
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
            alert('저장 되었습니다.');
            window.document.domain = '10x10.co.kr'
            opener.document.location.reload();
            self.close();
        },
        //endregion
        //region close 팝업 닫기
        close() {
            self.close();
        },
        //endregion

        // region 이미지 관련
        //region changeImage 이미지 변경
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
        //region callAjaxUploadImage 이미지 업로드 업로드서버 ajax 호출
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
                    alert('이미지를 업로드 하는 중 에러가 발생했습니다.');
                    console.log(e);
                }
            });
        },
        //endregion
        //region createUploadImageData 이미지 업로드 Data 생성
        createUploadImageData(type) {
            const imgData = new FormData();
            imgData.append('image', document.getElementById(type + 'Image').files[0]);
            return imgData;
        },
        //endregion
        //region clickBackImageButton 이미지 등록 버튼 클릭
        clickImageButton(type) {
            document.getElementById(type + 'Image').click();
        },
        //endregion
        // endregion
    }
});