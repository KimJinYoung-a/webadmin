Vue.component('POST-ITEM', {
    template : `
        <div class="post-item">
            <div class="popContV19">
                <table class="tableV19A tabbarTemplate">
                    <!--region colgroup-->
                    <colgroup>
                        <col style="width:150px;">
                        <col style="width:auto;">
                    </colgroup>
                    <!--endregion-->
                    <tbody>
                        <tr>
                            <th>타이틀</th>
                            <td><input v-model="title" type="text" placeholder="타이틀"></td>
                        </tr>
                        <tr>
                            <th>서브 타이틀</th>
                            <td><input v-model="subTitle" type="text" placeholder="서브 타이틀"></td>
                        </tr>
                        <tr>
                            <th>링크</th>
                            <td><input v-model="link" type="text" placeholder="링크(ex:/event/eventmain.asp?eventid=123456)"></td>
                        </tr>
                        <tr>
                            <th>이미지</th>
                            <td>
                                <button @click="clickImageButton" class="btn4 btnBlue1">이미지 등록</button>
                                <input @change="changeImageFile" id="itemImageFile" type="file" style="display: none;">
                                <img v-if="image" :src="image" class="back-image"/>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!--region 저장,취소-->
            <div class="popBtnWrapV19">
                <button v-if="item" @click="modifyItem" class="btn4 btnBlue1">수정</button>
                <button v-else @click="postItem" class="btn4 btnBlue1">저장</button>
                <button @click="$emit('cancel')" class="btn4 btnWhite1">취소</button>
            </div>
            <!--endregion-->
        </div>
    `,
    mounted() {
        if( this.item )
            this.setModifyData();
    },
    data() {return {
        title : '',
        subTitle : '',
        link : '',
        image : '',
    }},
    props : {
        masterIndex : { type:Number, default:0 },
        //region item 아이템
        item : {
            itemIndex : { type:Number, default:0 },
            title : { type:String, default:'' },
            subTitle : { type:String, default:'' },
            link : { type:String, default:'' },
            image : { type:String, default:'' },
            sort : { type:Number, default:1 },
            selected : { type:Boolean, default:false },
        },
        //endregion
    },
    methods : {
        //region clickImageButton 이미지 등록 버튼 클릭
        clickImageButton() {
            document.getElementById('itemImageFile').click();
        },
        //endregion
        //region changeImageFile 이미지 변경
        changeImageFile(e) {
            if( e.target.value === '' ) {
                this.image = '';
                return false;
            }

            const _this = this;
            const imgData = this.createUploadImageData();
            this.callAjaxUploadImage(imgData, data => {
                const response = JSON.parse(data);

                if (response.response === 'ok') {
                    _this.image = response.filePath;
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
        createUploadImageData() {
            const imgData = new FormData();
            imgData.append('image', document.getElementById("itemImageFile").files[0]);
            return imgData;
        },
        //endregion
        //region postItem 아이템 등록
        postItem() {
            if( confirm('등록 하시겠습니까?') && this.validatePostItem() ) {
                this.callApi(2, 'POST', '/event/contents/tabbar/item', this.createPostItemData(), this.successPostItem);
            }
        },
        validatePostItem() {
            if( !this.title ) {
                alert('타이틀을 입력 해 주세요');
                return false;
            }
            return true;
        },
        createPostItemData() {
            return {
                masterIndex : this.masterIndex,
                title : this.title,
                subTitle : this.subTitle,
                link : this.link,
                image : this.image
            };
        },
        successPostItem() {
            this.$emit('postItem');
        },
        //endregion
        //region modifyItem 아이템 수정
        modifyItem() {
            if( confirm('수정 하시겠습니까?') && this.validatePostItem() ) {
                this.callApi(2, 'POST', '/event/contents/tabbar/item/update', this.createPutItemData(), this.successPostItem);
            }
        },
        createPutItemData() {
            return {
                itemIndex : this.item.itemIndex,
                title : this.title,
                subTitle : this.subTitle,
                link : this.link,
                image : this.image
            };
        },
        //endregion
        //region setModifyData Set 수정 아이템 데이터
        setModifyData() {
            this.title = this.item.title;
            this.subTitle = this.item.subTitle;
            this.link = this.item.link;
            this.image = this.decodeBase64(this.item.image);
        },
        //endregion
    }
});