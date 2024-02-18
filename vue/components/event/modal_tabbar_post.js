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
                            <th>Ÿ��Ʋ</th>
                            <td><input v-model="title" type="text" placeholder="Ÿ��Ʋ"></td>
                        </tr>
                        <tr>
                            <th>���� Ÿ��Ʋ</th>
                            <td><input v-model="subTitle" type="text" placeholder="���� Ÿ��Ʋ"></td>
                        </tr>
                        <tr>
                            <th>��ũ</th>
                            <td><input v-model="link" type="text" placeholder="��ũ(ex:/event/eventmain.asp?eventid=123456)"></td>
                        </tr>
                        <tr>
                            <th>�̹���</th>
                            <td>
                                <button @click="clickImageButton" class="btn4 btnBlue1">�̹��� ���</button>
                                <input @change="changeImageFile" id="itemImageFile" type="file" style="display: none;">
                                <img v-if="image" :src="image" class="back-image"/>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!--region ����,���-->
            <div class="popBtnWrapV19">
                <button v-if="item" @click="modifyItem" class="btn4 btnBlue1">����</button>
                <button v-else @click="postItem" class="btn4 btnBlue1">����</button>
                <button @click="$emit('cancel')" class="btn4 btnWhite1">���</button>
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
        //region item ������
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
        //region clickImageButton �̹��� ��� ��ư Ŭ��
        clickImageButton() {
            document.getElementById('itemImageFile').click();
        },
        //endregion
        //region changeImageFile �̹��� ����
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
        createUploadImageData() {
            const imgData = new FormData();
            imgData.append('image', document.getElementById("itemImageFile").files[0]);
            return imgData;
        },
        //endregion
        //region postItem ������ ���
        postItem() {
            if( confirm('��� �Ͻðڽ��ϱ�?') && this.validatePostItem() ) {
                this.callApi(2, 'POST', '/event/contents/tabbar/item', this.createPostItemData(), this.successPostItem);
            }
        },
        validatePostItem() {
            if( !this.title ) {
                alert('Ÿ��Ʋ�� �Է� �� �ּ���');
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
        //region modifyItem ������ ����
        modifyItem() {
            if( confirm('���� �Ͻðڽ��ϱ�?') && this.validatePostItem() ) {
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
        //region setModifyData Set ���� ������ ������
        setModifyData() {
            this.title = this.item.title;
            this.subTitle = this.item.subTitle;
            this.link = this.item.link;
            this.image = this.decodeBase64(this.item.image);
        },
        //endregion
    }
});