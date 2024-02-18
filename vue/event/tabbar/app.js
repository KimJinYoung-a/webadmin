const app = new Vue({
    el : '#app',
    mixins : [api_mixin],
    template : /*html*/`
        <div class="popV19">
            <div class="popHeadV19">
                <h1>�ǹ�</h1>
            </div>
            <div class="popContV19">
                <!--region ��� ��-->
                <div class="tabV19">
                    <ul>
                        <li :class="{'selected' : device === 'M'}"><a @click="device = 'M'">Mobile / App</a></li>
                        <li :class="{'selected' : device === 'W'}"><a @click="device = 'W'">PC</a></li>
                    </ul>
                </div>
                <!--endregion-->
                <table class="tableV19A tabbarTemplate">
                    <!--region colgroup-->
                    <colgroup>
                        <col style="width:150px;">
                        <col style="width:auto;">
                        <col v-if="device === 'M'" style="width:600px;">
                    </colgroup>
                    <!--endregion-->
                    <tbody>
                        
                        <tr v-if="device === 'W'">
                            <td class="preview pc" colspan="2" align="center">
                                <div class="sliderArea" :style="sliderAreaStyle">
                                    <div class="swiper-container" :style="swiperContainerStyle">
                                        <ul class="swiper-wrapper">
                                            <li v-for="item in items" class="swiper-slide" :style="slideStyle(item.selected)">
                                                <span>{{item.title}}</span>
                                            </li>
                                        </ul>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-prev">����</button>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-next">����</button>
                                    </div>
                                </div>
                                <div class="etcArea" :style="previewEtcAreaStyle">
                                    <img v-if="sliderStyle[device].previewBackImage" :src="sliderStyle[device].previewBackImage"/>
                                </div>
                                <div class="inputArea">
                                    <input v-model="sliderStyle[device].previewBackColor" type="text" placeholder="��� �� �ڵ�(#����)">
                                    <button v-if="!sliderStyle[device].previewBackImage" @click="clickPreviewBackImageButton" class="btn4 btnBlue1">�̸����� �̹��� ���</button>
                                    <button v-else @click="sliderStyle[device].previewBackImage = ''" class="btn4 btnBlue1">�̸����� �̹��� ����</button>
                                    <input type="file" id="previewBackFile" @change="changePreviewBackFile" style="display: none;">
                                </div>
                            </td>
                        </tr>
                        
                        <!--region ����, �̸�����-->
                        <tr>
                            <th>���</th>
                            <td><button @click="openManageItemsModal" class="btn4 btnBlue1">��� ����</button></td>
                            <!--region �����&�� �̸�����-->
                            <td v-if="device === 'M'" class="preview mobile" rowspan="0" align="center">
                                <div class="sliderArea" :style="sliderAreaStyle">
                                    <div class="swiper-container" :style="swiperContainerStyle">
                                        <ul class="swiper-wrapper">
                                            <li v-for="item in items" class="swiper-slide" :style="slideStyle(item.selected)">
                                                <span>{{item.title}}</span>
                                            </li>
                                        </ul>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-prev">����</button>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-next">����</button>
                                    </div>
                                </div>
                                <div class="etcArea" :style="previewEtcAreaStyle">
                                    <img :src="sliderStyle[device].previewBackImage"/>
                                </div>
                                <div class="inputArea">
                                    <input v-model="sliderStyle[device].previewBackColor" type="text" placeholder="��� �� �ڵ�(#����)">
                                    <button v-if="!sliderStyle[device].previewBackImage" @click="clickPreviewBackImageButton" class="btn4 btnBlue1">�̸����� �̹��� ���</button>
                                    <button v-else @click="sliderStyle[device].previewBackImage = ''" class="btn4 btnBlue1">�̸����� �̹��� ����</button>
                                    <input type="file" id="previewBackFile" @change="changePreviewBackFile" style="display: none;">
                                </div>
                            </td>
                            <!--endregion-->
                        </tr>
                        <!--endregion-->
                        <!--region ���-->
                        <tr>
                            <th>���</th>
                            <td>
                                <p class="backTypeArea">
                                    <label class="chedkLabel"><input v-model="sliderStyle[device].backType" value="color" type="radio"> ���ڵ�</label>
                                    <label class="chedkLabel"><input v-model="sliderStyle[device].backType" value="image" type="radio"> �̹���</label>
                                </p>
                                <input v-show="sliderStyle[device].backType === 'color'" v-model="sliderStyle[device].backgroundColor" type="text" class="short" placeholder="#����">
                                <p v-show="sliderStyle[device].backType !== 'color'">
                                    <button @click="clickBackImageButton" class="btn4 btnBlue1">�̹��� ���</button>
                                    <input @change="changeBackFile" id="backFile" type="file" style="display: none;">
                                    <img v-if="sliderStyle[device].backImage" :src="sliderStyle[device].backImage" class="back-image"/>
                                </p>
                            </td>
                        </tr>
                        <!--endregion-->
                        <!--region �۾�-->
                        <tr>
                            <th>�۾� �� �ڵ�</th>
                            <td>
                                <p class="settingArea">
                                    <strong>����</strong>
                                    <input v-model="sliderStyle[device].selectedFontColor" type="text" class="short" placeholder="#����">
                                </p>
                                <p class="settingArea">
                                    <strong>����</strong>
                                    <input v-model="sliderStyle[device].unSelectedFontColor" type="text" class="short" placeholder="#����">
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <th>�۾� ũ��</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].fontSize" @keydown="upDownFontSize"> px</td>
                        </tr>
                        <tr>
                            <th>���� �۾� ȿ��</th>
                            <td>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].selectedFontEffect" value="" type="radio"> ����</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].selectedFontEffect" value="reddot" type="radio"> ���� ������</label>
                            </td>
                        </tr>
                        <!--endregion-->
                        <!--region ȭ��ǥ-->
                        <tr>
                            <th>ȭ��ǥ ����</th>
                            <td>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].showArrow" :value="true" type="radio"> Y</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].showArrow" :value="false" type="radio"> N</label>
                            </td>
                        </tr>
                        <tr v-if="sliderStyle[device].showArrow">
                            <th>ȭ��ǥ �� �ڵ�</th>
                            <td><input type="text" class="short" placeholder="#����"></td>
                        </tr>
                        <!--endregion-->
                        <!--region �����̴�-->
                        <tr>
                            <th>�����̴� ����</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].sliderHeight" @keydown="upDownSliderHeight"> px</td>
                        </tr>
                        <tr>
                            <th>�����̴� ����</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].sliderWidth" @keydown="upDownSliderWidth" @blur="refreshSwiper"> %</td>
                        </tr>
                        <tr>
                            <th>�����̴� ����</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].sliderSpace" @keydown="upDownSliderSpace" @blur="refreshSwiper"> px</td>
                        </tr>
                        <tr v-show="sliderStyle[device].sliderWidth < 100">
                            <th>�����̴� ����</th>
                            <td>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].sliderAlign" value="left" type="radio"> ��</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].sliderAlign" value="" type="radio"> �߰�</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].sliderAlign" value="right" type="radio"> ��</label>
                            </td>
                        </tr>
                        <!--endregion-->
                    </tbody>
                </table>
            </div>
            <!--region ����,���-->
            <div class="popBtnWrapV19">
                <button class="btn4 btnWhite1">���</button>
                <button class="btn4 btnBlue1">����</button>
            </div>
            <!--endregion-->
            
            <!--region ��� ���� ���-->
            <MODAL ref="manageItemsModal" title="��� ����" :width="950">
                <MANAGE-ITEMS slot="body" :masterIndex="masterIndex" :items="items"
                    @postItem="openPostItemModal" @modifyItem="openModifyItemModal" @deleteItem="getItems(true)"
                    @saveSortAndSelected="getItems"/>
            </MODAL>
            <!--endregion-->
            
            <!--region ������ ��� ���-->
            <MODAL ref="postItemModal" :title="postItemModalTitle" @closeModal="closePostItemModal">
                <POST-ITEM slot="body" :masterIndex="masterIndex" :item="postItem"
                    @postItem="successPostItem" @cancel="cancelPostItem"/>
            </MODAL>
            <!--endregion-->
        </div>
    `,
    created() {
        this.masterIndex = Number(masterIndex);
        this.device = device;
        this.getItems();
    },
    mounted() {
        $(document).ready(this.createSwiper);
    },
    data() {return {
        swiper : null, // Swiper
        masterIndex : null, // ������ ������ �Ϸù�ȣ
        device : 'M', // ���� ä��(��ġ)

        postItemModalTitle : '������ ���', // ������ ���/���� ��� Ÿ��Ʋ
        items : [], // ������ ����Ʈ
        postItem : null, // ���� �� ������

        sliderStyle : {
            //region �����̴� ��Ÿ�� M - Mobile
            'M' : {
                //region ���
                backType : 'color', // ��� ����
                backgroundColor : 'fff', // ��� ��
                backImage : '', // ��� �̹���
                previewBackImage : '', // �̸����� ��� �̹���
                previewBackColor : 'fff', // �̸����� ��� �� �ڵ�
                //endregion

                showArrow : true, // ȭ��ǥ ���� ����

                //region �۾�
                fontSize : 15, // �۾� ũ��
                selectedFontColor : '000', // ���þȵ� ��Ʈ �� �ڵ�
                unSelectedFontColor : 'c3c3c3', // ���þȵ� ��Ʈ �� �ڵ�
                selectedFontEffect : '', // ���� �۾� ȿ��
                //endregion

                //region �����̴�
                sliderWidth : 100, // �����̴� ����
                sliderAlign : 'left', // �����̴� ���ı���
                sliderSpace : 10, // �����̴� ����
                sliderHeight : 50, // �����̴� ����
                //endregion
            },
            //endregion
            //region �����̴� ��Ÿ�� W - PCWeb
            'W' : {
                //region ���
                backType : 'color', // ��� ����
                backgroundColor : 'fff', // ��� ��
                backImage : '', // ��� �̹���
                previewBackImage : '', // �̸����� ��� �̹���
                previewBackColor : 'fff', // �̸����� ��� �� �ڵ�
                //endregion

                showArrow : true, // ȭ��ǥ ���� ����

                //region �۾�
                fontSize : 15, // �۾� ũ��
                selectedFontColor : '000', // ���þȵ� ��Ʈ �� �ڵ�
                unSelectedFontColor : 'c3c3c3', // ���þȵ� ��Ʈ �� �ڵ�
                selectedFontEffect : '', // ���� �۾� ȿ��
                //endregion

                //region �����̴�
                sliderWidth : 100, // �����̴� ����
                sliderAlign : 'left', // �����̴� ���ı���
                sliderSpace : 10, // �����̴� ����
                sliderHeight : 50, // �����̴� ����
                //endregion
            }
            //endregion
        },
    }},
    computed : {
        //region sliderAreaStyle �����̴� ���� ��Ÿ��
        sliderAreaStyle() {
            if( this.sliderStyle[this.device].backType === 'color' ) {
                return {
                    'background-color' : '#' + this.sliderStyle[this.device].backgroundColor,
                }
            } else {
                return {
                    'background-image' : 'url(' + this.sliderStyle[this.device].backImage + ')',
                    'background-size' : 'cover'
                }
            }
        },
        //endregion
        //region swiperContainerStyle swiper-container ��Ÿ��
        swiperContainerStyle() {
            return {
                'padding' : this.sliderStyle[this.device].showArrow ? '0 37px' : '',
                'width' : this.sliderStyle[this.device].sliderWidth + '%',
                'float' : this.sliderStyle[this.device].sliderAlign,
            }
        },
        //endregion
        //region previewEtcAreaStyle �̸����� Etc���� ��Ÿ��
        previewEtcAreaStyle() {
            return { 'background-color' : '#' + this.sliderStyle[this.device].previewBackColor };
        },
        //endregion
    },
    methods : {
        //region getItems ������ ����Ʈ ��ȸ
        getItems(flag) {
            const _this = this;
            this.callApi(2, 'GET', `/event/contents/${this.masterIndex}/tabbar/items`, null,
                data => {
                    _this.items = data;
                    if( flag )
                        setTimeout(this.refreshSwiper, 500);
                });
        },
        //endregion
        //region createSwiper Swiper ����
        createSwiper() {
            this.swiper = new Swiper('.preview .swiper-container',{
                initialSlide:0,
                slidesPerView:'auto',
                speed:300,
                prevButton:'.preview .btnPrev',
                nextButton:'.preview .btnNext'
            });
        },
        //endregion
        //region clickBackImageButton ��� �̹��� ��� ��ư Ŭ��
        clickBackImageButton() {
            document.getElementById('backFile').click();
        },
        //endregion
        //region changeBackFile ��� �̹��� ����
        changeBackFile(e) {
            const file = e.target.files[0];
            if( !file ) {
                this.sliderStyle[this.device].backImage = '';
                return false;
            }

            const _this = this;
            const imgData = this.createUploadImageData();
            this.callAjaxUploadImage(imgData, data => {
                const response = JSON.parse(data);

                if (response.response === 'ok') {
                    _this.sliderStyle[_this.device].backImage = response.filePath;
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
            imgData.append('image', document.getElementById("backFile").files[0]);
            return imgData;
        },
        //endregion
        //region clickBackImageButton ��� �̹��� ��� ��ư Ŭ��
        clickPreviewBackImageButton() {
            document.getElementById('previewBackFile').click();
        },
        //endregion
        //region changePreviewBackFile �̸����� ��� �̹��� ����
        changePreviewBackFile(e) {
            const file = e.target.files[0];
            if( !file ) {
                this.sliderStyle[this.device].previewBackImage = '';
                return false;
            }

            if (!file.type.match("image.*")) {
                alert("�̹��� ���ϸ� ����Ͻ� �� �ֽ��ϴ�.");
                return false;
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            const _this = this;
            reader.onload = function(e){
                _this.sliderStyle[_this.device].previewBackImage = e.target.result;
            }
        },
        //endregion
        //region upDownFontSize �۾�ũ�� +/-
        upDownFontSize(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].fontSize++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].fontSize--;
        },
        //endregion
        //region upDownSliderSpace �����̴����� +/-
        upDownSliderSpace(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].sliderSpace++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].sliderSpace--;
        },
        //endregion
        //region upDownSliderWidth �����̴����� +/-
        upDownSliderWidth(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].sliderWidth++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].sliderWidth--;
        },
        //endregion
        //region upDownSliderHeight �����̴����� +/-
        upDownSliderHeight(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].sliderHeight++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].sliderHeight--;
        },
        //endregion
        //region refreshSwiper Swiper ����
        refreshSwiper() {
            if( this.swiper ) {
                this.swiper.destroy();
                this.swiper = null;
                setTimeout(this.createSwiper, 500);
            }
        },
        //endregion
        //region openManageItemsModal ������ ���� ��� ����
        openManageItemsModal() {
            this.$refs.manageItemsModal.openModal();
        },
        //endregion
        //region openModifyItemModal ������ ��� ��� ����
        openPostItemModal() {
            this.$refs.manageItemsModal.closeModal();
            this.postItemModalTitle = '������ ���';
            this.postItem = null;
            this.$refs.postItemModal.openModal();
        },
        //endregion
        //region openModifyItemModal ������ ���� ��� ����
        openModifyItemModal(item) {
            this.$refs.manageItemsModal.closeModal();
            this.postItemModalTitle = '������ ����';
            this.postItem = item;
            this.$refs.postItemModal.openModal();
        },
        //endregion
        //region closePostItemModal ������ ���/���� ��� �ݱ�
        closePostItemModal() {
            this.postItemModalTitle = '������ ���';
            this.postItem = null;
        },
        //endregion
        //region slideStyle swiper-slide ��Ÿ��
        slideStyle(selected) {
            return {
                'padding' : '0 ' + this.sliderStyle[this.device].sliderSpace + 'px',
                'line-height' : this.sliderStyle[this.device].sliderHeight + 'px',
                'font-size' : this.sliderStyle[this.device].fontSize + 'px',
                'color' : '#' + (selected ? this.sliderStyle[this.device].selectedFontColor : this.sliderStyle[this.device].unSelectedFontColor),
                'font-weight' : selected ? 'bold' : ''
            }
        },
        //endregion
        //region successPostItem ������ ���/���� ����
        successPostItem() {
            this.$refs.postItemModal.closeModal();
            this.getItems(true);
            this.$refs.manageItemsModal.openModal();
        },
        //endregion
        //region cancelPostItem ������ ���/���� ���
        cancelPostItem() {
            this.$refs.postItemModal.closeModal();
            this.postItem = null;
            this.$refs.manageItemsModal.openModal();
        },
        //endregion
    },
    watch : {
        device() {
            this.refreshSwiper();
        },
    }
});