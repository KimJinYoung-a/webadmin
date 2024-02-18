const app = new Vue({
    el : '#app',
    mixins : [api_mixin],
    template : /*html*/`
        <div class="popV19">
            <div class="popHeadV19">
                <h1>탭바</h1>
            </div>
            <div class="popContV19">
                <!--region 상단 탭-->
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
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-prev">이전</button>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-next">다음</button>
                                    </div>
                                </div>
                                <div class="etcArea" :style="previewEtcAreaStyle">
                                    <img v-if="sliderStyle[device].previewBackImage" :src="sliderStyle[device].previewBackImage"/>
                                </div>
                                <div class="inputArea">
                                    <input v-model="sliderStyle[device].previewBackColor" type="text" placeholder="배경 색 코드(#제외)">
                                    <button v-if="!sliderStyle[device].previewBackImage" @click="clickPreviewBackImageButton" class="btn4 btnBlue1">미리보기 이미지 등록</button>
                                    <button v-else @click="sliderStyle[device].previewBackImage = ''" class="btn4 btnBlue1">미리보기 이미지 삭제</button>
                                    <input type="file" id="previewBackFile" @change="changePreviewBackFile" style="display: none;">
                                </div>
                            </td>
                        </tr>
                        
                        <!--region 내용, 미리보기-->
                        <tr>
                            <th>목록</th>
                            <td><button @click="openManageItemsModal" class="btn4 btnBlue1">목록 관리</button></td>
                            <!--region 모바일&앱 미리보기-->
                            <td v-if="device === 'M'" class="preview mobile" rowspan="0" align="center">
                                <div class="sliderArea" :style="sliderAreaStyle">
                                    <div class="swiper-container" :style="swiperContainerStyle">
                                        <ul class="swiper-wrapper">
                                            <li v-for="item in items" class="swiper-slide" :style="slideStyle(item.selected)">
                                                <span>{{item.title}}</span>
                                            </li>
                                        </ul>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-prev">이전</button>
                                        <button v-show="sliderStyle[device].showArrow" type="button" class="btn-next">다음</button>
                                    </div>
                                </div>
                                <div class="etcArea" :style="previewEtcAreaStyle">
                                    <img :src="sliderStyle[device].previewBackImage"/>
                                </div>
                                <div class="inputArea">
                                    <input v-model="sliderStyle[device].previewBackColor" type="text" placeholder="배경 색 코드(#제외)">
                                    <button v-if="!sliderStyle[device].previewBackImage" @click="clickPreviewBackImageButton" class="btn4 btnBlue1">미리보기 이미지 등록</button>
                                    <button v-else @click="sliderStyle[device].previewBackImage = ''" class="btn4 btnBlue1">미리보기 이미지 삭제</button>
                                    <input type="file" id="previewBackFile" @change="changePreviewBackFile" style="display: none;">
                                </div>
                            </td>
                            <!--endregion-->
                        </tr>
                        <!--endregion-->
                        <!--region 배경-->
                        <tr>
                            <th>배경</th>
                            <td>
                                <p class="backTypeArea">
                                    <label class="chedkLabel"><input v-model="sliderStyle[device].backType" value="color" type="radio"> 색코드</label>
                                    <label class="chedkLabel"><input v-model="sliderStyle[device].backType" value="image" type="radio"> 이미지</label>
                                </p>
                                <input v-show="sliderStyle[device].backType === 'color'" v-model="sliderStyle[device].backgroundColor" type="text" class="short" placeholder="#제외">
                                <p v-show="sliderStyle[device].backType !== 'color'">
                                    <button @click="clickBackImageButton" class="btn4 btnBlue1">이미지 등록</button>
                                    <input @change="changeBackFile" id="backFile" type="file" style="display: none;">
                                    <img v-if="sliderStyle[device].backImage" :src="sliderStyle[device].backImage" class="back-image"/>
                                </p>
                            </td>
                        </tr>
                        <!--endregion-->
                        <!--region 글씨-->
                        <tr>
                            <th>글씨 색 코드</th>
                            <td>
                                <p class="settingArea">
                                    <strong>선택</strong>
                                    <input v-model="sliderStyle[device].selectedFontColor" type="text" class="short" placeholder="#제외">
                                </p>
                                <p class="settingArea">
                                    <strong>비선택</strong>
                                    <input v-model="sliderStyle[device].unSelectedFontColor" type="text" class="short" placeholder="#제외">
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <th>글씨 크기</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].fontSize" @keydown="upDownFontSize"> px</td>
                        </tr>
                        <tr>
                            <th>선택 글씨 효과</th>
                            <td>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].selectedFontEffect" value="" type="radio"> 없음</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].selectedFontEffect" value="reddot" type="radio"> 우상단 빨간점</label>
                            </td>
                        </tr>
                        <!--endregion-->
                        <!--region 화살표-->
                        <tr>
                            <th>화살표 노출</th>
                            <td>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].showArrow" :value="true" type="radio"> Y</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].showArrow" :value="false" type="radio"> N</label>
                            </td>
                        </tr>
                        <tr v-if="sliderStyle[device].showArrow">
                            <th>화살표 색 코드</th>
                            <td><input type="text" class="short" placeholder="#제외"></td>
                        </tr>
                        <!--endregion-->
                        <!--region 슬라이더-->
                        <tr>
                            <th>슬라이더 높이</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].sliderHeight" @keydown="upDownSliderHeight"> px</td>
                        </tr>
                        <tr>
                            <th>슬라이더 길이</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].sliderWidth" @keydown="upDownSliderWidth" @blur="refreshSwiper"> %</td>
                        </tr>
                        <tr>
                            <th>슬라이더 간격</th>
                            <td><input type="text" class="shortest" v-model="sliderStyle[device].sliderSpace" @keydown="upDownSliderSpace" @blur="refreshSwiper"> px</td>
                        </tr>
                        <tr v-show="sliderStyle[device].sliderWidth < 100">
                            <th>슬라이더 정렬</th>
                            <td>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].sliderAlign" value="left" type="radio"> 좌</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].sliderAlign" value="" type="radio"> 중간</label>
                                <label class="chedkLabel"><input v-model="sliderStyle[device].sliderAlign" value="right" type="radio"> 우</label>
                            </td>
                        </tr>
                        <!--endregion-->
                    </tbody>
                </table>
            </div>
            <!--region 저장,취소-->
            <div class="popBtnWrapV19">
                <button class="btn4 btnWhite1">취소</button>
                <button class="btn4 btnBlue1">저장</button>
            </div>
            <!--endregion-->
            
            <!--region 목록 관리 모달-->
            <MODAL ref="manageItemsModal" title="목록 관리" :width="950">
                <MANAGE-ITEMS slot="body" :masterIndex="masterIndex" :items="items"
                    @postItem="openPostItemModal" @modifyItem="openModifyItemModal" @deleteItem="getItems(true)"
                    @saveSortAndSelected="getItems"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 아이템 등록 모달-->
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
        masterIndex : null, // 컨텐츠 마스터 일련번호
        device : 'M', // 설정 채널(장치)

        postItemModalTitle : '아이템 등록', // 아이템 등록/수정 모달 타이틀
        items : [], // 아이템 리스트
        postItem : null, // 수정 중 아이템

        sliderStyle : {
            //region 슬라이더 스타일 M - Mobile
            'M' : {
                //region 배경
                backType : 'color', // 배경 유형
                backgroundColor : 'fff', // 배경 색
                backImage : '', // 배경 이미지
                previewBackImage : '', // 미리보기 배경 이미지
                previewBackColor : 'fff', // 미리보기 배경 색 코드
                //endregion

                showArrow : true, // 화살표 노출 여부

                //region 글씨
                fontSize : 15, // 글씨 크기
                selectedFontColor : '000', // 선택안된 폰트 색 코드
                unSelectedFontColor : 'c3c3c3', // 선택안된 폰트 색 코드
                selectedFontEffect : '', // 선택 글씨 효과
                //endregion

                //region 슬라이더
                sliderWidth : 100, // 슬라이더 길이
                sliderAlign : 'left', // 슬라이더 정렬기준
                sliderSpace : 10, // 슬라이더 간격
                sliderHeight : 50, // 슬라이더 높이
                //endregion
            },
            //endregion
            //region 슬라이더 스타일 W - PCWeb
            'W' : {
                //region 배경
                backType : 'color', // 배경 유형
                backgroundColor : 'fff', // 배경 색
                backImage : '', // 배경 이미지
                previewBackImage : '', // 미리보기 배경 이미지
                previewBackColor : 'fff', // 미리보기 배경 색 코드
                //endregion

                showArrow : true, // 화살표 노출 여부

                //region 글씨
                fontSize : 15, // 글씨 크기
                selectedFontColor : '000', // 선택안된 폰트 색 코드
                unSelectedFontColor : 'c3c3c3', // 선택안된 폰트 색 코드
                selectedFontEffect : '', // 선택 글씨 효과
                //endregion

                //region 슬라이더
                sliderWidth : 100, // 슬라이더 길이
                sliderAlign : 'left', // 슬라이더 정렬기준
                sliderSpace : 10, // 슬라이더 간격
                sliderHeight : 50, // 슬라이더 높이
                //endregion
            }
            //endregion
        },
    }},
    computed : {
        //region sliderAreaStyle 슬라이더 영역 스타일
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
        //region swiperContainerStyle swiper-container 스타일
        swiperContainerStyle() {
            return {
                'padding' : this.sliderStyle[this.device].showArrow ? '0 37px' : '',
                'width' : this.sliderStyle[this.device].sliderWidth + '%',
                'float' : this.sliderStyle[this.device].sliderAlign,
            }
        },
        //endregion
        //region previewEtcAreaStyle 미리보기 Etc영역 스타일
        previewEtcAreaStyle() {
            return { 'background-color' : '#' + this.sliderStyle[this.device].previewBackColor };
        },
        //endregion
    },
    methods : {
        //region getItems 아이템 리스트 조회
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
        //region createSwiper Swiper 생성
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
        //region clickBackImageButton 배경 이미지 등록 버튼 클릭
        clickBackImageButton() {
            document.getElementById('backFile').click();
        },
        //endregion
        //region changeBackFile 배경 이미지 변경
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
            imgData.append('image', document.getElementById("backFile").files[0]);
            return imgData;
        },
        //endregion
        //region clickBackImageButton 배경 이미지 등록 버튼 클릭
        clickPreviewBackImageButton() {
            document.getElementById('previewBackFile').click();
        },
        //endregion
        //region changePreviewBackFile 미리보기 배경 이미지 변경
        changePreviewBackFile(e) {
            const file = e.target.files[0];
            if( !file ) {
                this.sliderStyle[this.device].previewBackImage = '';
                return false;
            }

            if (!file.type.match("image.*")) {
                alert("이미지 파일만 등록하실 수 있습니다.");
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
        //region upDownFontSize 글씨크기 +/-
        upDownFontSize(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].fontSize++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].fontSize--;
        },
        //endregion
        //region upDownSliderSpace 슬라이더간격 +/-
        upDownSliderSpace(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].sliderSpace++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].sliderSpace--;
        },
        //endregion
        //region upDownSliderWidth 슬라이더길이 +/-
        upDownSliderWidth(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].sliderWidth++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].sliderWidth--;
        },
        //endregion
        //region upDownSliderHeight 슬라이더높이 +/-
        upDownSliderHeight(e) {
            if( e.keyCode === 38 )
                this.sliderStyle[this.device].sliderHeight++;
            else if( e.keyCode === 40 )
                this.sliderStyle[this.device].sliderHeight--;
        },
        //endregion
        //region refreshSwiper Swiper 갱신
        refreshSwiper() {
            if( this.swiper ) {
                this.swiper.destroy();
                this.swiper = null;
                setTimeout(this.createSwiper, 500);
            }
        },
        //endregion
        //region openManageItemsModal 아이템 관리 모달 열기
        openManageItemsModal() {
            this.$refs.manageItemsModal.openModal();
        },
        //endregion
        //region openModifyItemModal 아이템 등록 모달 열기
        openPostItemModal() {
            this.$refs.manageItemsModal.closeModal();
            this.postItemModalTitle = '아이템 등록';
            this.postItem = null;
            this.$refs.postItemModal.openModal();
        },
        //endregion
        //region openModifyItemModal 아이템 수정 모달 열기
        openModifyItemModal(item) {
            this.$refs.manageItemsModal.closeModal();
            this.postItemModalTitle = '아이템 수정';
            this.postItem = item;
            this.$refs.postItemModal.openModal();
        },
        //endregion
        //region closePostItemModal 아이템 등록/수정 모달 닫기
        closePostItemModal() {
            this.postItemModalTitle = '아이템 등록';
            this.postItem = null;
        },
        //endregion
        //region slideStyle swiper-slide 스타일
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
        //region successPostItem 아이템 등록/수정 성공
        successPostItem() {
            this.$refs.postItemModal.closeModal();
            this.getItems(true);
            this.$refs.manageItemsModal.openModal();
        },
        //endregion
        //region cancelPostItem 아이템 등록/수정 취소
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