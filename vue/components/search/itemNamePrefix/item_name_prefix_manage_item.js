Vue.component('ITEM-NAME-PREFIX-MANAGE-ITEM', {
    template : `
        <div class="result">
            <!--region 버튼 영역-->
            <div class="result-btn-area">
                <div>
                    <button @click="clickAddItem" class="btn">상품 추가</button>
                    <button @click="deleteSelectedDetail" class="btn">선택 삭제</button>
                </div>
                <div class="result-state-check">
                    <span>
                        <input v-model="checkedStates" value="T" type="checkbox" id="viewPrevSave">
                        <label for="viewPrevSave" class="blue">등록대기</label>
                    </span>
                    <span>
                        <input v-model="checkedStates" value="S" type="checkbox" id="viewSave">
                        <label for="viewSave" class="green">등록됨</label>
                    </span>
                    <span>
                        <input v-model="checkedStates" value="F" type="checkbox" id="viewFail">
                        <label for="viewFail" class="red">실패</label>
                    </span>
                </div>
            </div>
            <!--endregion-->
            
            <div class="result-list" style="max-height: 350px;overflow: scroll;">
                <table>
                    <!--region colgroup-->
                    <colgroup>
                        <col style="width: 50px;">
                        <col style="width: 100px;">
                        <col style="width: auto;">
                        <col style="width: 100px;">
                        <col style="width: 80px;">
                        <col style="width: 100px;">
                        <col style="width: 80px;">
                        <col style="width: 100px;">
                    </colgroup>
                    <!--endregion-->
                    <!--region thead-->
                    <thead>
                        <tr>
                            <th>
                                <input type="checkbox" @click="checkAll" 
                                    :checked="details.length > 0 && checkedProductIds.length === details.length">
                            </th>
                            <th>상품코드</th>
                            <th>상품명</th>
                            <th>브랜드ID</th>
                            <th>판매상태</th>
                            <th>판매가격</th>
                            <th>등록결과</th>
                            <th>행삭제</th>
                        </tr>
                    </thead>
                    <!--endregion-->
                    <tbody>
                        <template v-if="searchDetails.length > 0">
                            <tr v-for="detail in searchDetails">
                                <td><input v-model="checkedProductIds" :value="detail.productId" type="checkbox"></td>
                                <td>{{detail.productId}}</td>
                                <td>{{detail.productName}}</td>
                                <td>{{detail.brandId}}</td>
                                <td>{{detail.use ? 'Y' : 'N'}}</td>
                                <td>{{numberFormat(detail.salesPrice)}}</td>
                                <td>
                                    <p :class="stateTdClass(detail.state)">{{stateName(detail.state)}}</p>
                                </td>
                                <td><button @click="deleteDetail(detail)" class="btn">삭제</button></td>
                            </tr>
                        </template>
                        <tr v-else>
                            <td colspan="8">상품이 없습니다.</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!--region 실패 사유-->
            <div v-if="failProducts.length > 0" class="modal-alert">
                <strong>작업 요약</strong>
                <p>
                    총 {{requestPostProductCount}}건 등록 요청 중 
                    {{requestPostProductCount - failPostProductCount}}개 성공, 
                    {{failPostProductCount}}개 실패
                </p>
                <strong>실패 사유</strong>
                <p v-for="product in failProducts">
                    - 상품코드 : {{product.duplicatedProductId}}, 상품명 : "{{product.duplicatedProductName}}" 
                    > 말머리 [{{product.duplicatedPrefixWord}}]와(과) 이벤트기간 겹침
                </p>
                <a @click="resetFailProducts" class="close">x</a>
            </div>
            <!--endregion-->
            
            <div class="modal-btn-area">
                <button @click="postDetails" class="btn">저장</button>
            </div>
        </div>
    `,
    mounted() {
        this.getDetails();
    },
    data() {return {
        details : [], // 말머리 상세 리스트

        checkedProductIds : [], // 체크한 상품 ID 리스트
        checkedStates : [], // 선택된 상태 값
        deletedProductIds : [], // 삭제된 상품 id 리스트(실제 등록됐었던 것만 담음 ※실패, 등록대기는 담지 않음)

        requestPostProductCount : 0, // 등록 요청 상품 수
        failPostProductCount : 0, // 등록 요청 상품 중 실패 수
        failProducts : [], // 실패한 상품 리스트
    }},
    computed : {
        //region searchDetails 검색된 상세 리스트
        searchDetails() {
            if( this.checkedStates.length === 0 )
                return this.details;
            else
                return this.details.filter(d => this.checkedStates.indexOf(d.state) > -1);
        },
        //endregion
        //region detailsToPost 등록 할 상세 리스트
        detailsToPost() {
            return this.details.filter(d => d.state !== 'S');
        },
        //endregion
    },
    props : {
        prefixIdx : { type:Number, default:0 }, // 말머리 일련번호
    },
    methods : {
        //region getDetails 상세 리스트 조회
        getDetails() {
            const url = `/search/prefix/${this.prefixIdx}/details`;
            this.callApi(1, 'GET', url, null, this.successGetDetails);
        },
        successGetDetails(data) {
            data.forEach(d => d.state = 'S');
            this.details = data;
        },
        //endregion
        //region clickAddItem 상품 추가 모달 열기
        clickAddItem() {
            this.$emit('clickAddItem');
        },
        //endregion
        //region addProducts 상품 추가
        addProducts(products) {
            if( products == null )
                return false;

            products.forEach(p => {
                this.details.push({
                    productId : p.productId,
                    productName : p.productName,
                    brandId : p.brandId,
                    use : p.use,
                    salesPrice : p.salesPrice,
                    state : 'T'
                });
            });
        }
        ,
        //endregion
        //region numberFormat 숫자 천자리 (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
        //region stateTdClass 상태 TD 클래스 리스트
        stateTdClass(state) {
            const classes = ['state'];
            switch (state) {
                case 'S' : classes.push('save'); break;
                case 'F' : classes.push('fail'); break;
                default : classes.push('staging'); break;
            }
            return classes;
        },
        //endregion
        //region stateName 상태명
        stateName(state) {
            switch (state) {
                case 'S' : return '등록됨';
                case 'F' : return '실패';
                default : return '등록대기';
            }
        },
        //endregion
        //region checkAll 전체 체크박스 체크/해제
        checkAll(e) {
            if( e.target.checked ) {
                this.checkedProductIds = this.details.map(p => p.productId);
            } else {
                this.checkedProductIds = [];
            }
        },
        //endregion
        //region deleteSelectedDetail 선택된 항목 삭제
        deleteSelectedDetail() {
            this.details.filter(d => this.checkedProductIds.indexOf(d.productId) > -1 && d.state === 'S')
                .forEach(d => this.deletedProductIds.push(d.productId));

            this.details = this.details.filter(d => this.checkedProductIds.indexOf(d.productId) === -1);
            this.checkedProductIds = [];
        },
        //endregion
        //region deleteDetail 항목 한건 삭제
        deleteDetail(product) {
            if( product.state === 'S' )
                this.deletedProductIds.push(product.productId);

            const index = this.details.findIndex(d => d.productId === product.productId);
            this.details.splice(index, 1);
        },
        //endregion
        //region postDetails 상세 저장
        postDetails() {
            if( !confirm('저장 하시겠습니까?') )
                return false;

            this.resetFailProducts();
            const postDetailProductIds = this.detailsToPost.map(p => p.productId);
            this.requestPostProductCount = postDetailProductIds.length;

            const url = '/search/prefix/details';
            const data = this.createPostDetailsApiData(postDetailProductIds);
            this.callApi(1, 'POST', url, data, this.successPostDetails);
        },
        createPostDetailsApiData(postDetailProductIds) {
            return {
                prefixIdx : this.prefixIdx,
                productIds : postDetailProductIds.join(','),
                deleteProductIds : this.deletedProductIds.length > 0 ? this.deletedProductIds.join(',') : []
            };
        },
        successPostDetails(data) {
            this.$emit('updateItemCount', this.prefixIdx, data.prefixItemCount);
            this.failProducts = data.failProducts;
            this.failPostProductCount = this.failProducts.length;

            if( this.failPostProductCount > 0 ) {
                this.showFailAndSuccessProducts();
            } else {
                this.detailsToPost.forEach(d => d.state = 'S');
            }

            this.sortDetailsByState();
        },
        showFailAndSuccessProducts() {
            const failProductIds = this.failProducts.map(p => p.duplicatedProductId);
            this.detailsToPost.forEach(d => {
                if( failProductIds.indexOf(d.productId) > -1 ) {
                    d.state = 'F';
                } else {
                    d.state = 'S';
                }
            });
        },
        sortDetailsByState() {
            this.details.sort((a, b) => {
                return a.state > b.state;
            });
        },
        //endregion
        //region resetFailProducts 등록 실패 상품 초기화
        resetFailProducts() {
            this.failProducts = [];
            this.failPostProductCount = 0;
        },
        //endregion
    }
});