Vue.component('ITEM-NAME-PREFIX-SEARCH-ITEM', {
    template : `
        <div class="result">
        
            <!--region 검색-->
            <div class="search">
                <div class="search-group">
                    <label>상품ID</label>
                    <textarea v-model="searchProductIdValue" rows="3"></textarea>
                </div>
                <div class="search-group">
                    <label>브랜드ID</label>
                    <textarea v-model="searchBrandIdValue" rows="3"></textarea>
                </div>
                <button @click="search" class="btn">검색</button>
            </div>
            <!--endregion-->
            
            <div class="result-list" style="max-height: 400px;overflow: scroll;">
                <table>
                    <!--region colgroup-->
                    <colgroup>
                        <col style="width: 50px;">
                        <col style="width: 100px;">
                        <col style="width: auto;">
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
                                    :checked="products.length > 0 && checkedProductIds.length === products.length">
                            </th>
                            <th>상품코드</th>
                            <th>상품명</th>
                            <th>브랜드ID</th>
                            <th>판매상태</th>
                            <th>판매가격</th>
                        </tr>
                    </thead>
                    <!--endregion-->
                    <tbody>
                        <template v-if="products.length > 0">
                            <tr v-for="product in products">
                                <td>
                                    <input type="checkbox" :value="product.productId" v-model="checkedProductIds">
                                </td>
                                <td>{{product.productId}}</td>
                                <td>{{product.productName}}</td>
                                <td>{{product.brandId}}</td>
                                <td>{{product.use ? 'Y' : 'N'}}</td>
                                <td>{{product.salesPrice}}</td>
                            </tr>
                        </template>
                        <tr v-else>
                            <td colspan="6">상품이 없습니다.</td><
                        </tr>
                    </tbody>
                </table>
                
                <PAGINATION :currentPage="currentPage" :lastPage="lastPage" @clickPage="goPage"/>
            </div>
            
            <div class="modal-btn-area">
                <button @click="addProducts" class="btn">등록</button>
            </div>
        </div>
    `,
    mounted() {
        this.getProducts();
    },
    data() {return {
        currentPage : 1,
        lastPage : 1,
        products : [], // 상품 리스트

        checkedProductIds : [], // 체크한 상품 ID 리스트
        searchProductIdValue : '', // 상품 ID 검색 textarea 값
        searchBrandIdValue : '', // 브랜드 ID 검색 textarea 값
    }},
    computed : {
        //region productIds 상품ID 리스트
        productIds() {
            const value = this.searchProductIdValue.trim();
            if( value === '' )
                return [];
            else
                return value.replace(/ /g, '')
                    .replace(/\n/g, ',')
                    .split(',');
        },
        //endregion
        //region brandIds 브랜드ID 리스트
        brandIds() {
            const value = this.searchBrandIdValue.trim();
            if( value === '' )
                return [];
            else
                return value.replace(/ /g, '')
                    .replace(/\n/g, ',')
                    .split(',');
        },
        //endregion
    },
    props : {
        prefixIdx : { type:Number, default:0 }, // 말머리 일련번호
    },
    methods : {
        //region getProducts 상품 리스트 조회
        getProducts() {
            const url = '/search/prefix/products/search';
            const data = {
                prefixIdx : this.prefixIdx,
                productIds : this.productIds.join(','),
                brandIds : this.brandIds.join(','),
                page : this.currentPage
            };
            this.callApi(1, 'GET', url, data, this.successGetProducts);
        },
        successGetProducts(data) {
            this.lastPage = data.lastPage;
            this.products = data.products;
            const area = this.$el.querySelector('.result-list');
            $(area).animate({
                scrollTop : 0
            }, 200);
        },
        //endregion
        //region search 검색
        search() {
            this.currentPage = 1;
            this.getProducts();
        },
        //endregion
        //region goPage 페이지 이동
        goPage(page) {
            this.currentPage = page;
            this.getProducts();
        },
        //endregion
        //region checkAll 전체 체크박스 체크/해제
        checkAll(e) {
            if( e.target.checked ) {
                this.checkedProductIds = this.products.map(p => p.productId);
            } else {
                this.checkedProductIds = [];
            }
        },
        //endregion
        //region addProducts 상품 등록
        addProducts() {
            if( this.checkedProductIds.length === 0 ) {
                alert('상품을 체크 해 주세요');
                return false;
            } else {
                const products = this.products.filter(p => this.checkedProductIds.indexOf(p.productId) >= 0);
                this.$emit('addProducts', products);
            }
        },
        //endregion
    }
});