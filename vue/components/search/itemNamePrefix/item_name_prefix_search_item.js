Vue.component('ITEM-NAME-PREFIX-SEARCH-ITEM', {
    template : `
        <div class="result">
        
            <!--region �˻�-->
            <div class="search">
                <div class="search-group">
                    <label>��ǰID</label>
                    <textarea v-model="searchProductIdValue" rows="3"></textarea>
                </div>
                <div class="search-group">
                    <label>�귣��ID</label>
                    <textarea v-model="searchBrandIdValue" rows="3"></textarea>
                </div>
                <button @click="search" class="btn">�˻�</button>
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
                            <th>��ǰ�ڵ�</th>
                            <th>��ǰ��</th>
                            <th>�귣��ID</th>
                            <th>�ǸŻ���</th>
                            <th>�ǸŰ���</th>
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
                            <td colspan="6">��ǰ�� �����ϴ�.</td><
                        </tr>
                    </tbody>
                </table>
                
                <PAGINATION :currentPage="currentPage" :lastPage="lastPage" @clickPage="goPage"/>
            </div>
            
            <div class="modal-btn-area">
                <button @click="addProducts" class="btn">���</button>
            </div>
        </div>
    `,
    mounted() {
        this.getProducts();
    },
    data() {return {
        currentPage : 1,
        lastPage : 1,
        products : [], // ��ǰ ����Ʈ

        checkedProductIds : [], // üũ�� ��ǰ ID ����Ʈ
        searchProductIdValue : '', // ��ǰ ID �˻� textarea ��
        searchBrandIdValue : '', // �귣�� ID �˻� textarea ��
    }},
    computed : {
        //region productIds ��ǰID ����Ʈ
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
        //region brandIds �귣��ID ����Ʈ
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
        prefixIdx : { type:Number, default:0 }, // ���Ӹ� �Ϸù�ȣ
    },
    methods : {
        //region getProducts ��ǰ ����Ʈ ��ȸ
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
        //region search �˻�
        search() {
            this.currentPage = 1;
            this.getProducts();
        },
        //endregion
        //region goPage ������ �̵�
        goPage(page) {
            this.currentPage = page;
            this.getProducts();
        },
        //endregion
        //region checkAll ��ü üũ�ڽ� üũ/����
        checkAll(e) {
            if( e.target.checked ) {
                this.checkedProductIds = this.products.map(p => p.productId);
            } else {
                this.checkedProductIds = [];
            }
        },
        //endregion
        //region addProducts ��ǰ ���
        addProducts() {
            if( this.checkedProductIds.length === 0 ) {
                alert('��ǰ�� üũ �� �ּ���');
                return false;
            } else {
                const products = this.products.filter(p => this.checkedProductIds.indexOf(p.productId) >= 0);
                this.$emit('addProducts', products);
            }
        },
        //endregion
    }
});