Vue.component('ITEM-NAME-PREFIX-MANAGE-ITEM', {
    template : `
        <div class="result">
            <!--region ��ư ����-->
            <div class="result-btn-area">
                <div>
                    <button @click="clickAddItem" class="btn">��ǰ �߰�</button>
                    <button @click="deleteSelectedDetail" class="btn">���� ����</button>
                </div>
                <div class="result-state-check">
                    <span>
                        <input v-model="checkedStates" value="T" type="checkbox" id="viewPrevSave">
                        <label for="viewPrevSave" class="blue">��ϴ��</label>
                    </span>
                    <span>
                        <input v-model="checkedStates" value="S" type="checkbox" id="viewSave">
                        <label for="viewSave" class="green">��ϵ�</label>
                    </span>
                    <span>
                        <input v-model="checkedStates" value="F" type="checkbox" id="viewFail">
                        <label for="viewFail" class="red">����</label>
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
                            <th>��ǰ�ڵ�</th>
                            <th>��ǰ��</th>
                            <th>�귣��ID</th>
                            <th>�ǸŻ���</th>
                            <th>�ǸŰ���</th>
                            <th>��ϰ��</th>
                            <th>�����</th>
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
                                <td><button @click="deleteDetail(detail)" class="btn">����</button></td>
                            </tr>
                        </template>
                        <tr v-else>
                            <td colspan="8">��ǰ�� �����ϴ�.</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!--region ���� ����-->
            <div v-if="failProducts.length > 0" class="modal-alert">
                <strong>�۾� ���</strong>
                <p>
                    �� {{requestPostProductCount}}�� ��� ��û �� 
                    {{requestPostProductCount - failPostProductCount}}�� ����, 
                    {{failPostProductCount}}�� ����
                </p>
                <strong>���� ����</strong>
                <p v-for="product in failProducts">
                    - ��ǰ�ڵ� : {{product.duplicatedProductId}}, ��ǰ�� : "{{product.duplicatedProductName}}" 
                    > ���Ӹ� [{{product.duplicatedPrefixWord}}]��(��) �̺�Ʈ�Ⱓ ��ħ
                </p>
                <a @click="resetFailProducts" class="close">x</a>
            </div>
            <!--endregion-->
            
            <div class="modal-btn-area">
                <button @click="postDetails" class="btn">����</button>
            </div>
        </div>
    `,
    mounted() {
        this.getDetails();
    },
    data() {return {
        details : [], // ���Ӹ� �� ����Ʈ

        checkedProductIds : [], // üũ�� ��ǰ ID ����Ʈ
        checkedStates : [], // ���õ� ���� ��
        deletedProductIds : [], // ������ ��ǰ id ����Ʈ(���� ��ϵƾ��� �͸� ���� �ؽ���, ��ϴ��� ���� ����)

        requestPostProductCount : 0, // ��� ��û ��ǰ ��
        failPostProductCount : 0, // ��� ��û ��ǰ �� ���� ��
        failProducts : [], // ������ ��ǰ ����Ʈ
    }},
    computed : {
        //region searchDetails �˻��� �� ����Ʈ
        searchDetails() {
            if( this.checkedStates.length === 0 )
                return this.details;
            else
                return this.details.filter(d => this.checkedStates.indexOf(d.state) > -1);
        },
        //endregion
        //region detailsToPost ��� �� �� ����Ʈ
        detailsToPost() {
            return this.details.filter(d => d.state !== 'S');
        },
        //endregion
    },
    props : {
        prefixIdx : { type:Number, default:0 }, // ���Ӹ� �Ϸù�ȣ
    },
    methods : {
        //region getDetails �� ����Ʈ ��ȸ
        getDetails() {
            const url = `/search/prefix/${this.prefixIdx}/details`;
            this.callApi(1, 'GET', url, null, this.successGetDetails);
        },
        successGetDetails(data) {
            data.forEach(d => d.state = 'S');
            this.details = data;
        },
        //endregion
        //region clickAddItem ��ǰ �߰� ��� ����
        clickAddItem() {
            this.$emit('clickAddItem');
        },
        //endregion
        //region addProducts ��ǰ �߰�
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
        //region numberFormat ���� õ�ڸ� (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
        //region stateTdClass ���� TD Ŭ���� ����Ʈ
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
        //region stateName ���¸�
        stateName(state) {
            switch (state) {
                case 'S' : return '��ϵ�';
                case 'F' : return '����';
                default : return '��ϴ��';
            }
        },
        //endregion
        //region checkAll ��ü üũ�ڽ� üũ/����
        checkAll(e) {
            if( e.target.checked ) {
                this.checkedProductIds = this.details.map(p => p.productId);
            } else {
                this.checkedProductIds = [];
            }
        },
        //endregion
        //region deleteSelectedDetail ���õ� �׸� ����
        deleteSelectedDetail() {
            this.details.filter(d => this.checkedProductIds.indexOf(d.productId) > -1 && d.state === 'S')
                .forEach(d => this.deletedProductIds.push(d.productId));

            this.details = this.details.filter(d => this.checkedProductIds.indexOf(d.productId) === -1);
            this.checkedProductIds = [];
        },
        //endregion
        //region deleteDetail �׸� �Ѱ� ����
        deleteDetail(product) {
            if( product.state === 'S' )
                this.deletedProductIds.push(product.productId);

            const index = this.details.findIndex(d => d.productId === product.productId);
            this.details.splice(index, 1);
        },
        //endregion
        //region postDetails �� ����
        postDetails() {
            if( !confirm('���� �Ͻðڽ��ϱ�?') )
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
        //region resetFailProducts ��� ���� ��ǰ �ʱ�ȭ
        resetFailProducts() {
            this.failProducts = [];
            this.failPostProductCount = 0;
        },
        //endregion
    }
});