Vue.component('ITEM-NAME-PREFIX-RESULT', {
    template : `
        <div class="result">
            <div class="result-btn-area">
                <button @click="$emit('postPrefix')" class="btn">�ű� ����</button>
            </div>
            
            <div class="result-list">
                <table>
                    <!--region colgroup-->
                    <colgroup>
                        <col style="width: 80px;">
                        <col style="width: 150px;">
                        <col style="width: auto;">
                        <col style="width: auto;">
                        <col style="width: 80px;">
                        <col style="width: 80px;">
                        <col style="width: 150px;">
                        <col style="width: 100px;">
                        <col style="width: 100px;">
                    </colgroup>
                    <!--endregion-->
                    <!--region thead-->
                    <thead>
                        <tr>
                            <th>idx</th>
                            <th>���Ӹ�</th>
                            <th>�����Ͻ�</th>
                            <th>�����Ͻ�</th>
                            <th>����</th>
                            <th>��ǰ��</th>
                            <th>����Ͻ�</th>
                            <th>�����</th>
                            <th>��</th>
                        </tr>
                    </thead>
                    <!--endregion-->
                    <tbody v-if="prefixes.length > 0">
                        <ITEM-NAME-PREFIX-RESULT-ITEM v-for="prefix in prefixes"
                            @updatePrefix="updatePrefix" @manageProduct="manageProduct"
                            :key="prefix.prefixIdx" :prefix="prefix"/>
                    </tbody>
                    <tbody v-else>
                        <tr><td colspan="9">��ϵ� ���Ӹ��� �����ϴ�.</td></tr>
                    </tbody>
                </table>
            </div>
            
            <PAGINATION :currentPage="currentPage" :lastPage="lastPage" @clickPage="clickPage"/>
        </div>
    `,
    data() {return {
        currentPage : 1, // ���� ������

    }},
    props : {
        prefixes : { type:Array, default:function() {return []} }, // ���Ӹ� ����Ʈ
        lastPage : { type:Number, default:0 },
    },
    methods : {
        updatePrefix(prefix) {
            this.$emit('updatePrefix', prefix);
        },
        manageProduct(prefixIdx) {
            this.$emit('manageProduct', prefixIdx);
        },
        clickPage(page) {
            this.$emit('clickPage', page);
        },
    }
});