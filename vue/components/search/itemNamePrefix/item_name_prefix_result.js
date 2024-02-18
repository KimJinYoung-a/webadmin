Vue.component('ITEM-NAME-PREFIX-RESULT', {
    template : `
        <div class="result">
            <div class="result-btn-area">
                <button @click="$emit('postPrefix')" class="btn">신규 생성</button>
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
                            <th>말머리</th>
                            <th>시작일시</th>
                            <th>종료일시</th>
                            <th>상태</th>
                            <th>상품수</th>
                            <th>등록일시</th>
                            <th>등록자</th>
                            <th>상세</th>
                        </tr>
                    </thead>
                    <!--endregion-->
                    <tbody v-if="prefixes.length > 0">
                        <ITEM-NAME-PREFIX-RESULT-ITEM v-for="prefix in prefixes"
                            @updatePrefix="updatePrefix" @manageProduct="manageProduct"
                            :key="prefix.prefixIdx" :prefix="prefix"/>
                    </tbody>
                    <tbody v-else>
                        <tr><td colspan="9">등록된 말머리가 없습니다.</td></tr>
                    </tbody>
                </table>
            </div>
            
            <PAGINATION :currentPage="currentPage" :lastPage="lastPage" @clickPage="clickPage"/>
        </div>
    `,
    data() {return {
        currentPage : 1, // 현재 페이지

    }},
    props : {
        prefixes : { type:Array, default:function() {return []} }, // 말머리 리스트
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