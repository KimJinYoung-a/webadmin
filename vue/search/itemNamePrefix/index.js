const app = new Vue({
    el: '#app',
    store : store,
    mixins: [api_mixin],
    template: `
        <div class="container">
            <h3 class="title">말머리 관리</h3>
            
            <div class="content">
                
                <ITEM-NAME-PREFIX-SEARCH ref="search" @search="searchPrefixes"/>
                
                <ITEM-NAME-PREFIX-RESULT @postPrefix="openPostPrefixModal" @clickPage="clickPrefixPage"
                    @updatePrefix="openUpdatePrefixModal" @manageProduct="openManagePrefixItemModal" 
                    :prefixes="prefixes" :lastPage="prefixLastPage"/>
                
            </div>
            
            <MODAL ref="postPrefixModal" title="말머리 등록/수정" @closeModal="updatePrefix = null">
                <ITEM-NAME-PREFIX-POST slot="body" @postPrefix="callbackPostPrefix" :updatePrefix="updatePrefix"/>
            </MODAL>
            
            <MODAL ref="manageItemModal" title="상품 관리" :width="950">
                <ITEM-NAME-PREFIX-MANAGE-ITEM slot="body" ref="manageItem" :prefixIdx="searchProductPrefixIdx" 
                    @clickAddItem="$refs.searchItemModal.openModal()" @updateItemCount="updateItemCount"/>
            </MODAL>
            
            <MODAL ref="searchItemModal" title="상품 조회" :width="800">
                <ITEM-NAME-PREFIX-SEARCH-ITEM slot="body" :prefixIdx="searchProductPrefixIdx" 
                    @addProducts="addPrefixProducts"/>
            </MODAL>
        </div>
    `,
    created() {
        this.$store.commit('SET_APP', this);
        this.$store.dispatch('GET_PREFIXES');
    },
    data() {return {
        searchProductPrefixIdx : 0, // 상품 조회할 말머리 일련번호
        updatePrefix : null, // 수정중인 말머리
    }},
    computed : {
        prefixCount() { return this.$store.getters.prefixCount; },
        prefixCurrentPage() { return this.$store.getters.prefixCurrentPage; },
        prefixLastPage() { return this.$store.getters.prefixLastPage; },
        prefixes() { return this.$store.getters.prefixes; },
    },
    methods : {
        //region openPostPrefixModal 말머리 등록 모달 열기
        openPostPrefixModal() {
            this.$refs.postPrefixModal.openModal();
        },
        //endregion
        //region searchPrefixes 말머리 검색
        searchPrefixes(data) {
            console.log(data)
            this.$store.commit('SET_PREFIX_SEARCH', data);
            this.$store.commit('SET_PREFIX_CURRENT_PAGE', 1);
            this.$store.dispatch('GET_PREFIXES');
        },
        //endregion
        //region clickPrefixPage 말머리 페이지 클릭
        clickPrefixPage(page) {
            this.$store.commit('SET_PREFIX_CURRENT_PAGE', page);
            this.$store.dispatch('GET_PREFIXES');
        },
        //endregion
        //region openUpdatePrefixModal 말머리 수정 모달 열기
        openUpdatePrefixModal(prefix) {
            this.updatePrefix = prefix;
            this.$refs.postPrefixModal.openModal();
        },
        //endregion
        //region callbackPostPrefix 말머리 등록/수정 후 처리
        callbackPostPrefix() {
            this.$refs.postPrefixModal.closeModal();
            this.$store.dispatch('GET_PREFIXES');
        },
        //endregion

        //region openManagePrefixItemModal 상품 관리 모달 열기
        openManagePrefixItemModal(prefixIdx) {
            this.searchProductPrefixIdx = prefixIdx;
            this.$refs.manageItemModal.openModal();
        },
        //endregion
        //region addPrefixProducts 말머리 상품 추가
        addPrefixProducts(products) {
            this.$refs.manageItem.addProducts(products);
            this.$refs.searchItemModal.closeModal();
        },
        //endregion
        //region updateItemCount 말머리 상품 수 수정
        updateItemCount(prefixIdx, itemCount) {
            this.$store.commit('UPDATE_PREFIX_ITEM_COUNT', {prefixIdx, itemCount})
        },
        //endregion
    }
});