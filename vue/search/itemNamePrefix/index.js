const app = new Vue({
    el: '#app',
    store : store,
    mixins: [api_mixin],
    template: `
        <div class="container">
            <h3 class="title">���Ӹ� ����</h3>
            
            <div class="content">
                
                <ITEM-NAME-PREFIX-SEARCH ref="search" @search="searchPrefixes"/>
                
                <ITEM-NAME-PREFIX-RESULT @postPrefix="openPostPrefixModal" @clickPage="clickPrefixPage"
                    @updatePrefix="openUpdatePrefixModal" @manageProduct="openManagePrefixItemModal" 
                    :prefixes="prefixes" :lastPage="prefixLastPage"/>
                
            </div>
            
            <MODAL ref="postPrefixModal" title="���Ӹ� ���/����" @closeModal="updatePrefix = null">
                <ITEM-NAME-PREFIX-POST slot="body" @postPrefix="callbackPostPrefix" :updatePrefix="updatePrefix"/>
            </MODAL>
            
            <MODAL ref="manageItemModal" title="��ǰ ����" :width="950">
                <ITEM-NAME-PREFIX-MANAGE-ITEM slot="body" ref="manageItem" :prefixIdx="searchProductPrefixIdx" 
                    @clickAddItem="$refs.searchItemModal.openModal()" @updateItemCount="updateItemCount"/>
            </MODAL>
            
            <MODAL ref="searchItemModal" title="��ǰ ��ȸ" :width="800">
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
        searchProductPrefixIdx : 0, // ��ǰ ��ȸ�� ���Ӹ� �Ϸù�ȣ
        updatePrefix : null, // �������� ���Ӹ�
    }},
    computed : {
        prefixCount() { return this.$store.getters.prefixCount; },
        prefixCurrentPage() { return this.$store.getters.prefixCurrentPage; },
        prefixLastPage() { return this.$store.getters.prefixLastPage; },
        prefixes() { return this.$store.getters.prefixes; },
    },
    methods : {
        //region openPostPrefixModal ���Ӹ� ��� ��� ����
        openPostPrefixModal() {
            this.$refs.postPrefixModal.openModal();
        },
        //endregion
        //region searchPrefixes ���Ӹ� �˻�
        searchPrefixes(data) {
            console.log(data)
            this.$store.commit('SET_PREFIX_SEARCH', data);
            this.$store.commit('SET_PREFIX_CURRENT_PAGE', 1);
            this.$store.dispatch('GET_PREFIXES');
        },
        //endregion
        //region clickPrefixPage ���Ӹ� ������ Ŭ��
        clickPrefixPage(page) {
            this.$store.commit('SET_PREFIX_CURRENT_PAGE', page);
            this.$store.dispatch('GET_PREFIXES');
        },
        //endregion
        //region openUpdatePrefixModal ���Ӹ� ���� ��� ����
        openUpdatePrefixModal(prefix) {
            this.updatePrefix = prefix;
            this.$refs.postPrefixModal.openModal();
        },
        //endregion
        //region callbackPostPrefix ���Ӹ� ���/���� �� ó��
        callbackPostPrefix() {
            this.$refs.postPrefixModal.closeModal();
            this.$store.dispatch('GET_PREFIXES');
        },
        //endregion

        //region openManagePrefixItemModal ��ǰ ���� ��� ����
        openManagePrefixItemModal(prefixIdx) {
            this.searchProductPrefixIdx = prefixIdx;
            this.$refs.manageItemModal.openModal();
        },
        //endregion
        //region addPrefixProducts ���Ӹ� ��ǰ �߰�
        addPrefixProducts(products) {
            this.$refs.manageItem.addProducts(products);
            this.$refs.searchItemModal.closeModal();
        },
        //endregion
        //region updateItemCount ���Ӹ� ��ǰ �� ����
        updateItemCount(prefixIdx, itemCount) {
            this.$store.commit('UPDATE_PREFIX_ITEM_COUNT', {prefixIdx, itemCount})
        },
        //endregion
    }
});