Vue.component('ITEM-NAME-PREFIX-SEARCH', {
    template : `
        <div class="search">
            <div class="search-group">
                <label>�Ⱓ</label>
                <DATE-PICKER @updateDate="updateStartDate" :date="startDate" id="searchStartDate"/>
                <DATE-PICKER @updateDate="updateEndDate" :date="endDate" id="searchEndDate"/>
            </div>
            <div class="search-group">
                <label>���Ӹ�</label>
                <input v-model="keyword" type="text">
            </div>
            <button @click="search" class="btn search-btn">�˻�</button>
            <button @click="reset" class="btn search-btn">�ʱ�ȭ</button>
        </div>
    `,
    data() {return {
        startDate : '',
        endDate : '',
        keyword : '',
    }},
    methods : {
        //region updateStartDate ������ ����
        updateStartDate(date) {
            this.startDate = date;
        },
        //endregion
        //region updateEndDate ������ ����
        updateEndDate(date) {
            this.endDate = date;
        },
        //endregion
        //region search ���Ӹ� �˻�
        search() {
            this.$emit('search', {
                startDate : this.startDate,
                endDate : this.endDate,
                keyword : this.keyword
            });
        },
        //endregion
        //region reset �ʱ�ȭ
        reset() {
            this.startDate = '';
            this.endDate = '';
            this.keyword = '';
            this.search();
        },
        //endregion
    }
});