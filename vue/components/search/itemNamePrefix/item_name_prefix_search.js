Vue.component('ITEM-NAME-PREFIX-SEARCH', {
    template : `
        <div class="search">
            <div class="search-group">
                <label>기간</label>
                <DATE-PICKER @updateDate="updateStartDate" :date="startDate" id="searchStartDate"/>
                <DATE-PICKER @updateDate="updateEndDate" :date="endDate" id="searchEndDate"/>
            </div>
            <div class="search-group">
                <label>말머리</label>
                <input v-model="keyword" type="text">
            </div>
            <button @click="search" class="btn search-btn">검색</button>
            <button @click="reset" class="btn search-btn">초기화</button>
        </div>
    `,
    data() {return {
        startDate : '',
        endDate : '',
        keyword : '',
    }},
    methods : {
        //region updateStartDate 시작일 수정
        updateStartDate(date) {
            this.startDate = date;
        },
        //endregion
        //region updateEndDate 종료일 수정
        updateEndDate(date) {
            this.endDate = date;
        },
        //endregion
        //region search 말머리 검색
        search() {
            this.$emit('search', {
                startDate : this.startDate,
                endDate : this.endDate,
                keyword : this.keyword
            });
        },
        //endregion
        //region reset 초기화
        reset() {
            this.startDate = '';
            this.endDate = '';
            this.keyword = '';
            this.search();
        },
        //endregion
    }
});