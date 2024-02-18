Vue.component('POSTING-LIST', {
    template : `
        <table class="forum-list-tbl">
            <!--region colgroup-->
            <colgroup>
                <col style="width: 35px;">
                <col style="width: 70px;">
                <col style="width: 215px;">
                <col>
                <col style="width: 90px;">
                <col style="width: 150px;">
                <col style="width: 170px;">
            </colgroup>
            <!--endregion-->
            <!--region Thead -->
            <thead>
                <tr>
                    <th></th>
                    <th>idx</th>
                    <th>작성자 정보</th>
                    <th>작성내용</th>
                    <th>상단 고정여부</th>
                    <th>작성일지</th>
                    <th></th>
                </tr>
            </thead>
            <!--endregion-->
            <tbody>
                <template v-if="postings.length > 0">
                    <POSTING v-for="posting in postings" 
                        :key="posting.postingIdx" :posting="posting"
                        :checkedPostingIdx="checkedPostingIdx" @clickPosting="clickPosting"
                        @checkPosting="checkPosting" @deletePosting="deletePosting"/>
                </template>
                <tr v-else>
                    <td colspan="7">등록된 포스팅이 없습니다.</td>
                </tr>
            </tbody>
        </table>
    `,
    props : {
        postings : { type:Array, default:() => { return []; } }, // 포스팅 리스트
        checkedPostingIdx : { type:Number, default:0 } // 체크된 포스팅 일련번호
    },
    methods : {
        //region clickPosting 포스팅 클릭
        clickPosting(posting) {
            this.$emit('clickPosting', posting);
        },
        //endregion
        //region checkPosting 포스팅 체크/해제
        checkPosting(e, postingIdx) {
            this.$emit('checkPosting', e, postingIdx);
        },
        //endregion
        //region deletePosting 포스팅 삭제
        deletePosting(postingIdx) {
            this.$emit('deletePosting', postingIdx);
        },
        //endregion
    }
});