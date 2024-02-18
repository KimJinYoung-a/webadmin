Vue.component('MANAGE-FORUM-SORT', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{forums.length}}</span>개</p>
                <div>
                    <button @click="modifySortNo" class="linker-btn">정렬순서 수정</button>
                </div>
            </div>

            <table class="forum-list-tbl">
                <colgroup>
                    <col style="width: 60px;">
                    <col>
                    <col style="width: 200px;">
                    <col style="width: 120px;">
                    <col style="width: 70px;">
                    <col style="width: 160px;">
                    <col style="width: 140px;">
                </colgroup>
                <!--region THead -->
                <thead>
                    <tr>
                        <th>NO.</th>
                        <th>포럼 제목</th>
                        <th>포럼 부제목</th>
                        <th>프론트 오픈 여부</th>
                        <th>정렬순서</th>
                        <th>운영기간</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                
                <tbody>
                    <tr v-for="forum in sortedForumsBySortNo">
                        <td>{{forum.forumIdx}}</td>
                        <td v-html="forum.title"></td>
                        <td v-html="forum.subTitle"></td>
                        <td>{{forum.useYn ? '오픈' : '오픈안함'}}</td>
                        <td><input type="text" v-model="forum.sortNo" class="forum-sort"></td>
                        <td>{{getForumPeriod(forum)}}</td>
                        <td>
                            <button @click="modifyForum(forum)" class="linker-btn">수정</button>
                            <button @click="deleteForum(forum.forumIdx)" class="linker-btn">삭제</button>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    mounted() {
        this.syncSortedForumsBySortNo();
    },
    data() {return {
        sortedForumsBySortNo : []
    }},
    props : {
        forums : { type:Array, default:function(){return [];} }, // 포럼 리스트
    },
    computed : {
        //region modifySortNoApiData 순서 정렬 수정 API 전달 데이터
        modifySortNoApiData() {
            const data = {};
            this.sortedForumsBySortNo.forEach((f, i) => {
                data[`forumSorts[${i}].forumIndex`] = f.forumIdx;
                data[`forumSorts[${i}].sortNo`] = f.sortNo;
            });
            return data;
        },
        //endregion
    },
    methods : {
        //region getForumPeriod Get 포럼 오픈 기간
        getForumPeriod(forum) {
            return this.getLocalDateTimeFormat(forum.startDate, 'yyyy-MM-dd')
                + ' ~ ' + this.getLocalDateTimeFormat(forum.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region modifySortNo 노출 순서 수정
        modifySortNo() {
            this.callApi(2, 'POST', '/linker/forum/sort/update', this.modifySortNoApiData, this.successModifySortNo);
        },
        successModifySortNo() {
            this.$emit('modifySortNo', this.sortedForumsBySortNo);
        },
        //endregion
        //region syncSortedForumsBySortNo 포럼 리스트 동기화
        syncSortedForumsBySortNo() {
            this.sortedForumsBySortNo = [];
            this.forums.forEach(f => this.sortedForumsBySortNo.push(f));
            this.sortedForumsBySortNo.sort((a, b) => {
                return a.sortNo - b.sortNo;
            });
        },
        //endregion
        //region modifyForum 포럼 수정
        modifyForum(forum) {
            this.$emit('modifyForum', forum);
        },
        //endregion
        //region deleteForum 포럼 삭제
        deleteForum(forumIdx) {
            this.$emit('deleteForum', forumIdx);
        },
        //endregion
    }
});