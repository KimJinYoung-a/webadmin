Vue.component('MANAGE-FORUM-SORT', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{forums.length}}</span>��</p>
                <div>
                    <button @click="modifySortNo" class="linker-btn">���ļ��� ����</button>
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
                        <th>���� ����</th>
                        <th>���� ������</th>
                        <th>����Ʈ ���� ����</th>
                        <th>���ļ���</th>
                        <th>��Ⱓ</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                
                <tbody>
                    <tr v-for="forum in sortedForumsBySortNo">
                        <td>{{forum.forumIdx}}</td>
                        <td v-html="forum.title"></td>
                        <td v-html="forum.subTitle"></td>
                        <td>{{forum.useYn ? '����' : '���¾���'}}</td>
                        <td><input type="text" v-model="forum.sortNo" class="forum-sort"></td>
                        <td>{{getForumPeriod(forum)}}</td>
                        <td>
                            <button @click="modifyForum(forum)" class="linker-btn">����</button>
                            <button @click="deleteForum(forum.forumIdx)" class="linker-btn">����</button>
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
        forums : { type:Array, default:function(){return [];} }, // ���� ����Ʈ
    },
    computed : {
        //region modifySortNoApiData ���� ���� ���� API ���� ������
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
        //region getForumPeriod Get ���� ���� �Ⱓ
        getForumPeriod(forum) {
            return this.getLocalDateTimeFormat(forum.startDate, 'yyyy-MM-dd')
                + ' ~ ' + this.getLocalDateTimeFormat(forum.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region modifySortNo ���� ���� ����
        modifySortNo() {
            this.callApi(2, 'POST', '/linker/forum/sort/update', this.modifySortNoApiData, this.successModifySortNo);
        },
        successModifySortNo() {
            this.$emit('modifySortNo', this.sortedForumsBySortNo);
        },
        //endregion
        //region syncSortedForumsBySortNo ���� ����Ʈ ����ȭ
        syncSortedForumsBySortNo() {
            this.sortedForumsBySortNo = [];
            this.forums.forEach(f => this.sortedForumsBySortNo.push(f));
            this.sortedForumsBySortNo.sort((a, b) => {
                return a.sortNo - b.sortNo;
            });
        },
        //endregion
        //region modifyForum ���� ����
        modifyForum(forum) {
            this.$emit('modifyForum', forum);
        },
        //endregion
        //region deleteForum ���� ����
        deleteForum(forumIdx) {
            this.$emit('deleteForum', forumIdx);
        },
        //endregion
    }
});