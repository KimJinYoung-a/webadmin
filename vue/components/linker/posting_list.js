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
                    <th>�ۼ��� ����</th>
                    <th>�ۼ�����</th>
                    <th>��� ��������</th>
                    <th>�ۼ�����</th>
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
                    <td colspan="7">��ϵ� �������� �����ϴ�.</td>
                </tr>
            </tbody>
        </table>
    `,
    props : {
        postings : { type:Array, default:() => { return []; } }, // ������ ����Ʈ
        checkedPostingIdx : { type:Number, default:0 } // üũ�� ������ �Ϸù�ȣ
    },
    methods : {
        //region clickPosting ������ Ŭ��
        clickPosting(posting) {
            this.$emit('clickPosting', posting);
        },
        //endregion
        //region checkPosting ������ üũ/����
        checkPosting(e, postingIdx) {
            this.$emit('checkPosting', e, postingIdx);
        },
        //endregion
        //region deletePosting ������ ����
        deletePosting(postingIdx) {
            this.$emit('deletePosting', postingIdx);
        },
        //endregion
    }
});