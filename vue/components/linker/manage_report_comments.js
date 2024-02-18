Vue.component('MANAGE-REPORT-COMMENTS', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{comments.length}}</span>��</p>
                <div>
                    <button @click="unBlockSelectedComments" class="linker-btn">��� ����</button>
                    <button @click="deleteSelectedComments" class="linker-btn">���� ��� ����</button>
                </div>
            </div>

            <table class="forum-list-tbl">
                <!--region colgroup-->
                <colgroup>
                    <col style="width: 30px;">
                    <col style="width: 60px;">
                    <col style="width: 80px;">
                    <col style="width: 170px;">
                    <col>
                    <col style="width: 100px;">
                    <col style="width: 200px;">
                </colgroup>
                <!--endregion-->
                <!--region THead-->
                <thead>
                    <tr>
                        <th><input v-if="comments.length > 0" @click="checkAll" :checked="checkedCommentIdxs.length === comments.length" type="checkbox"></th>
                        <th>idx</th>
                        <th>������ �̹���</th>
                        <th>�ۼ��� ����</th>
                        <th>�ۼ�����</th>
                        <th>������ idx</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <tbody>
                    <template v-if="comments && comments.length > 0">
                        <tr v-for="comment in comments">
                            <td><input type="checkbox" @click="checkComment(comment.commentIdx)" :checked="checkedCommentIdxs.indexOf(comment.commentIdx) > -1"></td>
                            <td>{{comment.commentIdx}}</td>
                            <td><img :src="decodeBase64(comment.creatorThumbnail)" class="thumb"></td>
                            <td>
                                {{fullCreatorType(comment.creatorType)}} 
                                / {{creatorDescription(comment.creatorType, comment.creatorDescr, comment.creatorLevelName)}} 
                                / {{comment.creatorNickname}}
                            </td>
                            <td>{{cutoutCommentContents(comment.commentContents)}}</td>
                            <td>{{comment.postingIdx}}</td>
                            <td>
                                <button @click="unBlockComments([comment.commentIdx])" class="linker-btn long">��� ����</button>
                                <button @click="deleteComments([comment.commentIdx])" class="linker-btn long">��� ����</button>
                            </td>
                        </tr>
                    </template>
                    <tr v-else>
                        <td colspan="7">�Ű�� ����� �������� �ʽ��ϴ�.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    data() {return {
        checkedCommentIdxs : [], // üũ�� ��� �Ϸù�ȣ ����Ʈ
    }},
    props : {
        //region comments ������ ����Ʈ
        comments : {
            commentIdx : { type:Number, default:0 }, // ��� �Ϸù�ȣ
            postingIdx : { type:Number, default:0 }, // ������ �Ϸù�ȣ
            creatorType : { type:String, default:'N' }, // ������ �ۼ��� ȸ�� ���� ��
            creatorDescr : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorLevelName : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ���
            creatorNickname : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorThumbnail : { type:String, default:'' }, // ������ �ۼ��� ����� �̹���
            commentContents : { type:Number, default:0 }, // ������ ����
        },
        //endregion
    },
    methods: {
        //region fullCreatorType ������ �ۼ��� ȸ�� ���� �� Ǯ����
        fullCreatorType(creatorType) {
            switch (creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region creatorDescription ������ �ۼ��� ����
        creatorDescription(type, description, levelName) {
            if( type === 'H' || type === 'G' ) {
                return description;
            } else {
                return levelName;
            }
        },
        //endregion
        //region cutoutCommentContents ��� ���� 60�� �̳��� �ڸ�
        cutoutCommentContents(contents) {
            if( contents.length > 60 )
                return contents.substr(0, 60) + '...';
            else
                return contents;
        },
        //endregion
        //region deleteComment ��� ����
        deleteComment(commentIdx) {
            this.$emit('deleteComment', commentIdx, true);
        },
        //endregion
        //region unBlockComments ��� ��� ����
        unBlockComments(commentIdxs) {
            if( confirm('�� ����� ����� �����Ͻðڽ��ϱ�?\n��� �Ű����� �ʱ�ȭ�˴ϴ�.') )
                this.$emit('unBlockComments', commentIdxs);
        },
        //endregion
        //region checkAll ��ü ��� ����
        checkAll() {
            if( this.checkedCommentIdxs.length !== this.comments.length ) {
                this.checkedCommentIdxs = this.comments.map(c => c.commentIdx);
            } else {
                this.checkedCommentIdxs = [];
            }
        },
        //endregion
        //region checkComment üũ/���� ���
        checkComment(commentIdx) {
            if( this.checkedCommentIdxs.indexOf(commentIdx) > -1 ) {
                this.checkedCommentIdxs.splice(this.checkedCommentIdxs.findIndex(i => i === commentIdx), 1);
            } else {
                this.checkedCommentIdxs.push(commentIdx);
            }
        },
        //endregion
        //region unBlockSelectedComments ���� ��۵� ��� ����
        unBlockSelectedComments() {
            if( this.checkedCommentIdxs.length > 0 &&
                confirm('������ ��۵��� ����� �����Ͻðڽ��ϱ�?\n�ش� ��۵��� ��� �Ű����� �ʱ�ȭ�˴ϴ�.') ) {
                this.$emit('unBlockComments', this.checkedCommentIdxs);
            }
        },
        //endregion
        //region deleteSelectedComments ���� ��۵� ����
        deleteSelectedComments() {
            if( this.checkedCommentIdxs.length > 0 &&
                confirm('������ ��۵��� �����Ͻðڽ��ϱ�?') ) {
                this.$emit('deleteComments', this.checkedCommentIdxs);
            }
        },
        //endregion
    }
});