Vue.component('MANAGE-REPORT-COMMENTS', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{comments.length}}</span>개</p>
                <div>
                    <button @click="unBlockSelectedComments" class="linker-btn">블락 해제</button>
                    <button @click="deleteSelectedComments" class="linker-btn">선택 댓글 삭제</button>
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
                        <th>프로필 이미지</th>
                        <th>작성자 정보</th>
                        <th>작성내용</th>
                        <th>포스팅 idx</th>
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
                                <button @click="unBlockComments([comment.commentIdx])" class="linker-btn long">블락 해제</button>
                                <button @click="deleteComments([comment.commentIdx])" class="linker-btn long">댓글 삭제</button>
                            </td>
                        </tr>
                    </template>
                    <tr v-else>
                        <td colspan="7">신고된 댓글이 존재하지 않습니다.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    data() {return {
        checkedCommentIdxs : [], // 체크한 댓글 일련번호 리스트
    }},
    props : {
        //region comments 포스팅 리스트
        comments : {
            commentIdx : { type:Number, default:0 }, // 댓글 일련번호
            postingIdx : { type:Number, default:0 }, // 포스팅 일련번호
            creatorType : { type:String, default:'N' }, // 포스팅 작성자 회원 구분 값
            creatorDescr : { type:String, default:'' }, // 포스팅 작성자 회원 설명
            creatorLevelName : { type:String, default:'' }, // 포스팅 작성자 회원 등급
            creatorNickname : { type:String, default:'' }, // 포스팅 작성자 회원 별명
            creatorThumbnail : { type:String, default:'' }, // 포스팅 작성자 썸네일 이미지
            commentContents : { type:Number, default:0 }, // 포스팅 내용
        },
        //endregion
    },
    methods: {
        //region fullCreatorType 포스팅 작성자 회원 구분 값 풀네임
        fullCreatorType(creatorType) {
            switch (creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region creatorDescription 포스팅 작성자 설명
        creatorDescription(type, description, levelName) {
            if( type === 'H' || type === 'G' ) {
                return description;
            } else {
                return levelName;
            }
        },
        //endregion
        //region cutoutCommentContents 댓글 내용 60자 이내로 자름
        cutoutCommentContents(contents) {
            if( contents.length > 60 )
                return contents.substr(0, 60) + '...';
            else
                return contents;
        },
        //endregion
        //region deleteComment 댓글 삭제
        deleteComment(commentIdx) {
            this.$emit('deleteComment', commentIdx, true);
        },
        //endregion
        //region unBlockComments 댓글 블락 해제
        unBlockComments(commentIdxs) {
            if( confirm('이 댓글의 블락을 해제하시겠습니까?\n모든 신고기록이 초기화됩니다.') )
                this.$emit('unBlockComments', commentIdxs);
        },
        //endregion
        //region checkAll 전체 댓글 선택
        checkAll() {
            if( this.checkedCommentIdxs.length !== this.comments.length ) {
                this.checkedCommentIdxs = this.comments.map(c => c.commentIdx);
            } else {
                this.checkedCommentIdxs = [];
            }
        },
        //endregion
        //region checkComment 체크/해제 댓글
        checkComment(commentIdx) {
            if( this.checkedCommentIdxs.indexOf(commentIdx) > -1 ) {
                this.checkedCommentIdxs.splice(this.checkedCommentIdxs.findIndex(i => i === commentIdx), 1);
            } else {
                this.checkedCommentIdxs.push(commentIdx);
            }
        },
        //endregion
        //region unBlockSelectedComments 선택 댓글들 블락 해제
        unBlockSelectedComments() {
            if( this.checkedCommentIdxs.length > 0 &&
                confirm('선택한 댓글들의 블락을 해제하시겠습니까?\n해당 댓글들의 모든 신고기록이 초기화됩니다.') ) {
                this.$emit('unBlockComments', this.checkedCommentIdxs);
            }
        },
        //endregion
        //region deleteSelectedComments 선택 댓글들 삭제
        deleteSelectedComments() {
            if( this.checkedCommentIdxs.length > 0 &&
                confirm('선택한 댓글들을 삭제하시겠습니까?') ) {
                this.$emit('deleteComments', this.checkedCommentIdxs);
            }
        },
        //endregion
    }
});