Vue.component('POSTING', {
    template : `
        <tr>
            <td><input @click="checkPosting($event)" type="checkbox" :checked="posting.postingIdx === checkedPostingIdx" :disabled="posting.fixed"></td>
            <td>{{posting.postingIdx}}</td>
            <td>{{fullCreatorType}} / {{creatorDescription}} / {{posting.creatorNickname}}</td>
            <td @click="$emit('clickPosting', posting)" class="posting-content">{{cutOutPostingContent}}</td>
            <td :class="{'posting-red' : posting.fixed}">{{posting.fixed ? 'Y' : 'N'}}</td>
            <td>{{getLocalDateTimeFormat(posting.regDate, 'yyyy-MM-dd HH:mm:ss')}}</td>
            <td>
                <button @click="$emit('clickPosting', posting)" class="linker-btn">수정</button>
                <button @click="$emit('deletePosting', posting.postingIdx)" class="linker-btn">삭제</button>
            </td>
        </tr>
    `,
    props : {
        //region posting 포스팅
        posting : {
            postingIdx : { type:Number, default:0 }, // 포스팅 일련번호
            creatorType : { type:String, default:'N' }, // 포스팅 작성자 회원 구분 값
            creatorDescr : { type:String, default:'' }, // 포스팅 작성자 회원 설명
            creatorLevelName : { type:String, default:'' }, // 포스팅 작성자 회원 등급
            creatorNickname : { type:String, default:'' }, // 포스팅 작성자 회원 별명
            postingContents : { type:Number, default:0 }, // 포스팅 내용
            fixed : { type:Boolean, default:false }, // 포스팅 고정 여부
        },
        //endregion
        checkedPostingIdx : { type:Number, default:0 }, // 선택한 포스팅 일련번호
    },
    computed : {
        //region cutOutPostingContent 포스팅 내용 자른 값
        cutOutPostingContent() {
            const contents = this.posting.postingContents;
            return contents.length > 60 ? contents.substr(0, 60) + '...' : contents;
        },
        //endregion
        //region fullCreatorType 포스팅 작성자 회원 구분 값 풀네임
        fullCreatorType() {
            switch (this.posting.creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region creatorDescription 포스팅 작성자 설명
        creatorDescription() {
            if( this.posting.creatorType === 'H' || this.posting.creatorType === 'G' ) {
                return this.posting.creatorDescr;
            } else {
                return this.posting.creatorLevelName;
            }
        },
        //endregion
    },
    methods: {
        //region checkPosting 포스팅 체크박스 체크
        checkPosting(e) {
            this.$emit('checkPosting', e, this.posting.postingIdx);
        },
        //endregion
    }
})