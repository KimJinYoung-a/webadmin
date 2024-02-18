Vue.component('MANAGE-REPORT-POSTINGS', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{numberFormat(postingCount)}}</span>개</p>
                <div>
                    <button @click="unBlockSelectedPostings" class="linker-btn">블락 해제</button>
                    <button @click="deleteSelectedPostings" class="linker-btn">선택 포스팅 삭제</button>
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
                    <col style="width: 170px;">
                    <col style="width: 200px;">
                </colgroup>
                <!--endregion-->
                <!--region THead-->
                <thead>
                    <tr>
                        <th><input v-if="postings.length > 0" @click="checkAll" :checked="checkedPostingIdxs.length === postings.length" type="checkbox"></th>
                        <th>idx</th>
                        <th>프로필 이미지</th>
                        <th>작성자 정보</th>
                        <th>작성내용</th>
                        <th>연결된 컨텐츠</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <tbody>
                    <template v-if="postings && postings.length > 0">
                        <tr v-for="posting in postings">
                            <td><input type="checkbox" @click="checkPosting(posting.postingIdx)" :checked="checkedPostingIdxs.indexOf(posting.postingIdx) > -1"></td>
                            <td>{{posting.postingIdx}}</td>
                            <td><img :src="decodeBase64(posting.creatorThumbnail)" class="thumb"></td>
                            <td>
                                {{fullCreatorType(posting.creatorType)}} 
                                / {{creatorDescription(posting.creatorType, posting.creatorDescr, posting.creatorLevelName)}} 
                                / {{posting.creatorNickname}}
                            </td>
                            <td>{{cutoutPostingContents(posting.postingContents)}}</td>
                            <td>
                                <button v-if="posting.linkValue" @click="clickLink(posting.linkType, posting.linkValue)" class="linker-btn link">
                                    {{linkButtonText(posting.linkType, posting.linkValue)}}
                                </button>
                            </td>
                            <td>
                                <button @click="unBlockPostings([posting.postingIdx])" class="linker-btn long">블락 해제</button>
                                <button @click="deletePosting(posting.postingIdx)" class="linker-btn long">포스팅 삭제</button>
                            </td>
                        </tr>
                    </template>
                    <tr v-else>
                        <td colspan="7">신고된 포스팅이 존재하지 않습니다.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    data() {return {
        checkedPostingIdxs : [], // 체크한 포스팅 일련번호 리스트
    }},
    props : {
        postingCount : { type:Number, default:0 }, // 포스팅 갯수
        //region postings 포스팅 리스트
        postings : {
            postingIdx : { type:Number, default:0 }, // 포스팅 일련번호
            creatorType : { type:String, default:'N' }, // 포스팅 작성자 회원 구분 값
            creatorDescr : { type:String, default:'' }, // 포스팅 작성자 회원 설명
            creatorLevelName : { type:String, default:'' }, // 포스팅 작성자 회원 등급
            creatorNickname : { type:String, default:'' }, // 포스팅 작성자 회원 별명
            creatorThumbnail : { type:String, default:'' }, // 포스팅 작성자 썸네일 이미지
            postingContents : { type:Number, default:0 }, // 포스팅 내용
            linkType : { type:String, default:'' }, // 링크 유형
            linkValue : { type:String, default:'' }, // 링크 값
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
        //region cutoutPostingContents 포스팅 내용 60자 이내로 자름
        cutoutPostingContents(contents) {
            if( contents.length > 60 )
                return contents.substr(0, 60) + '...';
            else
                return contents;
        },
        //endregion
        //region linkButtonText 링크 버튼 텍스트
        linkButtonText(linkType, linkValue) {
            switch( linkType ) {
                case 1: return '상품 : ' + linkValue;
                case 2: return '이벤트 : ' + linkValue;
                default: return '외부 URL';
            }
        },
        //endregion
        //region clickLink 링크 클릭
        clickLink(linkType, linkValue) {
            let url;
            const wwwUrl = this.isDevelop ? '//2015www.10x10.co.kr' : '//www.10x10.co.kr';

            if( linkType === 1 )
                url = wwwUrl + '/shopping/category_prd.asp?itemid=' + linkValue;
            else if( linkType === 2 )
                url = wwwUrl + '/event/eventmain.asp?eventid=' + linkValue;
            else
                url = linkValue;

            window.open(url, '_blank');
        },
        //endregion
        //region deletePosting 포스팅 삭제
        deletePosting(postingIdx) {
            this.$emit('deletePosting', postingIdx, true);
        },
        //endregion
        //region unBlockPostings 포스팅 블락 해제
        unBlockPostings(postingIdxs) {
            if( confirm('이 포스팅의 블락을 해제하시겠습니까?\n모든 신고기록이 초기화됩니다.') )
                this.$emit('unBlockPostings', postingIdxs);
        },
        //endregion
        //region checkAll 전체 포스팅 선택
        checkAll() {
            if( this.checkedPostingIdxs.length !== this.postings.length ) {
                this.checkedPostingIdxs = this.postings.map(p => p.postingIdx);
            } else {
                this.checkedPostingIdxs = [];
            }
        },
        //endregion
        //region checkPosting 체크/해제 포스팅
        checkPosting(postingIdx) {
            if( this.checkedPostingIdxs.indexOf(postingIdx) > -1 ) {
                this.checkedPostingIdxs.splice(this.checkedPostingIdxs.findIndex(p => p === postingIdx), 1);
            } else {
                this.checkedPostingIdxs.push(postingIdx);
            }
        },
        //endregion
        //region unBlockSelectedPostings 선택 포스팅들 블락 해제
        unBlockSelectedPostings() {
            if( this.checkedPostingIdxs.length > 0 &&
                confirm('선택한 포스팅들의 블락을 해제하시겠습니까?\n해당 포스팅들의 모든 신고기록이 초기화됩니다.') ) {
                this.$emit('unBlockPostings', this.checkedPostingIdxs);
            }
        },
        //endregion
        //region deleteSelectedPostings 선택 포스팅들 삭제
        deleteSelectedPostings() {
            if( this.checkedPostingIdxs.length > 0 &&
                confirm('선택한 포스팅들을 삭제하시겠습니까?') ) {
                this.$emit('deletePostings', this.checkedPostingIdxs);
            }
        },
        //endregion
        //region numberFormat 숫자 천자리 (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
    }
});