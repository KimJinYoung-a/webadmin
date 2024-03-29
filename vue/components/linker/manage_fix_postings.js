Vue.component('MANAGE-FIX-POSTINGS', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{postings.length}}</span>개</p>
                <div>
                    <button @click="clearFixPostings" class="linker-btn">고정 해제</button>
                    <button @click="modifyFixPositions" class="linker-btn">노출위치 수정</button>
                </div>
            </div>

            <table class="forum-list-tbl">
                <!--region Colgroup-->
                <colgroup>
                    <col style="width: 30px;">
                    <col style="width: 60px;">
                    <col style="width: 170px;">
                    <col>
                    <col style="width: 100px;">
                    <col style="width: 160px;">
                    <col style="width: 170px;">
                </colgroup>
                <!--endregion-->
                <!--region THead-->
                <thead>
                    <tr>
                        <th><input v-if="postings.length > 0" @click="checkAll" :checked="checkedPostingIdxs.length === postings.length" type="checkbox"></th>
                        <th>idx</th>
                        <th>작성자 정보</th>
                        <th>작성내용</th>
                        <th>고정 노출 위치</th>
                        <th>고정 기간</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <tbody>
                    <template v-if="postings && postings.length > 0">
                        <tr v-for="(posting, index) in postings">
                            <td><input type="checkbox" @click="checkPosting(posting.postingIdx)" :checked="checkedPostingIdxs.indexOf(posting.postingIdx) > -1"></td>
                            <td>{{posting.postingIdx}}</td>
                            <td>
                                {{fullCreatorType(posting.creatorType)}} 
                                / {{creatorDescription(posting.creatorType, posting.creatorDescr, posting.creatorLevelName)}} 
                                / {{posting.creatorNickname}}
                            </td>
                            <td>{{cutoutPostingContents(posting.postingContents)}}</td>
                            <td><input type="text" v-if="positionNumbers[index]" class="forum-sort" v-model.number="positionNumbers[index].positionNo"></td>
                            <td>{{fixPeriod(posting)}}</td>
                            <td>
                                <button @click="$emit('modifyPosting', posting)" class="linker-btn">수정</button>
                                <button @click="clearFixPosting(posting.postingIdx)" class="linker-btn long">고정 해제</button>
                            </td>
                        </tr>
                    </template>
                    <tr v-else>
                        <td colspan="7">고정 포스팅이 존재하지 않습니다.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    data() {return {
        checkedPostingIdxs : [], // 체크한 포스팅 일련번호 리스트
        positionNumbers : [], // 정렬 순서 리스트
    }},
    props : {
        //region postings 고정 포스팅 리스트
        postings : {
            postingIdx : { type:Number, default:0 }, // 포스팅 일련번호
            creatorType : { type:String, default:'N' }, // 포스팅 작성자 회원 구분 값
            creatorDescr : { type:String, default:'' }, // 포스팅 작성자 회원 설명
            creatorLevelName : { type:String, default:'' }, // 포스팅 작성자 회원 등급
            creatorNickname : { type:String, default:'' }, // 포스팅 작성자 회원 별명
            creatorThumbnail : { type:String, default:'' }, // 포스팅 작성자 썸네일 이미지
            postingContents : { type:Number, default:0 }, // 포스팅 내용
            positionNo : { type:Number, default:0 }, // 고정 노출 위치
            startDate : { type:String, default:'' }, // 고정 포스팅 시작일자
            endDate : { type:String, default:'' }, // 고정 포스팅 시작일자
        },
        //endregion
    },
    computed : {
        modifyPositionApiData() {
            const data = {};
            if( this.positionNumbers.length > 0 ) {
                this.positionNumbers.forEach((p, index) => {
                    data[`postingsFixed[${index}].postingIndex`] = p.postingIndex;
                    data[`postingsFixed[${index}].positionNo`] = p.positionNo;
                });
            }
            return data;
        },
    },
    methods : {
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
        //region fixPeriod 고정 기간
        fixPeriod(posting) {
            return this.getLocalDateTimeFormat(posting.startDate, 'yyyy-MM-dd') + ' ~ '
                + this.getLocalDateTimeFormat(posting.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region clearFixPosting 포스팅 하나 고정 해제
        clearFixPosting(postingIdx) {
            if( confirm('이 포스팅의 고정을 해제 하시겠습니까?') )
                this.$emit('clearPostings', [postingIdx]);
        },
        //endregion
        //region clearFixPostings 선택한 포스팅들 고정 해제
        clearFixPostings() {
            if( confirm('선택한 포스팅들의 고정을 해제 하시겠습니까?') )
                this.$emit('clearPostings', this.checkedPostingIdxs);
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
        //region modifyFixPositions 노출 위치 수정
        modifyFixPositions() {
            if( confirm('현재 상태로 노출 위치를 수정하시겠습니까?') )
                this.$emit('modifyFixPositions', this.modifyPositionApiData);
        },
        //endregion
    },
    watch : {
        postings(postings) {
            this.positionNumbers = postings.map(p => {
                return {
                    postingIndex: p.postingIdx,
                    positionNo: p.positionNo
                };
            });
        }
    }
});