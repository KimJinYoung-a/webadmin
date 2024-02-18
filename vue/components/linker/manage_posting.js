Vue.component('MANAGE-POSTING', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:120px;">
                    <col>
                </colgroup>
                <tbody>
                    <tr>
                        <th>idx</th>
                        <td class="content">{{posting.postingIdx}}</td>
                    </tr>
                    <tr>
                        <th>회원구분</th>
                        <td class="content">{{fullCreatorType}}</td>
                    </tr>
                    <tr v-if="posting.creatorType === 'H' || posting.creatorType === 'G'">
                        <th>회원설명</th>
                        <td class="content">{{posting.creatorDescr}}</td>
                    </tr>
                    <tr v-else>
                        <th>회원등급</th>
                        <td class="content">{{posting.creatorLevelName}}</td>
                    </tr>
                    <tr>
                        <th>닉네임</th>
                        <td class="content">{{posting.creatorNickname}}</td>
                    </tr>
                    <tr>
                        <th>작성내용</th>
                        <td><textarea v-model="tempContent" rows="9"></textarea></td>
                    </tr>
                    <tr>
                        <th>연결컨텐츠</th>
                        <td>
                            <button v-if="posting.linkValue" @click="clickLink" class="linker-btn">
                                {{linkButtonText}}
                            </button>
                        </td>
                    </tr>
                    <tr>
                        <th>상단 고정 여부</th>
                        <td>
                            <input @click="clickFixRadio($event, true)" id="fixPostingY" type="radio" name="fix" :checked="fix.use">
                            <label for="fixPostingY">고정</label>
                            <input @click="clickFixRadio($event, false)" id="fixPostingN" type="radio" name="fix" :checked="!fix.use">
                            <label for="fixPostingN">고정안함</label>
                        </td>
                    </tr>
                    <tr>
                        <th>작성일시</th>
                        <td class="content">
                            <strong>{{getLocalDateTimeFormat(posting.regDate, 'yyyy-MM-dd HH:mm:ss')}}</strong>
                            <span v-if="posting.updateDate" class="posting-update">
                                {{getLocalDateTimeFormat(posting.updateDate, 'yyyy-MM-dd HH:mm:ss')}} 수정
                            </span>
                        </td>
                    </tr>
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="savePosting" class="linker-btn">저장</button>
            </div>
        </div>
    `,
    mounted() {
        this.tempContent = this.posting.postingContents;
        if( this.posting.fixed ) {
            this.fix = {
                positionNo : this.posting.positionNo,
                startDate : this.getLocalDateTimeFormat(this.posting.startDate, 'yyyy-MM-dd'),
                endDate : this.getLocalDateTimeFormat(this.posting.endDate, 'yyyy-MM-dd'),
                use : true
            }
        }
    },
    data() {return {
        tempContent : null,
        fix : {
            positionNo : null,
            startDate : null,
            endDate : null,
            use : false
        },
    }},
    props : {
        //region posting 포스팅
        posting: {
            postingIdx : { type:Number, default:0 }, // 포스팅 일련번호
            creatorType : { type:String, default:'N' }, // 포스팅 작성자 회원 구분 값
            creatorDescr : { type:String, default:'' }, // 포스팅 작성자 회원 설명
            creatorLevelName : { type:String, default:'' }, // 포스팅 작성자 회원 등급
            creatorNickname : { type:String, default:'' }, // 포스팅 작성자 회원 별명
            postingContents : { type:Number, default:0 }, // 포스팅 내용
            linkType : { type:Number, default:null }, // 포스팅 링크 유형
            linkValue : { type:String, default:'' }, // 포스팅 링크 값
            fixed : { type:Boolean, default:false }, // 포스팅 고정 여부
            regDate : { type:String, default:'' }, // 포스팅 고정 여부
        },
        //endregion
    },
    computed : {
        //region fullCreatorType 포스팅 작성자 회원 구분 값 풀네임
        fullCreatorType() {
            switch (this.posting.creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region linkButtonText 링크 버튼 텍스트
        linkButtonText() {
            let linkTitle;
            switch( this.posting.linkType ) {
                case 1: linkTitle = '상품'; break;
                case 2: linkTitle = '이벤트'; break;
            }

            return (linkTitle ? (linkTitle + ' : ') : '') + this.posting.linkValue;
        },
        //endregion
    },
    methods : {
        //region clickLink 링크 클릭
        clickLink() {
            let url;
            const wwwUrl = this.isDevelop ? '//2015www.10x10.co.kr' : '//www.10x10.co.kr';
            const linkType = this.posting.linkType;
            const linkValue = this.posting.linkValue;

            if( linkType === 1 )
                url = wwwUrl + '/shopping/category_prd.asp?itemid=' + linkValue;
            else if( linkType === 2 )
                url = wwwUrl + '/event/eventmain.asp?eventid=' + linkValue;
            else
                url = linkValue;

            window.open(url, '_blank');
        },
        //endregion
        //region clickFixRadio 고정 여부 클릭
        clickFixRadio(e, fixFlag) {
            if( fixFlag ) {
                e.preventDefault();
                this.$emit('openFixPostingModal', this.fix);
            } else {
                this.clearFix();
            }
        },
        //endregion
        //region savePosting 포스팅 저장
        savePosting() {
            if( confirm('저장 하시겠습니까?') )
                this.$emit('savePosting', this.createSavePostingData());
        },
        createSavePostingData() {
            const data = {
                postingIndex : this.posting.postingIdx,
                postingCotents : this.tempContent,
                useYn : this.fix.use
            }
            if( this.fix.use ) {
                data.startDate = this.fix.startDate;
                data.endDate = this.fix.endDate;
                data.positionNo = this.fix.positionNo;
            }

            return data;
        },
        //endregion
        //region setFixPosting 포스팅 고정정보 Set
        setFixPosting(fix) {
            this.fix = {
                positionNo : fix.positionNo,
                startDate : fix.startDate,
                endDate : fix.endDate,
                use : true
            }
        },
        //endregion
        //region clearFix 포스팅 고정정보 제거
        clearFix() {
            this.fix = {
                positionNo : null,
                startDate : null,
                endDate : null,
                use : false
            }
        },
        //endregion
    }
});