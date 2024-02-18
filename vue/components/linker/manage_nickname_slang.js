Vue.component('MANAGE-NICKNAME-SLANG', {
    template : `
        <div>
            <div class="search">
                <div>
                    <div class="search-group">
                        <label>비속어:</label>
                        <input id="keyword" @keydown.enter="updateKeyword" type="text">
                    </div>
                    <button @click="updateKeyword" class="linker-btn">검색</button>
                </div>
            </div>

            <div>
                <div class="nicknames-btn-area">
                    <button @click="$emit('openPostModal', 3)" class="linker-btn">신규등록</button>
                    <button @click="deleteWords" class="linker-btn">삭제</button>
                </div>

                <table class="forum-list-tbl">
                    <!--region Colgroup-->
                    <colgroup>
                        <col style="width: 50px;">
                        <col style="width: 60px;">
                        <col>
                        <col style="width: 150px;">
                    </colgroup>
                    <!--endregion-->
                    <!--region THead-->
                    <thead>
                        <tr>
                            <th>
                                <input v-if="searchWords.length > 0" @click="checkAll" 
                                    :checked="checkedWordIdxs.length === searchWords.length" type="checkbox">
                            </th>
                            <th>NO.</th>
                            <th>비속어</th>
                            <th></th>
                        </tr>
                    </thead>
                    <!--endregion-->
                    <tbody>
                        <template v-if="searchWords.length > 0">
                            <tr v-for="word in searchWords">
                                <td>
                                    <input type="checkbox" @click="checkWord(word.wordIdx)"
                                        :checked="checkedWordIdxs.indexOf(word.wordIdx) > -1">
                                </td>
                                <td>{{word.wordIdx}}</td>
                                <td>{{word.word}}</td>
                                <td>
                                    <button @click="$emit('modifyWord', 3, word)" class="linker-btn">수정</button>
                                    <button @click="deleteWord(word)" class="linker-btn">삭제</button>
                                </td>
                            </tr>
                        </template>
                        <tr v-else>
                            <td colspan="4">비속어가 없습니다.</td>
                        </tr>
                    </tbody>
                </table>
            </div>

        </div>
    `,
    data() {return {
        keyword : '', // 검색 키워드
        checkedWordIdxs : [], // 체크한 단어 일련번호 리스트
    }},
    computed : {
        //region searchWords 단어 검색 결과 리스트
        searchWords() {
            const keyword = this.keyword.trim();
            if( keyword ) {
                return this.words.filter(w => w.word.indexOf(keyword) > -1);
            } else {
                return this.words;
            }
        },
        //endregion
    },
    props : {
        words : {
            wordIdx : { type:Number, default:0 },
            word : { type:String, default:'' },
        },
    },
    methods : {
        //region updateKeyword 검색 키워드 수정
        updateKeyword() {
            this.checkedWordIdxs = [];
            this.keyword = document.getElementById('keyword').value;
        },
        //endregion
        //region checkAll 전체 선택
        checkAll() {
            if( this.checkedWordIdxs.length !== this.searchWords.length ) {
                this.checkedWordIdxs = this.searchWords.map(w => w.wordIdx);
            } else {
                this.checkedWordIdxs = [];
            }
        },
        //endregion
        //region checkWord 단어 선택
        checkWord(wordIdx) {
            if( this.checkedWordIdxs.indexOf(wordIdx) > -1 ) {
                this.checkedWordIdxs.splice(this.checkedWordIdxs.findIndex(w => w === wordIdx), 1);
            } else {
                this.checkedWordIdxs.push(wordIdx);
            }
        },
        //endregion
        //region deleteWord 단어 하나 삭제
        deleteWord(word) {
            if( confirm(`'${word.word}' 이 단어를 삭제하시겠습니까?`) )
                this.$emit('deleteWords', 3, [word.wordIdx]);
        },
        //endregion
        //region deleteWords 단어 여러개 삭제
        deleteWords() {
            if( this.checkedWordIdxs.length > 0 && confirm('선택한 단어들을 삭제하시겠습니까?') )
                this.$emit('deleteWords', 3, this.checkedWordIdxs);
        },
        //endregion
    }
});