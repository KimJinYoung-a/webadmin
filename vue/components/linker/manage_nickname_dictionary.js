Vue.component('MANAGE-NICKNAME-DICTIONARY', {
    template : `
        <div>
            <!--region 단어 검색-->
            <div class="search">
                <div>
                    <div class="search-group">
                        <select v-model.number="searchWordNumber">
                            <option value="1">단어1</option>
                            <option value="2">단어2</option>
                        </select>
                        :
                        <input id="keyword" @keydown.enter="updateKeyword" type="text">
                    </div>
                    <button @click="updateKeyword" class="linker-btn">검색</button>
                </div>
            </div>
            <!--endregion-->

            <div class="modal-nicknames-area">
                <!--region 단어1-->
                <div class="modal-nicknames-content">
                    <div class="nicknames-btn-area">
                        <button @click="openPostModal(1)" class="linker-btn">신규등록</button>
                        <button @click="deleteWords(1)" class="linker-btn">삭제</button>
                    </div>

                    <table class="forum-list-tbl">
                        <!--region colgroup-->
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
                                    <input v-if="searchWords1.length > 0" @click="checkAll(1)" 
                                        :checked="checkedWords1Idxs.length === searchWords1.length" type="checkbox">
                                </th>
                                <th>NO.</th>
                                <th>단어1</th>
                                <th></th>
                            </tr>
                        </thead>
                        <!--endregion-->
                        <tbody>
                            <template v-if="searchWords1.length > 0">
                                <tr v-for="word in searchWords1">
                                    <td><input @click="checkWord(1, word.wordIdx)" :checked="checkedWords1Idxs.indexOf(word.wordIdx) > -1" type="checkbox"></td>
                                    <td>{{word.wordIdx}}</td>
                                    <td>{{word.word}}</td>
                                    <td>
                                        <button @click="$emit('modifyWord', 1, word)" class="linker-btn">수정</button>
                                        <button @click="deleteWord(1, word)" class="linker-btn">삭제</button>
                                    </td>
                                </tr>
                            </template>
                            <tr v-else>
                                <td colspan="4">단어가 없습니다.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <!--endregion-->

                <!--region 단어2(명사)-->
                <div class="modal-nicknames-content">
                    <div class="nicknames-btn-area">
                        <button @click="openPostModal(2)" class="linker-btn">신규등록</button>
                        <button @click="deleteWords(2)" class="linker-btn">삭제</button>
                    </div>

                    <table class="forum-list-tbl">
                        <!--region colgroup-->
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
                                    <input v-if="searchWords2.length > 0" @click="checkAll(2)" 
                                        :checked="checkedWords2Idxs.length === searchWords2.length" type="checkbox">
                                </th>
                                <th>NO.</th>
                                <th>단어2</th>
                                <th></th>
                            </tr>
                        </thead>
                        <!--endregion-->
                        <tbody>
                            <template v-if="searchWords2.length > 0">
                                <tr v-for="word in searchWords2">
                                    <td><input @click="checkWord(2, word.wordIdx)" :checked="checkedWords2Idxs.indexOf(word.wordIdx) > -1" type="checkbox"></td>
                                    <td>{{word.wordIdx}}</td>
                                    <td>{{word.word}}</td>
                                    <td>
                                        <button @click="$emit('modifyWord', 2, word)"class="linker-btn">수정</button>
                                        <button @click="deleteWord(2, word)" class="linker-btn">삭제</button>
                                    </td>
                                </tr>
                            </template>
                            <tr v-else>
                                <td colspan="4">단어가 없습니다.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <!--endregion-->
            </div>
        </div>
    `,
    data() {return {
        searchWordNumber : 1, // 검색할 단어 번호
        keyword : '', // 검색 키워드
        checkedWords1Idxs : [], // 체크한 단어1 일련번호 리스트
        checkedWords2Idxs : [], // 체크한 단어2 일련번호 리스트
    }},
    computed : {
        //region searchWords1 단어1 검색 결과 리스트
        searchWords1() {
            const keyword = this.keyword.trim();
            if( this.searchWordNumber === 1 && keyword ) {
                return this.words1.filter(w => w.word.indexOf(keyword) > -1);
            } else {
                return this.words1;
            }
        },
        //endregion
        //region searchWords2 단어2 검색 결과 리스트
        searchWords2() {
            const keyword = this.keyword.trim();
            if( this.searchWordNumber === 2 && keyword ) {
                return this.words2.filter(w => w.word.indexOf(keyword) > -1);
            } else {
                return this.words2;
            }
        },
        //endregion
    },
    props : {
        words1 : { type:Array, default:function(){return [];} }, // 단어1(형용사) 리스트
        words2 : { type:Array, default:function(){return [];} }, // 단어2(명사) 리스트
    },
    methods : {
        //region updateKeyword 검색 키워드 수정
        updateKeyword() {
            this.clearCheck();
            this.keyword = document.getElementById('keyword').value;
        },
        //endregion
        //region openPostModal 등록 모달 열기
        openPostModal(num) {
            this.$emit('openPostModal', num);
        },
        //endregion
        //region checkAll 전체 선택
        checkAll(num) {
            if( num === 1 ) {
                if( this.checkedWords1Idxs.length !== this.searchWords1.length ) {
                    this.checkedWords1Idxs = this.searchWords1.map(w => w.wordIdx);
                } else {
                    this.checkedWords1Idxs = [];
                }
            } else {
                if( this.checkedWords2Idxs.length !== this.searchWords2.length ) {
                    this.checkedWords2Idxs = this.searchWords2.map(w => w.wordIdx);
                } else {
                    this.checkedWords2Idxs = [];
                }
            }
        },
        //endregion
        //region checkWord 단어 선택
        checkWord(num, wordIdx) {
            if( num === 1 ) {
                if( this.checkedWords1Idxs.indexOf(wordIdx) > -1 ) {
                    this.checkedWords1Idxs.splice(this.checkedWords1Idxs.findIndex(w => w === wordIdx), 1);
                } else {
                    this.checkedWords1Idxs.push(wordIdx);
                }
            } else {
                if( this.checkedWords2Idxs.indexOf(wordIdx) > -1 ) {
                    this.checkedWords2Idxs.splice(this.checkedWords2Idxs.findIndex(w => w === wordIdx), 1);
                } else {
                    this.checkedWords2Idxs.push(wordIdx);
                }
            }
        },
        //endregion
        //region clearCheck 체크 모두 해제
        clearCheck() {
            this.checkedWords1Idxs = [];
            this.checkedWords2Idxs = [];
        },
        //endregion
        //region deleteWord 단어 하나 삭제
        deleteWord(num, word) {
            if( confirm(`'${word.word}' 이 단어를 삭제하시겠습니까?`) )
                this.$emit('deleteWords', num, [word.wordIdx]);
        },
        //endregion
        //region deleteWords 단어 여러개 삭제
        deleteWords(num) {
            const checkWords = num === 1 ? this.checkedWords1Idxs : this.checkedWords2Idxs;
            if( checkWords.length > 0 && confirm('선택한 단어들을 삭제하시겠습니까?') )
                this.$emit('deleteWords', num, checkWords);
        },
        //endregion
    }
});