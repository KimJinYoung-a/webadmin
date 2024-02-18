Vue.component('MANAGE-NICKNAME-SLANG', {
    template : `
        <div>
            <div class="search">
                <div>
                    <div class="search-group">
                        <label>��Ӿ�:</label>
                        <input id="keyword" @keydown.enter="updateKeyword" type="text">
                    </div>
                    <button @click="updateKeyword" class="linker-btn">�˻�</button>
                </div>
            </div>

            <div>
                <div class="nicknames-btn-area">
                    <button @click="$emit('openPostModal', 3)" class="linker-btn">�űԵ��</button>
                    <button @click="deleteWords" class="linker-btn">����</button>
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
                            <th>��Ӿ�</th>
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
                                    <button @click="$emit('modifyWord', 3, word)" class="linker-btn">����</button>
                                    <button @click="deleteWord(word)" class="linker-btn">����</button>
                                </td>
                            </tr>
                        </template>
                        <tr v-else>
                            <td colspan="4">��Ӿ �����ϴ�.</td>
                        </tr>
                    </tbody>
                </table>
            </div>

        </div>
    `,
    data() {return {
        keyword : '', // �˻� Ű����
        checkedWordIdxs : [], // üũ�� �ܾ� �Ϸù�ȣ ����Ʈ
    }},
    computed : {
        //region searchWords �ܾ� �˻� ��� ����Ʈ
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
        //region updateKeyword �˻� Ű���� ����
        updateKeyword() {
            this.checkedWordIdxs = [];
            this.keyword = document.getElementById('keyword').value;
        },
        //endregion
        //region checkAll ��ü ����
        checkAll() {
            if( this.checkedWordIdxs.length !== this.searchWords.length ) {
                this.checkedWordIdxs = this.searchWords.map(w => w.wordIdx);
            } else {
                this.checkedWordIdxs = [];
            }
        },
        //endregion
        //region checkWord �ܾ� ����
        checkWord(wordIdx) {
            if( this.checkedWordIdxs.indexOf(wordIdx) > -1 ) {
                this.checkedWordIdxs.splice(this.checkedWordIdxs.findIndex(w => w === wordIdx), 1);
            } else {
                this.checkedWordIdxs.push(wordIdx);
            }
        },
        //endregion
        //region deleteWord �ܾ� �ϳ� ����
        deleteWord(word) {
            if( confirm(`'${word.word}' �� �ܾ �����Ͻðڽ��ϱ�?`) )
                this.$emit('deleteWords', 3, [word.wordIdx]);
        },
        //endregion
        //region deleteWords �ܾ� ������ ����
        deleteWords() {
            if( this.checkedWordIdxs.length > 0 && confirm('������ �ܾ���� �����Ͻðڽ��ϱ�?') )
                this.$emit('deleteWords', 3, this.checkedWordIdxs);
        },
        //endregion
    }
});