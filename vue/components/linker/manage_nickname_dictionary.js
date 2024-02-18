Vue.component('MANAGE-NICKNAME-DICTIONARY', {
    template : `
        <div>
            <!--region �ܾ� �˻�-->
            <div class="search">
                <div>
                    <div class="search-group">
                        <select v-model.number="searchWordNumber">
                            <option value="1">�ܾ�1</option>
                            <option value="2">�ܾ�2</option>
                        </select>
                        :
                        <input id="keyword" @keydown.enter="updateKeyword" type="text">
                    </div>
                    <button @click="updateKeyword" class="linker-btn">�˻�</button>
                </div>
            </div>
            <!--endregion-->

            <div class="modal-nicknames-area">
                <!--region �ܾ�1-->
                <div class="modal-nicknames-content">
                    <div class="nicknames-btn-area">
                        <button @click="openPostModal(1)" class="linker-btn">�űԵ��</button>
                        <button @click="deleteWords(1)" class="linker-btn">����</button>
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
                                <th>�ܾ�1</th>
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
                                        <button @click="$emit('modifyWord', 1, word)" class="linker-btn">����</button>
                                        <button @click="deleteWord(1, word)" class="linker-btn">����</button>
                                    </td>
                                </tr>
                            </template>
                            <tr v-else>
                                <td colspan="4">�ܾ �����ϴ�.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <!--endregion-->

                <!--region �ܾ�2(���)-->
                <div class="modal-nicknames-content">
                    <div class="nicknames-btn-area">
                        <button @click="openPostModal(2)" class="linker-btn">�űԵ��</button>
                        <button @click="deleteWords(2)" class="linker-btn">����</button>
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
                                <th>�ܾ�2</th>
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
                                        <button @click="$emit('modifyWord', 2, word)"class="linker-btn">����</button>
                                        <button @click="deleteWord(2, word)" class="linker-btn">����</button>
                                    </td>
                                </tr>
                            </template>
                            <tr v-else>
                                <td colspan="4">�ܾ �����ϴ�.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <!--endregion-->
            </div>
        </div>
    `,
    data() {return {
        searchWordNumber : 1, // �˻��� �ܾ� ��ȣ
        keyword : '', // �˻� Ű����
        checkedWords1Idxs : [], // üũ�� �ܾ�1 �Ϸù�ȣ ����Ʈ
        checkedWords2Idxs : [], // üũ�� �ܾ�2 �Ϸù�ȣ ����Ʈ
    }},
    computed : {
        //region searchWords1 �ܾ�1 �˻� ��� ����Ʈ
        searchWords1() {
            const keyword = this.keyword.trim();
            if( this.searchWordNumber === 1 && keyword ) {
                return this.words1.filter(w => w.word.indexOf(keyword) > -1);
            } else {
                return this.words1;
            }
        },
        //endregion
        //region searchWords2 �ܾ�2 �˻� ��� ����Ʈ
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
        words1 : { type:Array, default:function(){return [];} }, // �ܾ�1(�����) ����Ʈ
        words2 : { type:Array, default:function(){return [];} }, // �ܾ�2(���) ����Ʈ
    },
    methods : {
        //region updateKeyword �˻� Ű���� ����
        updateKeyword() {
            this.clearCheck();
            this.keyword = document.getElementById('keyword').value;
        },
        //endregion
        //region openPostModal ��� ��� ����
        openPostModal(num) {
            this.$emit('openPostModal', num);
        },
        //endregion
        //region checkAll ��ü ����
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
        //region checkWord �ܾ� ����
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
        //region clearCheck üũ ��� ����
        clearCheck() {
            this.checkedWords1Idxs = [];
            this.checkedWords2Idxs = [];
        },
        //endregion
        //region deleteWord �ܾ� �ϳ� ����
        deleteWord(num, word) {
            if( confirm(`'${word.word}' �� �ܾ �����Ͻðڽ��ϱ�?`) )
                this.$emit('deleteWords', num, [word.wordIdx]);
        },
        //endregion
        //region deleteWords �ܾ� ������ ����
        deleteWords(num) {
            const checkWords = num === 1 ? this.checkedWords1Idxs : this.checkedWords2Idxs;
            if( checkWords.length > 0 && confirm('������ �ܾ���� �����Ͻðڽ��ϱ�?') )
                this.$emit('deleteWords', num, checkWords);
        },
        //endregion
    }
});