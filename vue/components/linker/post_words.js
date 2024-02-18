Vue.component('POST-WORDS', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:120px;">
                    <col>
                </colgroup>
                <tbody>
                    <tr>
                        <th>{{wordNumber < 3 ? '단어'+wordNumber : '비속어'}}</th>
                        <td>
                            <textarea v-model="word" rows="3"></textarea>
                            <p class="descr">여러 단어를 추가할 경우 ',(쉼표)' 또는 Enter로 구분하여 입력해주세요.</p>
                            <p class="descr">Space(공백)는 무시됩니다.</p>
                        </td>
                    </tr>
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveNicknames" class="linker-btn">저장</button>
            </div>
        </div>
    `,
    data() {return {
        word : '', // 단어
    }},
    computed : {
        //region words 닉네임 배열
        words() {
            if( !this.word )
                return [];

            return this.word.replaceAll(/\n/gi, ',')
                .split(',')
                .map(w => w.trim())
                .filter(w => w !== '');
        },
        //endregion
    },
    props : {
        wordNumber : { type:Number, default:1 }, // 단어 번호
    },
    methods : {
        //region saveNicknames 닉네임 저장
        saveNicknames() {
            let message = '아래와 같은 단어들이 등록됩니다.\n저장하시겠습니까?\n';
            this.words.forEach(w => message += '\n' + w);
            if( confirm(message) )
                this.$emit('saveNicknames', this.words);
        },
        //endregion
    }
});