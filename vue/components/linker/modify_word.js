Vue.component('MODIFY-WORD', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:120px;">
                    <col>
                </colgroup>
                <tbody>
                    <tr>
                        <th class="mid">단어{{wordNumber}}</th>
                        <td>
                            <input v-model="modifyWord" type="text">
                        </td>
                    </tr>
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="save" class="linker-btn">저장</button>
            </div>
        </div>
    `,
    mounted() {
        this.modifyWord = this.word.word;
    },
    data() {return {
        modifyWord : '', // 단어
    }},
    props : {
        wordNumber : { type:Number, default:1 }, // 단어 번호
        word : {
            wordIdx : { type:Number, default:1 }, // 단어 일련번호
            word : { type:String, default:'' }, // 단어
        }
    },
    methods : {
        //region save 저장
        save() {
            const message = `'${this.word.word}' ⇒ '${this.modifyWord}'\n저장하시겠습니까?`;
            if( confirm(message) )
                this.$emit('save', this.word.wordIdx, this.modifyWord);
        },
        //endregion
    }
});