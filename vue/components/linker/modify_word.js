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
                        <th class="mid">�ܾ�{{wordNumber}}</th>
                        <td>
                            <input v-model="modifyWord" type="text">
                        </td>
                    </tr>
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="save" class="linker-btn">����</button>
            </div>
        </div>
    `,
    mounted() {
        this.modifyWord = this.word.word;
    },
    data() {return {
        modifyWord : '', // �ܾ�
    }},
    props : {
        wordNumber : { type:Number, default:1 }, // �ܾ� ��ȣ
        word : {
            wordIdx : { type:Number, default:1 }, // �ܾ� �Ϸù�ȣ
            word : { type:String, default:'' }, // �ܾ�
        }
    },
    methods : {
        //region save ����
        save() {
            const message = `'${this.word.word}' �� '${this.modifyWord}'\n�����Ͻðڽ��ϱ�?`;
            if( confirm(message) )
                this.$emit('save', this.word.wordIdx, this.modifyWord);
        },
        //endregion
    }
});