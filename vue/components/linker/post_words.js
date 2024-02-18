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
                        <th>{{wordNumber < 3 ? '�ܾ�'+wordNumber : '��Ӿ�'}}</th>
                        <td>
                            <textarea v-model="word" rows="3"></textarea>
                            <p class="descr">���� �ܾ �߰��� ��� ',(��ǥ)' �Ǵ� Enter�� �����Ͽ� �Է����ּ���.</p>
                            <p class="descr">Space(����)�� ���õ˴ϴ�.</p>
                        </td>
                    </tr>
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveNicknames" class="linker-btn">����</button>
            </div>
        </div>
    `,
    data() {return {
        word : '', // �ܾ�
    }},
    computed : {
        //region words �г��� �迭
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
        wordNumber : { type:Number, default:1 }, // �ܾ� ��ȣ
    },
    methods : {
        //region saveNicknames �г��� ����
        saveNicknames() {
            let message = '�Ʒ��� ���� �ܾ���� ��ϵ˴ϴ�.\n�����Ͻðڽ��ϱ�?\n';
            this.words.forEach(w => message += '\n' + w);
            if( confirm(message) )
                this.$emit('saveNicknames', this.words);
        },
        //endregion
    }
});