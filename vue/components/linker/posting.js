Vue.component('POSTING', {
    template : `
        <tr>
            <td><input @click="checkPosting($event)" type="checkbox" :checked="posting.postingIdx === checkedPostingIdx" :disabled="posting.fixed"></td>
            <td>{{posting.postingIdx}}</td>
            <td>{{fullCreatorType}} / {{creatorDescription}} / {{posting.creatorNickname}}</td>
            <td @click="$emit('clickPosting', posting)" class="posting-content">{{cutOutPostingContent}}</td>
            <td :class="{'posting-red' : posting.fixed}">{{posting.fixed ? 'Y' : 'N'}}</td>
            <td>{{getLocalDateTimeFormat(posting.regDate, 'yyyy-MM-dd HH:mm:ss')}}</td>
            <td>
                <button @click="$emit('clickPosting', posting)" class="linker-btn">����</button>
                <button @click="$emit('deletePosting', posting.postingIdx)" class="linker-btn">����</button>
            </td>
        </tr>
    `,
    props : {
        //region posting ������
        posting : {
            postingIdx : { type:Number, default:0 }, // ������ �Ϸù�ȣ
            creatorType : { type:String, default:'N' }, // ������ �ۼ��� ȸ�� ���� ��
            creatorDescr : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorLevelName : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ���
            creatorNickname : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            postingContents : { type:Number, default:0 }, // ������ ����
            fixed : { type:Boolean, default:false }, // ������ ���� ����
        },
        //endregion
        checkedPostingIdx : { type:Number, default:0 }, // ������ ������ �Ϸù�ȣ
    },
    computed : {
        //region cutOutPostingContent ������ ���� �ڸ� ��
        cutOutPostingContent() {
            const contents = this.posting.postingContents;
            return contents.length > 60 ? contents.substr(0, 60) + '...' : contents;
        },
        //endregion
        //region fullCreatorType ������ �ۼ��� ȸ�� ���� �� Ǯ����
        fullCreatorType() {
            switch (this.posting.creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region creatorDescription ������ �ۼ��� ����
        creatorDescription() {
            if( this.posting.creatorType === 'H' || this.posting.creatorType === 'G' ) {
                return this.posting.creatorDescr;
            } else {
                return this.posting.creatorLevelName;
            }
        },
        //endregion
    },
    methods: {
        //region checkPosting ������ üũ�ڽ� üũ
        checkPosting(e) {
            this.$emit('checkPosting', e, this.posting.postingIdx);
        },
        //endregion
    }
})