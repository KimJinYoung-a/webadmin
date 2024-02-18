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
                        <th>ȸ������</th>
                        <td class="content">{{fullCreatorType}}</td>
                    </tr>
                    <tr v-if="posting.creatorType === 'H' || posting.creatorType === 'G'">
                        <th>ȸ������</th>
                        <td class="content">{{posting.creatorDescr}}</td>
                    </tr>
                    <tr v-else>
                        <th>ȸ�����</th>
                        <td class="content">{{posting.creatorLevelName}}</td>
                    </tr>
                    <tr>
                        <th>�г���</th>
                        <td class="content">{{posting.creatorNickname}}</td>
                    </tr>
                    <tr>
                        <th>�ۼ�����</th>
                        <td><textarea v-model="tempContent" rows="9"></textarea></td>
                    </tr>
                    <tr>
                        <th>����������</th>
                        <td>
                            <button v-if="posting.linkValue" @click="clickLink" class="linker-btn">
                                {{linkButtonText}}
                            </button>
                        </td>
                    </tr>
                    <tr>
                        <th>��� ���� ����</th>
                        <td>
                            <input @click="clickFixRadio($event, true)" id="fixPostingY" type="radio" name="fix" :checked="fix.use">
                            <label for="fixPostingY">����</label>
                            <input @click="clickFixRadio($event, false)" id="fixPostingN" type="radio" name="fix" :checked="!fix.use">
                            <label for="fixPostingN">��������</label>
                        </td>
                    </tr>
                    <tr>
                        <th>�ۼ��Ͻ�</th>
                        <td class="content">
                            <strong>{{getLocalDateTimeFormat(posting.regDate, 'yyyy-MM-dd HH:mm:ss')}}</strong>
                            <span v-if="posting.updateDate" class="posting-update">
                                {{getLocalDateTimeFormat(posting.updateDate, 'yyyy-MM-dd HH:mm:ss')}} ����
                            </span>
                        </td>
                    </tr>
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="savePosting" class="linker-btn">����</button>
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
        //region posting ������
        posting: {
            postingIdx : { type:Number, default:0 }, // ������ �Ϸù�ȣ
            creatorType : { type:String, default:'N' }, // ������ �ۼ��� ȸ�� ���� ��
            creatorDescr : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorLevelName : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ���
            creatorNickname : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            postingContents : { type:Number, default:0 }, // ������ ����
            linkType : { type:Number, default:null }, // ������ ��ũ ����
            linkValue : { type:String, default:'' }, // ������ ��ũ ��
            fixed : { type:Boolean, default:false }, // ������ ���� ����
            regDate : { type:String, default:'' }, // ������ ���� ����
        },
        //endregion
    },
    computed : {
        //region fullCreatorType ������ �ۼ��� ȸ�� ���� �� Ǯ����
        fullCreatorType() {
            switch (this.posting.creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region linkButtonText ��ũ ��ư �ؽ�Ʈ
        linkButtonText() {
            let linkTitle;
            switch( this.posting.linkType ) {
                case 1: linkTitle = '��ǰ'; break;
                case 2: linkTitle = '�̺�Ʈ'; break;
            }

            return (linkTitle ? (linkTitle + ' : ') : '') + this.posting.linkValue;
        },
        //endregion
    },
    methods : {
        //region clickLink ��ũ Ŭ��
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
        //region clickFixRadio ���� ���� Ŭ��
        clickFixRadio(e, fixFlag) {
            if( fixFlag ) {
                e.preventDefault();
                this.$emit('openFixPostingModal', this.fix);
            } else {
                this.clearFix();
            }
        },
        //endregion
        //region savePosting ������ ����
        savePosting() {
            if( confirm('���� �Ͻðڽ��ϱ�?') )
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
        //region setFixPosting ������ �������� Set
        setFixPosting(fix) {
            this.fix = {
                positionNo : fix.positionNo,
                startDate : fix.startDate,
                endDate : fix.endDate,
                use : true
            }
        },
        //endregion
        //region clearFix ������ �������� ����
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