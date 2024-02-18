Vue.component('MANAGE-REPORT-POSTINGS', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{numberFormat(postingCount)}}</span>��</p>
                <div>
                    <button @click="unBlockSelectedPostings" class="linker-btn">��� ����</button>
                    <button @click="deleteSelectedPostings" class="linker-btn">���� ������ ����</button>
                </div>
            </div>

            <table class="forum-list-tbl">
                <!--region colgroup-->
                <colgroup>
                    <col style="width: 30px;">
                    <col style="width: 60px;">
                    <col style="width: 80px;">
                    <col style="width: 170px;">
                    <col>
                    <col style="width: 170px;">
                    <col style="width: 200px;">
                </colgroup>
                <!--endregion-->
                <!--region THead-->
                <thead>
                    <tr>
                        <th><input v-if="postings.length > 0" @click="checkAll" :checked="checkedPostingIdxs.length === postings.length" type="checkbox"></th>
                        <th>idx</th>
                        <th>������ �̹���</th>
                        <th>�ۼ��� ����</th>
                        <th>�ۼ�����</th>
                        <th>����� ������</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <tbody>
                    <template v-if="postings && postings.length > 0">
                        <tr v-for="posting in postings">
                            <td><input type="checkbox" @click="checkPosting(posting.postingIdx)" :checked="checkedPostingIdxs.indexOf(posting.postingIdx) > -1"></td>
                            <td>{{posting.postingIdx}}</td>
                            <td><img :src="decodeBase64(posting.creatorThumbnail)" class="thumb"></td>
                            <td>
                                {{fullCreatorType(posting.creatorType)}} 
                                / {{creatorDescription(posting.creatorType, posting.creatorDescr, posting.creatorLevelName)}} 
                                / {{posting.creatorNickname}}
                            </td>
                            <td>{{cutoutPostingContents(posting.postingContents)}}</td>
                            <td>
                                <button v-if="posting.linkValue" @click="clickLink(posting.linkType, posting.linkValue)" class="linker-btn link">
                                    {{linkButtonText(posting.linkType, posting.linkValue)}}
                                </button>
                            </td>
                            <td>
                                <button @click="unBlockPostings([posting.postingIdx])" class="linker-btn long">��� ����</button>
                                <button @click="deletePosting(posting.postingIdx)" class="linker-btn long">������ ����</button>
                            </td>
                        </tr>
                    </template>
                    <tr v-else>
                        <td colspan="7">�Ű�� �������� �������� �ʽ��ϴ�.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    data() {return {
        checkedPostingIdxs : [], // üũ�� ������ �Ϸù�ȣ ����Ʈ
    }},
    props : {
        postingCount : { type:Number, default:0 }, // ������ ����
        //region postings ������ ����Ʈ
        postings : {
            postingIdx : { type:Number, default:0 }, // ������ �Ϸù�ȣ
            creatorType : { type:String, default:'N' }, // ������ �ۼ��� ȸ�� ���� ��
            creatorDescr : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorLevelName : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ���
            creatorNickname : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorThumbnail : { type:String, default:'' }, // ������ �ۼ��� ����� �̹���
            postingContents : { type:Number, default:0 }, // ������ ����
            linkType : { type:String, default:'' }, // ��ũ ����
            linkValue : { type:String, default:'' }, // ��ũ ��
        },
        //endregion
    },
    methods: {
        //region fullCreatorType ������ �ۼ��� ȸ�� ���� �� Ǯ����
        fullCreatorType(creatorType) {
            switch (creatorType) {
                case 'H': return 'Host';
                case 'G': return 'Guest';
                default: return 'User';
            }
        },
        //endregion
        //region creatorDescription ������ �ۼ��� ����
        creatorDescription(type, description, levelName) {
            if( type === 'H' || type === 'G' ) {
                return description;
            } else {
                return levelName;
            }
        },
        //endregion
        //region cutoutPostingContents ������ ���� 60�� �̳��� �ڸ�
        cutoutPostingContents(contents) {
            if( contents.length > 60 )
                return contents.substr(0, 60) + '...';
            else
                return contents;
        },
        //endregion
        //region linkButtonText ��ũ ��ư �ؽ�Ʈ
        linkButtonText(linkType, linkValue) {
            switch( linkType ) {
                case 1: return '��ǰ : ' + linkValue;
                case 2: return '�̺�Ʈ : ' + linkValue;
                default: return '�ܺ� URL';
            }
        },
        //endregion
        //region clickLink ��ũ Ŭ��
        clickLink(linkType, linkValue) {
            let url;
            const wwwUrl = this.isDevelop ? '//2015www.10x10.co.kr' : '//www.10x10.co.kr';

            if( linkType === 1 )
                url = wwwUrl + '/shopping/category_prd.asp?itemid=' + linkValue;
            else if( linkType === 2 )
                url = wwwUrl + '/event/eventmain.asp?eventid=' + linkValue;
            else
                url = linkValue;

            window.open(url, '_blank');
        },
        //endregion
        //region deletePosting ������ ����
        deletePosting(postingIdx) {
            this.$emit('deletePosting', postingIdx, true);
        },
        //endregion
        //region unBlockPostings ������ ��� ����
        unBlockPostings(postingIdxs) {
            if( confirm('�� �������� ����� �����Ͻðڽ��ϱ�?\n��� �Ű����� �ʱ�ȭ�˴ϴ�.') )
                this.$emit('unBlockPostings', postingIdxs);
        },
        //endregion
        //region checkAll ��ü ������ ����
        checkAll() {
            if( this.checkedPostingIdxs.length !== this.postings.length ) {
                this.checkedPostingIdxs = this.postings.map(p => p.postingIdx);
            } else {
                this.checkedPostingIdxs = [];
            }
        },
        //endregion
        //region checkPosting üũ/���� ������
        checkPosting(postingIdx) {
            if( this.checkedPostingIdxs.indexOf(postingIdx) > -1 ) {
                this.checkedPostingIdxs.splice(this.checkedPostingIdxs.findIndex(p => p === postingIdx), 1);
            } else {
                this.checkedPostingIdxs.push(postingIdx);
            }
        },
        //endregion
        //region unBlockSelectedPostings ���� �����õ� ��� ����
        unBlockSelectedPostings() {
            if( this.checkedPostingIdxs.length > 0 &&
                confirm('������ �����õ��� ����� �����Ͻðڽ��ϱ�?\n�ش� �����õ��� ��� �Ű����� �ʱ�ȭ�˴ϴ�.') ) {
                this.$emit('unBlockPostings', this.checkedPostingIdxs);
            }
        },
        //endregion
        //region deleteSelectedPostings ���� �����õ� ����
        deleteSelectedPostings() {
            if( this.checkedPostingIdxs.length > 0 &&
                confirm('������ �����õ��� �����Ͻðڽ��ϱ�?') ) {
                this.$emit('deletePostings', this.checkedPostingIdxs);
            }
        },
        //endregion
        //region numberFormat ���� õ�ڸ� (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
    }
});