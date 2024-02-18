Vue.component('MANAGE-FIX-POSTINGS', {
    template : `
        <div class="forum-posting-result">
            <div class="forum-posting-top">
                <p><span>{{postings.length}}</span>��</p>
                <div>
                    <button @click="clearFixPostings" class="linker-btn">���� ����</button>
                    <button @click="modifyFixPositions" class="linker-btn">������ġ ����</button>
                </div>
            </div>

            <table class="forum-list-tbl">
                <!--region Colgroup-->
                <colgroup>
                    <col style="width: 30px;">
                    <col style="width: 60px;">
                    <col style="width: 170px;">
                    <col>
                    <col style="width: 100px;">
                    <col style="width: 160px;">
                    <col style="width: 170px;">
                </colgroup>
                <!--endregion-->
                <!--region THead-->
                <thead>
                    <tr>
                        <th><input v-if="postings.length > 0" @click="checkAll" :checked="checkedPostingIdxs.length === postings.length" type="checkbox"></th>
                        <th>idx</th>
                        <th>�ۼ��� ����</th>
                        <th>�ۼ�����</th>
                        <th>���� ���� ��ġ</th>
                        <th>���� �Ⱓ</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <tbody>
                    <template v-if="postings && postings.length > 0">
                        <tr v-for="(posting, index) in postings">
                            <td><input type="checkbox" @click="checkPosting(posting.postingIdx)" :checked="checkedPostingIdxs.indexOf(posting.postingIdx) > -1"></td>
                            <td>{{posting.postingIdx}}</td>
                            <td>
                                {{fullCreatorType(posting.creatorType)}} 
                                / {{creatorDescription(posting.creatorType, posting.creatorDescr, posting.creatorLevelName)}} 
                                / {{posting.creatorNickname}}
                            </td>
                            <td>{{cutoutPostingContents(posting.postingContents)}}</td>
                            <td><input type="text" v-if="positionNumbers[index]" class="forum-sort" v-model.number="positionNumbers[index].positionNo"></td>
                            <td>{{fixPeriod(posting)}}</td>
                            <td>
                                <button @click="$emit('modifyPosting', posting)" class="linker-btn">����</button>
                                <button @click="clearFixPosting(posting.postingIdx)" class="linker-btn long">���� ����</button>
                            </td>
                        </tr>
                    </template>
                    <tr v-else>
                        <td colspan="7">���� �������� �������� �ʽ��ϴ�.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    data() {return {
        checkedPostingIdxs : [], // üũ�� ������ �Ϸù�ȣ ����Ʈ
        positionNumbers : [], // ���� ���� ����Ʈ
    }},
    props : {
        //region postings ���� ������ ����Ʈ
        postings : {
            postingIdx : { type:Number, default:0 }, // ������ �Ϸù�ȣ
            creatorType : { type:String, default:'N' }, // ������ �ۼ��� ȸ�� ���� ��
            creatorDescr : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorLevelName : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ���
            creatorNickname : { type:String, default:'' }, // ������ �ۼ��� ȸ�� ����
            creatorThumbnail : { type:String, default:'' }, // ������ �ۼ��� ����� �̹���
            postingContents : { type:Number, default:0 }, // ������ ����
            positionNo : { type:Number, default:0 }, // ���� ���� ��ġ
            startDate : { type:String, default:'' }, // ���� ������ ��������
            endDate : { type:String, default:'' }, // ���� ������ ��������
        },
        //endregion
    },
    computed : {
        modifyPositionApiData() {
            const data = {};
            if( this.positionNumbers.length > 0 ) {
                this.positionNumbers.forEach((p, index) => {
                    data[`postingsFixed[${index}].postingIndex`] = p.postingIndex;
                    data[`postingsFixed[${index}].positionNo`] = p.positionNo;
                });
            }
            return data;
        },
    },
    methods : {
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
        //region fixPeriod ���� �Ⱓ
        fixPeriod(posting) {
            return this.getLocalDateTimeFormat(posting.startDate, 'yyyy-MM-dd') + ' ~ '
                + this.getLocalDateTimeFormat(posting.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region clearFixPosting ������ �ϳ� ���� ����
        clearFixPosting(postingIdx) {
            if( confirm('�� �������� ������ ���� �Ͻðڽ��ϱ�?') )
                this.$emit('clearPostings', [postingIdx]);
        },
        //endregion
        //region clearFixPostings ������ �����õ� ���� ����
        clearFixPostings() {
            if( confirm('������ �����õ��� ������ ���� �Ͻðڽ��ϱ�?') )
                this.$emit('clearPostings', this.checkedPostingIdxs);
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
        //region modifyFixPositions ���� ��ġ ����
        modifyFixPositions() {
            if( confirm('���� ���·� ���� ��ġ�� �����Ͻðڽ��ϱ�?') )
                this.$emit('modifyFixPositions', this.modifyPositionApiData);
        },
        //endregion
    },
    watch : {
        postings(postings) {
            this.positionNumbers = postings.map(p => {
                return {
                    postingIndex: p.postingIdx,
                    positionNo: p.positionNo
                };
            });
        }
    }
});