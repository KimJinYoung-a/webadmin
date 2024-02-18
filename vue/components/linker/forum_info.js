Vue.component('FORUM-INFO', {
    template : `
        <div class="forum-info">
            <div class="title">
                <div>
                    <h3>���� �ȳ�</h3>
                    <span>���� �ȳ��� 5�������� ����� �� �ֽ��ϴ�.</span>
                </div>
                <div>
                    <button @click="$emit('postForumInfo')" class="linker-btn">���� �ȳ� ���</button>
                    <button @click="modifySort" class="linker-btn">���ļ���</button>
                    <button @click="deleteInfos" class="linker-btn">���� �׸� ����</button>
                </div>
            </div>

            <table id="forumInfoTbl" class="forum-list-tbl">
                <!--region colgroup-->
                <colgroup>
                    <col style="width: 50px;">
                    <col style="width: 100px;">
                    <col style="width: 100px;">
                    <col style="width: 300px;">
                    <col>
                </colgroup>
                <!--endregion-->
                <!--region THead-->
                <thead>
                    <tr>
                        <th><input @click="checkAll($event)" id="forumInfoAll" type="checkbox"></th>
                        <th>ID</th>
                        <th>�������</th>
                        <th>�ȳ�����</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <draggable v-if="tempInfos.length > 0" v-model="tempInfos" tag="tbody">
                    <tr v-for="info in tempInfos">
                        <td><input @click="checkInfo(info.infoIdx, $event)" :checked="checkedInfos.indexOf(info.infoIdx) > -1" type="checkbox"></td>
                        <td>{{info.infoIdx}}</td>
                        <td>{{info.sortNo}}</td>
                        <td @click="$emit('postForumInfo', info)" v-html="info.appTitle" class="tl info-title" colspan="2"></td>                        
                    </tr>
                </draggable>
                <tbody v-else>
                    <tr>
                        <td colspan="5">��ϵ� �ȳ� ������ �����ϴ�.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    mounted() {
        this.setTempInfos();
    },
    data() {return {
        checkedInfos : [],
        tempInfos : [],
    }},
    props : {
        //region infos �ȳ� ����Ʈ
        infos : {
            infoIdx : { type:Number, default:0 }, // �ȳ� �Ϸù�ȣ
            sortNo : { type:Number, default:0 }, // ���Ĺ�ȣ
            appTitle : { type:String, default:'' }, // ���� - APP
            appContent : { type:String, default:'' }, // ���� - APP
            mobileTitle : { type:String, default:'' }, // ���� - Mobile
            mobileContent : { type:String, default:'' }, // ���� - Mobile
            pcTitle : { type:String, default:'' }, // ���� - PC
            pcContent : { type:String, default:'' } // ���� - PC
        },
        //endregion
    },
    methods : {
        //region setTempInfos Set �ӽ� �ȳ� ����Ʈ
        setTempInfos(infos) {
            if( infos )
                this.tempInfos = infos;
            else
                this.tempInfos = this.infos;
        },
        //endregion
        //region checkAll ��ü �׸� ����/����
        checkAll(e) {
            if( e.target.checked )
                this.checkedInfos = this.infos.map(i => i.infoIdx);
            else
                this.checkedInfos = [];
        },
        //endregion
        //region checkInfo �ȳ� check �߰�/����
        checkInfo(infoIdx, e) {
            if( e.target.checked ) {
                this.checkedInfos.push(infoIdx);
            } else {
                document.getElementById('forumInfoAll').checked = false;
                this.checkedInfos.splice(this.checkedInfos.findIndex(i => i === infoIdx), 1);
            }
        },
        //endregion
        //region deleteInfos ���� �׸� ����
        deleteInfos() {
            if( this.checkedInfos.length > 0 && confirm('������ �׸���� �����Ͻðڽ��ϱ�?') )
                this.$emit('deleteInfos', this.checkedInfos);
        },
        //endregion
        //region modifySort ���� ����
        modifySort() {
            if( this.tempInfos.length > 0 && confirm('������ �����Ͻðڽ��ϱ�?') ) {
                const idxs = this.tempInfos.map(i => i.infoIdx);
                this.$emit('modifySort', idxs);
            }
        },
        //endregion
    }
});