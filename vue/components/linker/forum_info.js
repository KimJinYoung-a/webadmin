Vue.component('FORUM-INFO', {
    template : `
        <div class="forum-info">
            <div class="title">
                <div>
                    <h3>포럼 안내</h3>
                    <span>포럼 안내는 5개까지만 등록할 수 있습니다.</span>
                </div>
                <div>
                    <button @click="$emit('postForumInfo')" class="linker-btn">포럼 안내 등록</button>
                    <button @click="modifySort" class="linker-btn">정렬수정</button>
                    <button @click="deleteInfos" class="linker-btn">선택 항목 삭제</button>
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
                        <th>노출순서</th>
                        <th>안내제목</th>
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
                        <td colspan="5">등록된 안내 정보가 없습니다.</td>
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
        //region infos 안내 리스트
        infos : {
            infoIdx : { type:Number, default:0 }, // 안내 일련번호
            sortNo : { type:Number, default:0 }, // 정렬번호
            appTitle : { type:String, default:'' }, // 제목 - APP
            appContent : { type:String, default:'' }, // 내용 - APP
            mobileTitle : { type:String, default:'' }, // 제목 - Mobile
            mobileContent : { type:String, default:'' }, // 내용 - Mobile
            pcTitle : { type:String, default:'' }, // 제목 - PC
            pcContent : { type:String, default:'' } // 내용 - PC
        },
        //endregion
    },
    methods : {
        //region setTempInfos Set 임시 안내 리스트
        setTempInfos(infos) {
            if( infos )
                this.tempInfos = infos;
            else
                this.tempInfos = this.infos;
        },
        //endregion
        //region checkAll 전체 항목 선택/삭제
        checkAll(e) {
            if( e.target.checked )
                this.checkedInfos = this.infos.map(i => i.infoIdx);
            else
                this.checkedInfos = [];
        },
        //endregion
        //region checkInfo 안내 check 추가/해제
        checkInfo(infoIdx, e) {
            if( e.target.checked ) {
                this.checkedInfos.push(infoIdx);
            } else {
                document.getElementById('forumInfoAll').checked = false;
                this.checkedInfos.splice(this.checkedInfos.findIndex(i => i === infoIdx), 1);
            }
        },
        //endregion
        //region deleteInfos 선택 항목 삭제
        deleteInfos() {
            if( this.checkedInfos.length > 0 && confirm('선택한 항목들을 삭제하시겠습니까?') )
                this.$emit('deleteInfos', this.checkedInfos);
        },
        //endregion
        //region modifySort 정렬 수정
        modifySort() {
            if( this.tempInfos.length > 0 && confirm('정렬을 수정하시겠습니까?') ) {
                const idxs = this.tempInfos.map(i => i.infoIdx);
                this.$emit('modifySort', idxs);
            }
        },
        //endregion
    }
});