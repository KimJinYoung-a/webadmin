Vue.component('MANAGE-ITEMS', {
    template : `
        <div class="manage-area">
            <div class="manage-button-area">
                <button @click="$emit('postItem')" class="add">신규등록</button>
                <button v-if="isUpdated" @click="saveSortAndSelected">저장</button>
            </div>
            <table>
                <!--region colgroup-->
                <colgroup>
                    <col style="width:50px;">
                    <col style="width:150px;">
                    <col style="width:150px;">
                    <col style="width:auto;">
                    <col style="width:80px;">
                    <col style="width:150px;">
                </colgroup>
                <!--endregion-->
                <!--region thead-->
                <thead>
                    <tr>
                        <th>순서</th>
                        <th>타이틀</th>
                        <th>서브타이틀</th>
                        <th>링크</th>
                        <th>초기값</th>
                        <th></th>
                    </tr>
                </thead>
                <!--endregion-->
                <draggable v-if="tempItems.length > 0" v-model="tempItems" tag="tbody" @change="changeSort">
                    <tr v-for="item in tempItems">
                        <td>{{item.sort}}</td>
                        <td>{{item.title}}</td>
                        <td>{{item.subTitle}}</td>
                        <td>{{decodeBase64(item.link) ? decodeBase64(item.link) : '-'}}</td>
                        <td><input v-model="tempSelectedIdx" :value="item.itemIndex" type="radio"></td>
                        <td>
                            <button @click="modifyItem(item)" class="add">수정</button>
                            <button @click="deleteItem(item.itemIndex)" class="add">삭제</button>
                        </td>
                    </tr>
                </draggable>
                <tbody v-else>
                    <tr>
                        <td colspan="5">등록된 아이템이 없습니다.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    mounted() {
        this.setTempItems(this.items);
    },
    data() {return {
        tempItems : [], // 수정중 아이템 리스트

        selectedIdx : -1, // 초기값
        tempSelectedIdx : -1, // 수정중 초기값
    }},
    props: {
        masterIndex : { type:Number, default:0 },
        //region items 아이템 리스트
        items : {
            itemIndex : { type:Number, default:0 },
            title : { type:String, default:'' },
            subTitle : { type:String, default:'' },
            link : { type:String, default:'' },
            image : { type:String, default:'' },
            sort : { type:Number, default:1 },
            selected : { type:Boolean, default:false },
        },
        //endregion
    },
    computed : {
        //region isUpdatedSort 순서가 변경되었는지 여부
        isUpdatedSort() {
            if( this.tempItems.length > 0 ) {
                for( let i=0 ; i<this.items.length ; i++ ) {
                    if( this.items[i].itemIndex !== this.tempItems[i].itemIndex ) {
                        return true;
                    }
                }
                return false;
            } else {
                return false;
            }
        },
        //endregion
        //region isUpdated 초기값 또는 순서가 변경되었는지 여부
        isUpdated() {
            return this.selectedIdx !== this.tempSelectedIdx || this.isUpdatedSort;
        },
        //endregion
    },
    methods : {
        //region deleteItem 아이템 삭제
        deleteItem(itemIndex) {
            if( confirm('삭제 하시겠습니까?') ) {
                this.callApi(2, 'POST', `/event/contents/tabbar/item/${itemIndex}/delete`, null, () => this.$emit('deleteItem'));
            }
        },
        //endregion
        //region modifyItem 아이템 수정
        modifyItem(item) {
            this.$emit('modifyItem', item);
        },
        //endregion
        //region setTempItems Set 수정중 아이템 리스트
        setTempItems(items) {
            this.tempItems = this.items;

            const selectedItem = this.tempItems.find(i => i.selected);
            if( selectedItem )
                this.selectedIdx = selectedItem.itemIndex;
            else
                this.selectedIdx = -1;

            this.tempSelectedIdx = this.selectedIdx;
        },
        //endregion
        //region changeSort 순서 변경
        changeSort(e) {
            const moved = e.moved;
            if( moved.oldIndex > moved.newIndex ) // 앞으로 이동
                this.moveForward(moved.element, moved.oldIndex+1, moved.newIndex+1);
            else // 뒤로 이동
                this.moveBack(moved.element, moved.oldIndex+1, moved.newIndex+1);
        },
        moveForward(item, oldSort, newSort) {
            this.tempItems.filter(i => i.sort >= newSort && i.sort < oldSort).forEach(i => i.sort++);
            item.sort = newSort;
        },
        moveBack(item, oldSort, newSort) {
            this.tempItems.filter(i => i.sort <= newSort && i.sort > oldSort).forEach(i => i.sort--);
            item.sort = newSort;
        },
        //endregion
        //region saveSortAndSelected 초기값, 순서변경 저장
        saveSortAndSelected() {
            if( !confirm('현재 상태로 저장하시겠습니까?') )
                return false;

            const url = '/event/contents/tabbar/sort/select/update';
            const data = this.createSaveSortAndSelectedData();
            this.callApi(2, 'POST', url, data, this.successSaveSortAndSelectedData);
        },
        createSaveSortAndSelectedData() {
            const data = {};
            this.tempItems.forEach((item, index) => {
                data[`items[${index}].itemIndex`] = item.itemIndex;
                data[`items[${index}].sort`] = item.sort;
                data[`items[${index}].selected`] = item.itemIndex === this.tempSelectedIdx;
            });
            return data;
        },
        successSaveSortAndSelectedData() {
            alert('저장 되었습니다.');
            this.$emit('saveSortAndSelected');
        },
        //endregion
    },
    watch : {
        items(items) {
            this.setTempItems(items);
        },
    }
});