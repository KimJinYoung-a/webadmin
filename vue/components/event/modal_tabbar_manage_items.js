Vue.component('MANAGE-ITEMS', {
    template : `
        <div class="manage-area">
            <div class="manage-button-area">
                <button @click="$emit('postItem')" class="add">�űԵ��</button>
                <button v-if="isUpdated" @click="saveSortAndSelected">����</button>
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
                        <th>����</th>
                        <th>Ÿ��Ʋ</th>
                        <th>����Ÿ��Ʋ</th>
                        <th>��ũ</th>
                        <th>�ʱⰪ</th>
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
                            <button @click="modifyItem(item)" class="add">����</button>
                            <button @click="deleteItem(item.itemIndex)" class="add">����</button>
                        </td>
                    </tr>
                </draggable>
                <tbody v-else>
                    <tr>
                        <td colspan="5">��ϵ� �������� �����ϴ�.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
    mounted() {
        this.setTempItems(this.items);
    },
    data() {return {
        tempItems : [], // ������ ������ ����Ʈ

        selectedIdx : -1, // �ʱⰪ
        tempSelectedIdx : -1, // ������ �ʱⰪ
    }},
    props: {
        masterIndex : { type:Number, default:0 },
        //region items ������ ����Ʈ
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
        //region isUpdatedSort ������ ����Ǿ����� ����
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
        //region isUpdated �ʱⰪ �Ǵ� ������ ����Ǿ����� ����
        isUpdated() {
            return this.selectedIdx !== this.tempSelectedIdx || this.isUpdatedSort;
        },
        //endregion
    },
    methods : {
        //region deleteItem ������ ����
        deleteItem(itemIndex) {
            if( confirm('���� �Ͻðڽ��ϱ�?') ) {
                this.callApi(2, 'POST', `/event/contents/tabbar/item/${itemIndex}/delete`, null, () => this.$emit('deleteItem'));
            }
        },
        //endregion
        //region modifyItem ������ ����
        modifyItem(item) {
            this.$emit('modifyItem', item);
        },
        //endregion
        //region setTempItems Set ������ ������ ����Ʈ
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
        //region changeSort ���� ����
        changeSort(e) {
            const moved = e.moved;
            if( moved.oldIndex > moved.newIndex ) // ������ �̵�
                this.moveForward(moved.element, moved.oldIndex+1, moved.newIndex+1);
            else // �ڷ� �̵�
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
        //region saveSortAndSelected �ʱⰪ, �������� ����
        saveSortAndSelected() {
            if( !confirm('���� ���·� �����Ͻðڽ��ϱ�?') )
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
            alert('���� �Ǿ����ϴ�.');
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