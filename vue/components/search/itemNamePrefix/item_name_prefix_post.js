Vue.component('ITEM-NAME-PREFIX-POST', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:100px;">
                    <col>
                </colgroup>
                <tbody>
                    <tr>
                        <th>���Ӹ�</th>
                        <td><input v-model="prefixWord" type="text" placeholder="���Ӹ��� �Է����ּ���"></td>
                    </tr>
                    <tr>
                        <th>�Ⱓ</th>
                        <td>
                            <span class="datepicker">
                                <label for="startDate">
                                    <strong>������</strong>
                                </label>
                                <DATE-PICKER @updateDate="updateStartDate" :date="startDate" id="startDate"/>
                                <input v-model="startTime" type="time">
                            </span>
                            <span class="datepicker">
                                <label for="endDate">
                                    <strong>������</strong>
                                </label>
                                <DATE-PICKER @updateDate="updateEndDate" :date="endDate" id="endDate"/>
                                <input v-model="endTime" type="time">
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <th>��뿩��</th>
                        <td>
                            <p class="radio-area">
                                <input v-model="use" :value="true" id="itemNamePrefixUseY" type="radio">
                                <label for="itemNamePrefixUseY">Y</label>
                                <input v-model="use" :value="false" id="itemNamePrefixUseN" type="radio">
                                <label for="itemNamePrefixUseN">N</label>
                            </p>
                        </td>
                    </tr>
                </tbody>
            </table>
            
            <!--region ���/���� ���� ����-->
            <div v-if="failProducts.length > 0" class="modal-alert">
                <a @click="failProducts = []" class="close">x</a>
                <strong>���� ����</strong>
                <p v-for="product in failProducts">
                    - ��ǰ�ڵ� : {{product.duplicatedProductId}}, 
                    ��ǰ�� : "{{product.duplicatedProductName}}" 
                    > ���Ӹ� [{{product.duplicatedPrefixWord}}]�� �̺�Ʈ�Ⱓ ��ħ
                </p>
            </div>
            <!--endregion-->
            
            <div class="modal-btn-area">
                <button @click="postPrefix" class="btn">����</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.updatePrefix ) {
            this.setUpdatePrefixData();
        }
    },
    data() {return {
        prefixWord : '', // ���Ӹ�
        startDate : '', // ���� ����
        startTime : '', // ���� �ð�
        endDate : '', // ���� ����
        endTime : '', // ���� �ð�
        use : false, // ��� ����
        failProducts : [], // ������ ��ǰ ����Ʈ
    }},
    props : {
        //region updatePrefix ������ ���Ӹ�
        updatePrefix : {
            prefixIdx : { type:Number, default:0 },
            prefixWord : { type:String, default:'' },
            startDate : { type:String, default:'' },
            endDate : { type:String, default:'' },
            use : { type:String, default:'' },
            state : { type:String, default:'' },
            itemCount : { type:Number, default:0 },
            regDate : { type:String, default:'' },
            regAdminId : { type:String, default:'' },
            regAdminName : { type:String, default:'' }
        },
        //endregion
    },
    computed : {
        //region postPrefixApiSendData ���Ӹ� ���� API ���� ������
        postPrefixApiSendData() {
            const data = {
                prefixWord : this.prefixWord,
                startDate : this.startDate + ' ' + this.startTime,
                endDate : this.endDate + ' ' + this.endTime,
                use : this.use
            };

            if( this.updatePrefix )
                data.prefixIdx = this.updatePrefix.prefixIdx;

            return data;
        },
        //endregion
    },
    methods : {
        //region postPrefix ���Ӹ� ����
        postPrefix() {
            if( !confirm('���� �Ͻðڽ��ϱ�?') )
                return false;

            this.callApi(1, 'POST', '/search/prefix', this.postPrefixApiSendData, this.successPostPrefix);
        },
        successPostPrefix(failProducts) {
            if( failProducts.length > 0 ) {
                this.failProducts = failProducts;
            } else {
                this.$emit('postPrefix');
            }
        },
        //endregion
        //region setUpdatePrefixData ���� ���Ӹ� ������ ����
        setUpdatePrefixData() {
            this.prefixWord = this.updatePrefix.prefixWord;
            this.use = this.updatePrefix.use;
            this.setUpdatePrefixDateAndTime('start', this.updatePrefix.startDate);
            this.setUpdatePrefixDateAndTime('end', this.updatePrefix.endDate);
        },
        setUpdatePrefixDateAndTime(flag, date) {
            const dateArr = date.split('T');
            this[`${flag}Date`] = dateArr[0];
            this[`${flag}Time`] = dateArr[1];
        },
        //endregion
        //region updateStartDate ������ ����
        updateStartDate(date) {
            this.startDate = date;
        },
        //endregion
        //region updateEndDate ������ ����
        updateEndDate(date) {
            this.endDate = date;
        },
        //endregion
    }
});