Vue.component('MANAGE-FIX-POSTING', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <!--region colgroup-->
                <colgroup>
                    <col style="width:120px;">
                    <col>
                </colgroup>
                <!--endregion-->
                <tbody>
                    <!--region ���� ��ġ-->
                    <tr>
                        <th class="mid"><span class="required">���� ��ġ<i></i></span></th>
                        <td><input v-model="fix.positionNo" type="text" style="width: 50px;"></td>
                    </tr>
                    <!--endregion-->
                    <!--region ���� �Ⱓ-->
                    <tr>
                        <th class="mid"><span class="required">���� �Ⱓ<i></i></span></th>
                        <td>
                            <span class="datepicker">
                                <label for="datepicker3">
                                    <strong>������</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setStartDate" :date="fix.startDate" id="fixStartDate"/>

                                <label for="datepicker4">
                                    <strong>������</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setEndDate" :date="fix.endDate" id="fixEndDate"/>
                            </span>
                        </td>
                    </tr>
                    <!--endregion-->
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveFixPosting" class="linker-btn">����</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.orgFix )
            this.syncOrgFixData();
    },
    data() {return {
        //region fix
        fix : {
            positionNo : '', // ���� ����
            startDate : '', // ��������
            endDate : '', // ��������
        },
        //endregion
    }},
    props : {
        //region orgFix ���� ���� ������ ����
        orgFix : {
            positionNo : { type:Number, default:0 },
            startDate : { type:String, default:'' },
            endDate : { type:String, default:'' },
        },
        //endregion
    },
    methods : {
        //region syncOrgFixData ���� ���������� ���� ����ȭ
        syncOrgFixData() {
            this.fix = {
                positionNo : this.orgFix.positionNo,
                startDate : this.orgFix.startDate,
                endDate : this.orgFix.endDate,
            };
        },
        //endregion
        //region setDate Set ����/���� ����
        setStartDate(date) {
            this.fix.startDate = date;
        },
        setEndDate(date) {
            this.fix.endDate = date;
        },
        //endregion
        //region saveFixPosting ���� ������ ���� ����
        saveFixPosting() {
            if( !this.validateFixPosting() )
                return false;

            this.$emit('saveFixPosting', this.fix);
        },
        validateFixPosting() {
            if( this.fix.positionNo.trim() === '' ) {
                alert('���� ��ġ�� �Է� �� �ּ���');
                return false;
            }
            if( isNaN(this.fix.positionNo) ) {
                alert('�߸��� ���� ��ġ�� �Դϴ�.');
                return false;
            }
            if( this.fix.startDate.trim() === '' || this.fix.endDate.trim() === '' ) {
                alert('���� �Ⱓ�� �Է� �� �ּ���');
                return false;
            }
            if( this.fix.startDate > this.fix.endDate ) {
                alert('�߸��� ���� �Ⱓ�Դϴ�');
                return false;
            }

            return true;
        }
        //endregion
    }
});