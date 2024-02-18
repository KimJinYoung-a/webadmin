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
                    <!--region 고정 위치-->
                    <tr>
                        <th class="mid"><span class="required">고정 위치<i></i></span></th>
                        <td><input v-model="fix.positionNo" type="text" style="width: 50px;"></td>
                    </tr>
                    <!--endregion-->
                    <!--region 고정 기간-->
                    <tr>
                        <th class="mid"><span class="required">고정 기간<i></i></span></th>
                        <td>
                            <span class="datepicker">
                                <label for="datepicker3">
                                    <strong>시작일</strong>
                                    <span class="mdi mdi-calendar-month"></span>
                                </label>
                                <DATE-PICKER @updateDate="setStartDate" :date="fix.startDate" id="fixStartDate"/>

                                <label for="datepicker4">
                                    <strong>종료일</strong>
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
                <button @click="saveFixPosting" class="linker-btn">저장</button>
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
            positionNo : '', // 고정 순서
            startDate : '', // 시작일자
            endDate : '', // 종료일자
        },
        //endregion
    }},
    props : {
        //region orgFix 기존 고정 포스팅 정보
        orgFix : {
            positionNo : { type:Number, default:0 },
            startDate : { type:String, default:'' },
            endDate : { type:String, default:'' },
        },
        //endregion
    },
    methods : {
        //region syncOrgFixData 기존 고정포스팅 정보 동기화
        syncOrgFixData() {
            this.fix = {
                positionNo : this.orgFix.positionNo,
                startDate : this.orgFix.startDate,
                endDate : this.orgFix.endDate,
            };
        },
        //endregion
        //region setDate Set 시작/종료 일자
        setStartDate(date) {
            this.fix.startDate = date;
        },
        setEndDate(date) {
            this.fix.endDate = date;
        },
        //endregion
        //region saveFixPosting 고정 포스팅 정보 저장
        saveFixPosting() {
            if( !this.validateFixPosting() )
                return false;

            this.$emit('saveFixPosting', this.fix);
        },
        validateFixPosting() {
            if( this.fix.positionNo.trim() === '' ) {
                alert('고정 위치를 입력 해 주세요');
                return false;
            }
            if( isNaN(this.fix.positionNo) ) {
                alert('잘못된 고정 위치값 입니다.');
                return false;
            }
            if( this.fix.startDate.trim() === '' || this.fix.endDate.trim() === '' ) {
                alert('고정 기간을 입력 해 주세요');
                return false;
            }
            if( this.fix.startDate > this.fix.endDate ) {
                alert('잘못된 고정 기간입니다');
                return false;
            }

            return true;
        }
        //endregion
    }
});