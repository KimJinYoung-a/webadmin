Vue.component('Content-Write',{
    template: `
        <div style="height: 600px; overflow-y:auto;">
            <p style="color:red;">
                * : 필수값입니다. 빈값일시 저장이 불가능합니다.
            </p>
            <form id="content" enctype="multipart/form-data">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>                        
                        <tr>
                            <th><b style="color: red">*</b> 앱 이벤트 코드</th>
                            <td>
                                <input v-if="write_mode == 'write'" v-model="current_content.evt_code" type="text" name="evt_code" class="must" />
                                <input v-else v-model="current_content.evt_code" type="text" name="evt_code" class="must" readonly style="background-color: grey"/>
                            </td>
                        </tr>
                        <tr>
                            <th><b style="color: red">*</b> 모바일 이벤트 코드</th>
                            <td>
                                <input v-if="write_mode == 'write'" v-model="current_content.mobile_evt_code" type="text" name="mobile_evt_code" class="must" />
                                <input v-else v-model="current_content.mobile_evt_code" type="text" name="mobile_evt_code" class="must" readonly style="background-color: grey"/>
                            </td>
                        </tr>
                        <tr>
                            <th>Text 1</th>
                            <td>
                                <textarea v-model="current_content.text1" name="text1" rows="8" style="width: 80%"></textarea>
                            </td>
                        </tr>                        
                        <tr>
                            <th>Mobile Text 1</th>
                            <td>
                                <textarea v-model="current_content.mobile_text1" name="mobile_text1" rows="8" style="width: 80%"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>참여 대상자 기간 시작일</th>
                            <td>
                                <input v-model="current_content.startdate_of_bill" type="text" name="startdate_of_bill" id="startdate_of_bill" />
                            </td>
                        </tr>   
                        <tr>
                            <th>참여 대상자 기간 종료일</th>
                            <td>
                                <input v-model="current_content.enddate_of_bill" type="text" name="enddate_of_bill" id="enddate_of_bill"/>
                            </td>
                        </tr>        
                        <tr>
                            <th>당첨자 수</th>
                            <td>
                                <input v-model="current_content.number_of_winner" type="text" name="number_of_winner" />
                            </td>
                        </tr>        
                        <tr>
                            <th>담첨자 발표일</th>
                            <td>
                                <input v-model="current_content.winner_notice_date" type="text" name="winner_notice_date" id="winner_notice_date"/>
                            </td>
                        </tr>            
                        <tr>
                            <th>앱 링크</th>
                            <td>
                                <input v-model="current_content.app_link" type="text" name="app_link" />
                            </td>
                        </tr>                        
                    </tbody>
                </table>                
            </form>
        </div>
    `
    , mounted() {
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#startdate_of_bill").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#enddate_of_bill").datepicker('setDate', min_date);
                $("#enddate_of_bill").datepicker('option', "minDate", min_date);

                _this.current_content.startdate_of_bill = document.getElementById("startdate_of_bill").value;
            }
        });

        $("#enddate_of_bill").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.current_content.enddate_of_bill = document.getElementById("enddate_of_bill").value;
            }
        });

        $("#winner_notice_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.current_content.winner_notice_date = document.getElementById("winner_notice_date").value;
            }
        });
    }
    , data() {
        return {
            current_content : {
                evt_code : null
                , mobile_evt_code : null
                , text1 : null
                , mobile_text1 : null
                , startdate_of_bill : null
                , enddate_of_bill : null
                , number_of_winner : 0
                , winner_notice_date : null
                , app_link : null
            }
        }
    }
    , props: {
        content : {
            evt_code : { type:Number, default: null }
            , mobile_evt_code : { type:Number, default: null }
            , text1 : { type:String, default: null }
            , mobile_text1 : { type:String, default: null }
            , startdate_of_bill : { type:String, default: null }
            , enddate_of_bill : { type:String, default: null }
            , number_of_winner : { type:Number, default: 0 }
            , winner_notice_date : { type:String, default: null }
            , app_link : { type:String, default: null }
        }
        , write_mode : {type:String, default:"wait"}
    }
    , methods : {
        init_write_data(){
            this.current_content = {
                evt_code : null
                , mobile_evt_code : null
                , text1 : null
                , mobile_text1 : null
                , startdate_of_bill : null
                , enddate_of_bill : null
                , number_of_winner : 0
                , winner_notice_date : null
                , app_link : null
            }
        }
    }
    , watch:{
        content(content){
            this.init_write_data();
            this.current_content = content;
        }
        , write_mode(write_mode){
            console.log(write_mode);
            const _this = this;
            switch (write_mode){
                case "write" : _this.init_write_data(); break;
            }
        }
    }
});