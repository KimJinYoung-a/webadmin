const app = new Vue({
    el: "#app"
    , store: store
    , mixins: [api_mixin]
    , template: `
        <div style="height: 400px; overflow-y:auto;">
            <form id="today_content" enctype="multipart/form-data">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>                        
                        <tr>
                            <th>구분명</th>
                            <td>
                                <select v-model="current_content.poscode" class="form-control inline" name="poscode">
                                    <option v-if="page_type == 'main'" value="2075">2017_메인롤링_MA</option>
                                    <option v-if="page_type == 'main'" value="2079">2017_이미지배너A/B</option>
                                    <option v-if="page_type == 'round'" value="10000">라운드배너</option>
                                    <option v-if="page_type == 'enjoy'" value="10001">메인상단기획전(모바일)</option>
                                    <option v-if="page_type == 'event'" value="10002">메인상단기획전(PC)</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>날짜</th>
                            <td>
                                <input type="text" name="start_date" id="start_date" class="form-control inline small" /> ~ <input type="text" name="end_date" id="end_date" class="form-control inline small" />
                            </td>
                        </tr>
                        <tr>
                            <th>이벤트</th>
                            <td>
                                <p style="color: red">-[ 숫자,숫자 ] 형식의 데이터만 저장됩니다.</p>
                                <textarea v-model="current_content.multicode" name="multicode" style="width: 80%" rows="10" ></textarea>
                            </td>
                        </tr>                        
                    </tbody>
                </table>                
            </form>

            <div style="text-align: right; margin: 11px 11px;">
                <button @click="save" class="button dark">저장</button>
                <button @click="close" class="button secondary">취소</button>
            </div>
        </div>
    `
    , data() {
        return {
            current_content : {
                poscode : ""
                , multicode : ""
                , start_date : ""
                , end_date : ""
            }
            , tmp_multicode_list : []
            , page_type : ""
        };
    }
    , created() {
        const _this = this;
        let query_param = new URLSearchParams(window.location.search);
        this.page_type = query_param.get("type");
    }
    , computed: {

    }
    , mounted() {
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#today_content #start_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                if(date > $("#end_date").val()){
                    _this.current_content.end_date = date;
                }

                $("#end_date").datepicker('option', "minDate", date);
                _this.current_content.start_date = date;
            }
        });
        $("#today_content #end_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.current_content.end_date = date;
            }
        });

        let now = new Date();
        let tomorrow = new Date(now.setDate(now.getDate()+1));

        $("#start_date").datepicker('setDate', tomorrow);
        this.current_content.start_date = $("#start_date").val();
        $("#end_date").datepicker('setDate', tomorrow);
        this.current_content.end_date = $("#end_date").val();
    }
    , methods: {
        save(){
            const _this = this;
            if(this.validate()){
                let api_data = {
                    "poscode" : this.current_content.poscode
                    , "multicode" : this.tmp_multicode_list
                    , "start_date" : this.current_content.start_date
                    , "end_date" : this.current_content.end_date
                };

                callApiHttps("POST", "/mobile/main/today-easy", api_data, function(data){
                    console.log(data);

                    if(data.result){
                        alert("저장이 완료되었습니다.");
                        _this.close();
                    }else{
                        alert(data.errmsg);
                    }
                });
            }
        }
        , close(){
            opener.location.reload();
            self.close();
        }
        , validate(){
            const multicode_regex = /(^[0-9]+,[0-9]+\n?$)+/;
            let multicode_list = this.current_content.multicode.split("\n");
            let result = true;
            let local_tmp = new Array();

            for(let i = 0; i < multicode_list.length; i++){
                if(!multicode_regex.test(multicode_list[i])){
                    result = false;
                    alert("이벤트항목 데이터 형식이 유효하지 못합니다.");

                    break;
                }
                local_tmp.push(multicode_list[i]);
            }

            this.tmp_multicode_list = local_tmp;
            return result;
        }
    }
    , watch : {
        page_type(page_type){
            const _this = this;

            switch (page_type){
                case "main" : _this.current_content.poscode = "2075"; break;
                case "round" : _this.current_content.poscode = "10000"; break;
                case "enjoy" : _this.current_content.poscode = "10001"; break;
                case "event" : _this.current_content.poscode = "10002"; break;
            }
        }
    }
});
