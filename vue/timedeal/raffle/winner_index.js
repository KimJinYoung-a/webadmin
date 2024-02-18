var app = new Vue({
    el: "#app"
    , store: store
    , template: `
        <div style="height: 300px;">
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:30%;">
                    <col style="width:30%;">
                    <col style="width:40%;">
                </colgroup>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>상품코드</th>
                        <th>연락처</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in winners" class="link">
                        <td>{{item.userid}}</td>
                        <td>{{item.itemid}}</td>
                        <td>{{item.phone}}</td>
                    </tr>
                </tbody>
            </table>
            
            <table class="table table-dark">
                <colgroup>
                    <col style="width:50%;">
                    <col style="width:50%;">
                </colgroup>
                <thead>
                    <tr>
                        <th>상품코드</th>
                        <th>응모자수</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in subscript_count" class="link">
                        <td>{{item.itemid}}</td>
                        <td>{{item.count}}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `
    , data() {
        return{
        }
    }
    , created() {
        let query_param = new URLSearchParams(window.location.search);
        this.$store.commit("SET_EVT_CODE", query_param.get("evt_code"));
        this.$store.commit("SET_SCHEDULE_IDX", query_param.get("schedule_idx"));

        this.$store.dispatch("GET_WINNER_INFO");
    }
    , mounted(){
        const _this = this;
    }
    , computed: {
        winners(){
            return this.$store.getters.winners;
        }
        , subscript_count(){
            return this.$store.getters.subscript_count;
        }
    }
    , methods: {

    }
});
