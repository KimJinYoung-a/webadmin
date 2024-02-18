Vue.component('Snack-Sort',{
    template: `
        <div style="height: 600px; overflow-y:auto;">
            <p style="color:red;">
                상태값이 진행중인것들만 노출됩니다.
            </p>
            <form id="snack_content" enctype="multipart/form-data">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody id="sorting_row">                        
                        <tr v-for="item in content" :data-idx="item.video_idx">
                            <th>정렬순서</th>
                            <td>
                                {{item.sort_no}}
                            </td>
                            
                            <th>idx</th>
                            <td>{{item.video_idx}}</td>
                            
                            <th>항목설명</th>
                            <td>{{item.entry_desc}}</td>
                        </tr>
                    </tbody>
                </table>                
            </form>
        </div>
    `
    , mounted() {
        const _this = this;

        $("#sorting_row").sortable({
            delay: 150
            , stop: function() {
                let sortedArrVar = new Array();

                $('#sorting_row > tr').each(function() {
                    sortedArrVar.push($(this).attr("data-idx"));
                });
                _this.sorted_arr = sortedArrVar;
            }
        });
    }
    , data() {
        return {
            sorted_arr : []
        }
    }
    , props: {
        content : {
            video_idx : { type:String, default: null }
            , entry_desc : { type:String, default: null }
        }
    }
    , methods : {
    }
    , watch:{
    }
});