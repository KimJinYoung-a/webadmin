Vue.component('Nickname-Write',{
    template: `
        <div style="height: 100px; overflow-y:auto;">
            <form name="play_nickname" id="play_nickname" >
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>직종</th>
                            <td>
                                <select name="occupation" class="form-control inline small" v-model="current_nickname.occupation">
                                    <option selected>직종선택</option>
                                    <option value="Member">Member</option>
                                    <option value="Planner">Planner</option>
                                    <option value="Designer">Designer</option>
                                    <option value="Publisher">Publisher</option>
                                    <option value="Developer">Developer</option>
                                    <option value="MD">MD</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>별명</th>
                            <td>
                                <input v-model="current_nickname.nickname" type="text" name="nickname" />
                            </td>
                        </tr>                        
                    </tbody>
                </table>                
            </form>
        </div>
    `,
    mounted() {
        const _this = this;
    },
    data() {
        return {
            current_nickname: {
                occupation : ""
                , nickname : ""
            }
        }
    },
    props: {
        nickname : {
            occupation : {type:String, default:""} // 컨텐츠 idx
            , nickname : {type:String, default:""}
        }
    },
    methods : {
    }
    , watch:{
        nickname(nickname) { // 컨텐츠 변경 시 현재구분값 set(팝업되었을 때)
            console.log("watch nickname");
            this.current_nickname = nickname;
        }
    }
});