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
                            <th>����</th>
                            <td>
                                <select name="occupation" class="form-control inline small" v-model="current_nickname.occupation">
                                    <option selected>��������</option>
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
                            <th>����</th>
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
            occupation : {type:String, default:""} // ������ idx
            , nickname : {type:String, default:""}
        }
    },
    methods : {
    }
    , watch:{
        nickname(nickname) { // ������ ���� �� ���籸�а� set(�˾��Ǿ��� ��)
            console.log("watch nickname");
            this.current_nickname = nickname;
        }
    }
});