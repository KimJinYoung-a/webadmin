Vue.component("SCHEDULE", {
    template : `
        <div style="margin-top: 15px;">
            <tr>
                <th>�Ⱓ</th>
                <td>
                    {{content_schedule.startDate.substr(0, 16)}} ~ {{content_schedule.endDate.substr(0, 16)}}
                </td>
                <td colspan="2">
                    <input type="button" value="����" @click="$emit('go_schedule', content_schedule.schedule_idx)"/>
                    <input type="button" value="����" @click="$emit('go_schedule_delete', content_schedule.schedule_idx)"/>
                    <input v-if="raffle_flag == 'Y'" type="button" value="��÷�� ����" @click="$emit('go_winner_popup', content_schedule.schedule_idx)"/>
                </td>
            </tr>
        </div>
    `
    , props : {
        content_schedule : {
            evt_code : {type :String, default : ""}
            , schedule_idx : {type :String, default : ""}
            , startDate : {type :String, default : ""}
            , endDate : {type :String, default : ""}
        }
        , raffle_flag : {type :String, default : "N"}
    }
});