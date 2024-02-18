Vue.component("SCHEDULE", {
    template : `
        <div style="margin-top: 15px;">
            <tr>
                <th>기간</th>
                <td>
                    {{content_schedule.startDate.substr(0, 16)}} ~ {{content_schedule.endDate.substr(0, 16)}}
                </td>
                <td colspan="2">
                    <input type="button" value="수정" @click="$emit('go_schedule', content_schedule.schedule_idx)"/>
                    <input type="button" value="삭제" @click="$emit('go_schedule_delete', content_schedule.schedule_idx)"/>
                    <input v-if="raffle_flag == 'Y'" type="button" value="당첨자 보기" @click="$emit('go_winner_popup', content_schedule.schedule_idx)"/>
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