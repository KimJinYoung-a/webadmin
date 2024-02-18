let store = new Vuex.Store({
    state: {
        winners : []
        , subscript_count : []
    }
    , actions: {
        GET_WINNER_INFO(context){
            const api_data = {
                evt_code: context.getters.evt_code
                , schedule_idx: context.getters.schedule_idx
            };

            callApiHttps("GET", "/event/timedeal-raffle-winner", api_data, function (data){
                //console.log("GET_LISTS", data);
                if(data){
                    context.commit("SET_WINNERS", data.winners);
                    context.commit("SET_SUBSCRIPT_COUNT", data.subscript_count);
                }
            });
        }
    }
    , mutations: {
        SET_EVT_CODE(state, data){
            state.evt_code = data;
        }
        , SET_SCHEDULE_IDX(state, data){
            state.schedule_idx = data;
        }
        , SET_WINNERS(state, data){
            state.winners = data;
        }
        , SET_SUBSCRIPT_COUNT(state, data){
            state.subscript_count = data;
        }
    }
    , getters: {
        evt_code(state){
            return state.evt_code;
        }
        , schedule_idx(state){
            return state.schedule_idx;
        }
        , winners(state){
            return state.winners;
        }
        , subscript_count(state){
            return state.subscript_count;
        }
    }
});
