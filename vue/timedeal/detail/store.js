let store = new Vuex.Store({
    state: {
        evt_code : ""
        , schedule_idx : ""
        , schedule : {}
        , content_mikki : {}
        , mikki_is_write : false
        , mikki_list : []
        , normal_list : []
    }
    , actions: {
        GET_SCHEDULE(context, schedule_idx){
            callApiHttps("GET", "/event/timedeal-schedule", {"evt_code" : context.getters.evt_code, "schedule_idx" : schedule_idx}, function (data){
                //console.log("GET_SCHEDULE", data);
                context.commit("SET_SCHEDULE", data);
            });
        }
        , GET_MIKKI_DETAIL(context, param){
            if(param){
                callApiHttps("GET", "/event/timedeal-mikki-detail"
                    , {"evt_code" : context.getters.evt_code, "schedule_idx" : context.getters.schedule_idx, "startDate" : param[0], "endDate" : param[1]}
                    , function (data){
                        //console.log("GET_MIKKI_DETAIL", data);
                        if(data){
                            context.commit("SET_MIKKI_IS_WRITE", true);
                        }else{
                            context.commit("SET_MIKKI_IS_WRITE", false);
                        }

                        data.orgPrice = data.orgPrice.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
                        data.sellCash = data.sellCash.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
                        context.commit("SET_MIKKI_DETAIL", data);
                    }
                );
            }else{
                context.commit("SET_MIKKI_IS_WRITE", false);
            }
        }
        , DELETE_MIKKI_DETAIL(context, param){
            callApiHttps("DELETE", "/event/timedeal-mikki-detail"
                , {"evt_code" : context.getters.evt_code, "schedule_idx" : context.getters.schedule_idx, "startDate" : param[0], "endDate" : param[1]}
            );
        }
        , GET_MIKKI_LIST(context, schedule_idx){
            callApiHttps("GET", "/event/timedeal-mikki-list", {"evt_code" : context.getters.evt_code, "schedule_idx" : schedule_idx}, function (data){
                //console.log("GET_MIKKI_LIST", data);
                context.commit("SET_MIKKI_LIST", data);
            });
        }
        , GET_NORMAL_LIST(context, schedule_idx){
            callApiHttps("GET", "/event/timedeal-normal-list", {"evt_code" : context.getters.evt_code, "schedule_idx" : schedule_idx}, function (data){
                //console.log("GET_NORMAL_LIST", data);
                context.commit("SET_NORMAL_LIST", data);
            });
        }
    }
    , mutations: {
        SET_SCHEDULE(state, data){
            state.schedule = data;
        }
        , SET_EVT_CODE(state, data){
            state.evt_code = data;
        }
        , SET_SCHEDULE_IDX(state, data){
            state.schedule_idx = data;
        }
        , SET_MIKKI_IS_WRITE(state, data){
            state.mikki_is_write = data;
        }
        , SET_MIKKI_LIST(state, data){
            state.mikki_list = data;
        }
        , SET_MIKKI_DETAIL(state, data){
            state.content_mikki = data;
        }
        , SET_NORMAL_LIST(state, data){
            state.normal_list = data;
        }
        , SET_NORMAL_LIST_EMPTY(state){
            state.normal_list = [];
        }
    }
    , getters: {
        schedule(state){
            return state.schedule;
        }
        , evt_code(state){
            return state.evt_code;
        }
        , schedule_idx(state){
            return state.schedule_idx;
        }
        , content_mikki(state){
            return state.content_mikki;
        }
        , mikki_list(state){
            return state.mikki_list;
        }
        , mikki_is_write(state){
            return state.mikki_is_write;
        }
        , normal_list(state){
            return state.normal_list;
        }
    }
});
