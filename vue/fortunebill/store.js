let store = new Vuex.Store({
    state: {
        list : []
        , content : {}
    }
    , actions: {
        GET_LIST(context){
            callApiHttps("GET", "/automatic-event/fortunebill-list", null, function (data){
                console.log("GET_LIST", data);
                context.commit("SET_LIST", data);
            });
        }
        , GET_CONTENT(context, evt_code){
            callApiHttps("GET", "/automatic-event/fortunebill", {"evt_code" : evt_code}, function (data){
                console.log("GET_CONTENT", data);
                context.commit("SET_CONTENT", data);
            });
        }
    }
    , mutations: {
        SET_LIST(state, data){
            state.list = data;
        }
        , SET_CONTENT(state, data){
            state.content = data;
        }
    }
    , getters: {
        list(state){
            return state.list;
        }
        , content(state){
            return state.content;
        }
    }
});
