let store = new Vuex.Store({
    state: {
        is_write : 0
        , content : {
            link_code : ""
            , applink : ""
            , snstitle : ""
            , snstext : ""
            , option1 : ""
            , option2 : ""
            , option4 : 0
            , option6 : ""
            , option7 : ""
            , option8 : ""
            , listimg : ""
            , open_date : ""
            , end_date : ""
            , pushTitle : ""
            , pushText : ""
        }
        , snstext_length : 0
        , pushtext_length : 0
    }
    , actions: {
        GET_IS_WRITE(context, evt_code){
            callApiHttps("GET", "/event/joobjoob?evt_code=" + evt_code, null, function (data){
                console.log("GET_IS_WRITE", data);
                context.commit("SET_IS_WRITE", data.joobjoobCount);

                if(data.joobjoobCount > 0){
                    data.joobjoob.option6 = data.joobjoob.option6.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",")
                    data.joobjoob.option7 = data.joobjoob.option7.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",")
                    context.commit("SET_CONTENT", data.joobjoob);
                    context.commit("SET_CONTENT_PRIZE", data.joobjoobPrize);
                    context.commit("SET_SNSTEXT_LENTH", data.joobjoob.snstext.length);
                    context.commit("SET_PUSHTEXT_LENTH", data.joobjoob.pushText.length);
                }
            });
        }
    }
    , mutations: {
        SET_IS_WRITE(state, data){
            state.is_write = data;
        }
        , SET_CONTENT(state, data){
            state.content = data;
        }
        , SET_CONTENT_PRIZE(state, data){
            state.content.prize = data;
        }
        , SET_SNSTEXT_LENTH(state, data){
            state.snstext_length = data;
        }
        , SET_PUSHTEXT_LENTH(state, data){
            state.pushtext_length = data;
        }
    }
    , getters: {
        is_write(state){
            return state.is_write;
        }
        , content(state){
            return state.content;
        }
        , snstext_length(state){
            return state.snstext_length;
        }
        , pushtext_length(state){
            return state.pushtext_length;
        }
    }
});
