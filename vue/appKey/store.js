let store = new Vuex.Store({
    state: {
        key_list : []
        , content : {
            validationKey : ""
        }
    }
    , actions: {
        GET_KEY_LIST(context){
            // 기존 callApiHttps가 /v1/ 경로를 포함중이라 v2를 사용할수 없는 문제가 있음.
            callApiHttpsV2("GET", "/v2/app/appkey-list", null, function (data){
                console.log("GET_KEY_LIST", data);
                context.commit("SET_KEY_LIST", data);
            });
        }
        , GET_KEY(context, idx){
            if(idx){
                // 기존 callApiHttps가 /v1/ 경로를 포함중이라 v2를 사용할수 없는 문제가 있음.
                callApiHttpsV2("GET", "/v2/app/appkey", {"idx" : idx}, function (data){
                    console.log("GET_KEY", data);
                    context.commit("SET_KEY", data);
                });
            }else{
                context.commit("SET_KEY", {"validationKey" : "", "isusing" : "", "description" : ""});
            }
        }
    }
    , mutations: {
        SET_KEY_LIST(state, data){
            state.key_list = data;
        }
        , SET_KEY(state, data){
            state.content = data;
        }
    }
    , getters: {
        key_list(state){
            return state.key_list;
        }
        , content(state){
            return state.content;
        }
    }
});
