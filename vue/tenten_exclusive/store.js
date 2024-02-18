let store = new Vuex.Store({
    state: {
        total_count : 0
        , items : []
        , write : {}
        , main : {}
        , detail : {}
        , is_written : false
    }
    , actions: {
        GET_ITEMS(context, searach_param){
            const api_data = searach_param;

            callApiHttps("GET", "/tenten-exclusive/list", api_data, function (data){
                console.log("GET_ITEMS", data);
                if(data){
                    context.commit("SET_TOTAL_COUNT", data.total_count);

                    if(data.total_count > 0){
                        context.commit("SET_ITEMS", data.items);
                    }
                }
            });
        }
        , GET_WRITE(context, exclusive_idx){
            if(exclusive_idx){
                callApiHttps("GET", "/tenten-exclusive/item", {"exclusive_idx" : exclusive_idx}, function (data){
                    console.log("GET_WRITE", data);
                    if(data){
                        context.commit("SET_IS_WRITTEN", true);
                    }
                    context.commit("SET_WRITE", data);
                });
            }else{
                context.commit("SET_IS_WRITTEN", false);
            }
        }
        , GET_MAIN(context, exclusive_idx){
            callApiHttps("GET", "/tenten-exclusive/item-main", {"exclusive_idx" : exclusive_idx}, function (data){
                console.log("GET_MAIN", data);
                context.commit("SET_MAIN", data);
            });
        }
        , GET_DETAIL(context, exclusive_idx){
            callApiHttps("GET", "/tenten-exclusive/item-detail", {"exclusive_idx" : exclusive_idx}, function (data){
                console.log("GET_DETAIL", data);
                context.commit("SET_DETAIL", data);
            });
        }
    }
    , mutations: {
        SET_TOTAL_COUNT(state, data){
            state.total_count = data;
        }
        , SET_ITEMS(state, data){
            state.items = data;
        }
        , SET_WRITE(state, data){
            state.write = data;
        }
        , SET_IS_WRITTEN(state, data){
            state.is_written = data;
        }
        , SET_MAIN(state, data){
            state.main = data;
        }
        , SET_DETAIL(state, data){
            state.detail = data;
        }
    }
    , getters: {
        total_count(state){
            return state.total_count;
        }
        , items(state){
            return state.items;
        }
        , write(state){
            return state.write;
        }
        , main(state){
            return state.main;
        }
        , detail(state){
            return state.detail;
        }
        , is_written(state){
            return state.is_written;
        }
    }
});
