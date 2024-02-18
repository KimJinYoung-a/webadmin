let store = new Vuex.Store({
    state: {
        is_write : 0
        , content : {
            open_date : ""
            , end_date : ""
            , subcopy : ""
            , mileage_name : ""
            , mileage_expire_date : ""
            , top_img : ""
            , cloud_img : ""
            , coin_img : ""
            , complete_img : ""
        }
    }
    , actions: {
        GET_IS_WRITE(context, evt_code){
            callApiHttps("GET", "/event/everyday-mileage?evt_code=" + evt_code, null, function (data){
                console.log("GET_IS_WRITE", data);
                context.commit("SET_IS_WRITE", data.count);

                if(data.count > 0){
                    context.commit("SET_CONTENT", data.content);
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
    }
    , getters: {
        is_write(state){
            return state.is_write;
        }
        , content(state){
            return state.content;
        }
    }
});
