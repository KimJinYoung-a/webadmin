let store = new Vuex.Store({
    state: {
        search_count : 0
        , lists : []
        , content : {}
        , content_schedule : []
        , timesale_is_write : false

        , current_page : 1
        , last_page : 1
        , page_size : 10
    }
    , actions: {
        GET_LISTS(context){
            const api_data = {
                current_page: context.getters.current_page
                , page_size: context.getters.page_size
                , raffleFlag : "Y"
            };

            callApiHttps("GET", "/event/timedeal-list", api_data, function (data){
                //console.log("GET_LISTS", data);
                if(data.timedealList){
                    context.commit("SET_SEARCH_COUNT", data.timedealList.length);

                    if(data.timedealList.length > 0){
                        context.commit("SET_LISTS", data.timedealList);
                        context.commit("SET_LAST_PAGE");
                    }
                }
            });
        }
        , GET_CONTENT(context, evt_code){
            if(evt_code){
                callApiHttps("GET", "/event/timedeal-detail", {"evt_code" : evt_code}, function (data){
                    //console.log("GET_CONTENT", data);
                    if(data){
                        context.commit("SET_TIMEDEAL_IS_WRITE", true);
                    }
                    context.commit("SET_CONTENT", data);
                });
            }else{
                context.commit("SET_TIMEDEAL_IS_WRITE", false);
                context.commit("SET_TIMEDEAL_SCHEDULE_NULL");
            }
        }
    }
    , mutations: {
        SET_SEARCH_COUNT(state, data){
            state.search_count = data;
        }
        , SET_LISTS(state, data){
            state.lists = data;
        }
        , SET_CONTENT(state, data){
            state.content = data.timedeal;
            state.content_schedule = data.timedealSchedule;
        }
        , SET_TIMEDEAL_IS_WRITE(state, data){
            state.timesale_is_write = data;
        }
        , SET_CURRENT_PAGE(state, page) {
            state.current_page = page;
        }
        , SET_LAST_PAGE(state, page) {
            state.last_page = page;
        }
        , SET_PAGE_SIZE(state, size) {
            state.page_size = size;
        }
        , SET_TIMEDEAL_SCHEDULE_NULL(state){
            state.content_schedule = [];
        }
    }
    , getters: {
        search_count(state){
            return state.search_count;
        }
        , lists(state){
            return state.lists;
        }
        , content(state){
            return state.content;
        }
        , content_schedule(state){
            return state.content_schedule;
        }
        , timesale_is_write(state){
            return state.timesale_is_write;
        }
        , current_page(state) {
            return state.current_page;
        }
        , last_page(state) {
            return state.last_page;
        }
        , page_size(state) {
            return state.page_size;
        }
    }
});
