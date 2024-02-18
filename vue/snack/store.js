let store = new Vuex.Store({
    state: {
        list : []
        , total_count : 0
        , current_page : 1
        , last_page : 1
        , content : {}
        , sort_list : []
    }
    , actions: {
        GET_LIST(context, search_param){
            let api_url;
            if (location.hostname.startsWith('webadmin')) {
                api_url = '//fapi.10x10.co.kr/api/web/v1';
            }else if(location.hostname.startsWith('localwebadmin')) {
                api_url = '//localhost:8080/api/web/v1';
            }else{
                api_url = '//testfapi.10x10.co.kr/api/web/v1';
            }

            let api_data = {
                "current_page" : context.getters.current_page
                , "search_state" : search_param.state
                , "search_entry_type" : search_param.search_entry_type
                , "search_keyword_option" : search_param.search_keyword_option
                , "search_keyword_text" : search_param.search_keyword_text
                , "search_encoding_state" : search_param.search_encoding_state
            };

            $.ajax({
                type: "GET"
                , url: api_url + "/snack/list"
                , data : api_data
                , crossDomain: true
                , xhrFields: {
                    withCredentials: true
                }
                , success: function(data){
                    context.commit("SET_LIST", data);
                }
            });
        }
        , GET_CONTENT(context, video_idx) {
            const _this = this;

            let api_url;
            if (location.hostname.startsWith('webadmin')) {
                api_url = '//fapi.10x10.co.kr/api/web/v1';
            }else if(location.hostname.startsWith('localwebadmin')) {
                api_url = '//localhost:8080/api/web/v1';
            }else{
                api_url = '//testfapi.10x10.co.kr/api/web/v1';
            }
            $.ajax({
                type: "GET"
                , url: api_url + "/snack/content"
                , data : {"video_idx" : video_idx}
                , crossDomain: true
                , xhrFields: {
                    withCredentials: true
                }
                , success: function(data){
                    context.commit("SET_CONTENT", data);
                }
            });
        }
        , GET_SORT_LIST(context){
            let api_url;
            if (location.hostname.startsWith('webadmin')) {
                api_url = '//fapi.10x10.co.kr/api/web/v1';
            }else if(location.hostname.startsWith('localwebadmin')) {
                api_url = '//localhost:8080/api/web/v1';
            }else{
                api_url = '//testfapi.10x10.co.kr/api/web/v1';
            }

            $.ajax({
                type: "GET"
                , url: api_url + "/snack/sort-list"
                , data : {}
                , crossDomain: true
                , xhrFields: {
                    withCredentials: true
                }
                , success: function(data){
                    context.commit("SET_SORT_LIST", data);
                }
            });
        }
    }
    , mutations: {
        SET_LIST(state, data){
            state.list = data.items;
            state.total_count = data.total_count;
            state.current_page = data.current_page;
            state.last_page = data.last_page;
        }
        , SET_CONTENT(state, data){
            state.content = data;
        }
        , SET_CURRENT_PAGE(state, data){
            state.current_page = data;
        }
        , SET_SORT_LIST(state, data){
            state.sort_list = data;
        }
    }
    , getters: {
        list(state){
            return state.list;
        }
        , content(state){
            return state.content;
        }
        , total_count(state){
            return state.total_count;
        }
        , current_page(state){
            return state.current_page;
        }
        , last_page(state){
            return state.last_page;
        }
        , sort_list(state){
            return state.sort_list;
        }
    }
});
