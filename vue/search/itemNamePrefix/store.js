const store = new Vuex.Store({
    state : {
        app : null, // Root

        prefixCurrentPage : 1, // 말머리 현재 페이지
        //region prefixSearch 말머리 검색 파라미터
        prefixSearch : {
            startDate : '', // 시작일자
            endDate : '', // 종료일자
            keyword : '' // 키워드
        },
        //endregion

        prefixCount : 0, // 말머리 갯수
        prefixLastPage : 1, // 말머리 마지막 페이지
        prefixes : [], // 말머리 리스트
    },
    getters : {
        app(state) {return state.app;},
        prefixCount(state) {return state.prefixCount;},
        prefixCurrentPage(state) {return state.prefixCurrentPage;},
        prefixSearch(state) {return state.prefixSearch;},
        prefixLastPage(state) {return state.prefixLastPage;},
        prefixes(state) {return state.prefixes;},
    },
    mutations : {
        SET_APP(state, app) { state.app = app; },
        SET_PREFIX_CURRENT_PAGE(state, page) { state.prefixCurrentPage = page; },
        SET_PREFIX_SEARCH(state, payload) { state.prefixSearch = payload; },
        SET_PREFIX_COUNT(state, count) { state.prefixCount = count; },
        SET_PREFIX_LAST_PAGE(state, page) { state.prefixLastPage = page; },
        SET_PREFIXES(state, prefixes) { state.prefixes = prefixes; },
        //region UPDATE_PREFIX_ITEM_COUNT 말머리 상품 수 수정
        UPDATE_PREFIX_ITEM_COUNT(state, payload) {
            const prefix = state.prefixes.find(p => p.prefixIdx === payload.prefixIdx);
            prefix.itemCount = payload.itemCount;
        },
        //endregion
    },
    actions : {
        //region GET_PREFIXES 말머리 리스트 조회
        GET_PREFIXES(context) {
            const app = context.getters.app;
            const data = {
                startDate : context.getters.prefixSearch.startDate,
                endDate : context.getters.prefixSearch.endDate,
                keyword : context.getters.prefixSearch.keyword,
                page : context.getters.prefixCurrentPage
            };
            app.callApi(1, 'GET', '/search/prefixes', data, data => {
                context.commit('SET_PREFIX_COUNT', data.totalCount);
                context.commit('SET_PREFIX_LAST_PAGE', data.lastPage);
                context.commit('SET_PREFIXES', data.prefixes);
            });
        },
        //endregion
    }
});