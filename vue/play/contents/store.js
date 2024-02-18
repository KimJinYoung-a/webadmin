let store = new Vuex.Store({
    state: {
        /* 기본 */
        content_count: 0, // 컨텐츠 총 개수
        contents: [], // 컨텐츠 리스트
        current_page: 1, // 현재 페이지
        last_page: 1 // 마지막 페이지
        , opening_list: []
        , pop_content: {}
        , pop_content_items: []
        , pop_content_tag: []
        , show_type: "list"
        , nickname : {}

        /* 검색조건 */
        , period: 1 // 기간 기준
        , startdate: "" // 시작일
        , enddate: "" // 종료일
        , page_size: 10 // 페이지 크기
        , uinumber: 0 // 레이아웃
        , contentsnumber: 0 // 컨텐츠
        , stateflag: 0 // 진행상태
        , searchkey: 1 // 키워드 검색 기준
        , searchstring: "" // 키워드

        /*contents의 검색조건*/
        , isusing : 3
        , isview : 3
        , keywordsearchtype : ""
        , keywordsearch : ""
    }
    , actions: {
        // 컨텐츠 리스트 조회
        GET_CONTENTS(context) {
          const getter = context.getters;
          const api_data = {
            page: getter.current_page
            , page_size: getter.page_size
          };

          const show_type = getter.show_type;
          if(show_type == "list"){
              api_data.period = getter.period;
              if (getter.startdate !== "") api_data.startdate = getter.startdate;
              if (getter.enddate !== "") api_data.enddate = getter.enddate;
              api_data.uinumber = getter.uinumber;
              api_data.contentsnumber = getter.contentsnumber;
              api_data.stateflag = getter.stateflag;
              api_data.searchkey = getter.searchkey;
              if (getter.searchstring !== "") api_data.searchstring = getter.searchstring;

              callApiHttps("GET", "/mobileSite/play/list", api_data, function (data) {
                  context.commit("SET_CONTENT_COUNT", data.count);
                  context.commit("SET_CONTENTS", data.playListDataResponse);
                  context.commit("SET_LAST_PAGE", data.last_page);
              });
          }else if(show_type == "contents"){
              api_data.isusing = getter.isusing;
              api_data.isview = getter.isview;
              api_data.keywordsearchtype = getter.keywordsearchtype;
              api_data.keywordsearch = getter.keywordsearch;

              console.log("api data : ", api_data);
              callApiHttps("GET", "/mobileSite/play/contents", api_data, function (data) {
                  context.commit("SET_CONTENT_COUNT", data.count);
                  context.commit("SET_CONTENTS", data.playContentsData);
                  context.commit("SET_LAST_PAGE", data.last_page);
              });
          }
        }
        , GET_OPENING_LIST(context) {
          if(context.getters.show_type == "list") {
              callApiHttps("GET", "/mobileSite/play/openinglist", null, function (data) {
                  context.commit("SET_OPENING_LIST", data);
              });
          }
        }
        , GET_POP_CONTENT(context, pidx) {
          const show_type = context.getters.show_type;

          if(show_type == "list"){
              callApiHttps("GET", "/mobileSite/play/list-content?pidx=" + pidx, null, function (data) {
                  context.commit("SET_POP_CONTENT", data);
              });

              callApiHttps("GET", "/mobileSite/play/list-content-items?pidx=" + pidx, null, function (data) {
                  context.commit("SET_POP_CONTENT_ITEMS", data);
              });

              callApiHttps("GET", "/mobileSite/play/list-content-tag?pidx=" + pidx, null, function (data) {
                  context.commit("SET_POP_CONTENT_TAG", data);
              });
          }else if(show_type == "contents"){
              callApiHttps("GET", "/mobileSite/play/contents-content?cidx=" + pidx, null, function (data) {
                  context.commit("SET_POP_CONTENT", data);
              });
          }

        }
        , GET_NICKNAME(context){
          callApiHttps("GET", "/mobileSite/play/nickname", null, function (data) {
              context.commit("SET_NICKNAME", data);
          });
        }
    }
    , mutations: {
        SET_CONTENT_COUNT(state, count) {
            // SET 컨텐츠 총 개수
            state.content_count = count;
        },
        SET_CONTENTS(state, contents) {
            // SET 컨텐츠 리스트
            state.contents = [];
            if (contents != null) {
                contents.forEach(content => {
                    content.listimage = decodeBase64(content.listimage);
                    state.contents.push(content);
                });
            }
        },
        SET_CURRENT_PAGE(state, page) {
            state.current_page = page;
        },
        SET_LAST_PAGE(state, page) {
            state.last_page = page;
        },
        SET_PAGE_SIZE(state, size) {
            // SET 페이지별 컨텐츠 수
            state.page_size = size;
        },
        SET_SEARCH_PARAMETER(state, parameter) {
            // SET 검색 파라미터들
            state.period = parameter.period;
            state.startdate = parameter.startdate;
            state.enddate = parameter.enddate;
            state.uinumber = parameter.uinumber;
            state.contentsnumber = parameter.contentsnumber;
            state.stateflag = parameter.stateflag;
            state.searchkey = parameter.searchkey;
            state.searchstring = parameter.searchstring;
        }
        , SET_CONTENTS_SEARCH_PARAMETER(state, parameter){
            console.log("test", parameter);
            state.isusing = parameter.isusing;
            state.isview = parameter.isview;
            state.keywordsearchtype = parameter.keywordsearchtype;
            state.keywordsearch = parameter.keywordsearch;
        }
        , SET_OPENING_LIST(state, openingList) {
            state.opening_list = openingList;
        }
        , SET_POP_CONTENT(state, content) { // SET 기존 컨텐츠
            if (content != null && content.thumbnail != null){
                //content.thumbnail = decodeBase64(content.thumbnail); //TODO : 이미지처리 부분 코드 확인필요.
            }
            state.pop_content = content;
        }
        , SET_POP_CONTENT_ITEMS(state, content){
            state.pop_content_items = content;
        }
        , SET_POP_CONTENT_TAG(state, tag){
            state.pop_content_tag = tag;
        }
        , SET_SHOW_TYPE(state, show_type){
            state.show_type = show_type;
        }
        , SET_NICKNAME(state, data){
            state.nickname = data;
        }
    }
    , getters: {
        content_count(state) {
            return state.content_count;
        },
        contents(state) {
            return state.contents;
        },
        page_size(state) {
            return state.page_size;
        },
        current_page(state) {
            return state.current_page;
        },
        last_page(state) {
            return state.last_page;
        },
        period(state) {
            return state.period;
        },
        startdate(state) {
            return state.startdate;
        },
        enddate(state) {
            return state.enddate;
        },
        uinumber(state) {
            return state.uinumber;
        },
        contentsnumber(state) {
            return state.contentsnumber;
        },
        stateflag(state) {
            return state.stateflag;
        },
        searchkey(state) {
            return state.searchkey;
        },
        searchstring(state) {
            return state.searchstring;
        }
        , opening_list(state){
            return state.opening_list;
        }
        , pop_content(state) {
            return state.pop_content;
        }
        , pop_content_items(state){
            return state.pop_content_items;
        }
        , pop_content_tag(state){
            return state.pop_content_tag;
        }
        , show_type(state){
            return state.show_type;
        }
        , isusing(state){
            return state.isusing;
        }
        , isview(state){
            return state.isview;
        }
        , keywordsearchtype(state){
            return state.keywordsearchtype;
        }
        , keywordsearch(state){
            return state.keywordsearch;
        }
        , nickname(state){
            return state.nickname;
        }
    }
});
