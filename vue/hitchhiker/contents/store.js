let store = new Vuex.Store({
  state: {
    /* 기본 */
    content_count: 0, // 컨텐츠 총 개수
    contents: [], // 컨텐츠 리스트
    current_page: 1, // 현재 페이지
    last_page: 1, // 마지막 페이지

    /* 검색조건 */
    gubun: "", // 구분
    use_yn: "", // 사용여부
    title: "", // 타이틀
    page_size: 10, // 페이지별 컨텐츠 수
    default_wallpaper_sizes: {
      // default 배경화면 사이즈 리스트
      mobile_wallpaper_sizes: [], // Mobile
      pc_wallpaper_sizes: [], // PC
    },
    pop_content: {}, // 팝업 컨텐츠
  },

  getters: {
    content_count(state) {
      return state.content_count;
    },
    contents(state) {
      return state.contents;
    },
    current_page(state) {
      return state.current_page;
    },
    last_page(state) {
      return state.last_page;
    },
    gubun(state) {
      return state.gubun;
    },
    use_yn(state) {
      return state.use_yn;
    },
    page_size(state) {
      return state.page_size;
    },
    title(state) {
      return state.title;
    },
    default_wallpaper_sizes(state) {
      return state.default_wallpaper_sizes;
    },
    pop_content(state) {
      return state.pop_content;
    },
  },

  mutations: {
    SET_CONTENT_COUNT(state, count) {
      // SET 컨텐츠 총 개수
      state.content_count = count;
    },
    SET_CONTENTS(state, contents) {
      // SET 컨텐츠 리스트
      state.contents = [];
      if (contents != null) {
        contents.forEach(content => {
          content.thumbnail = decodeBase64(content.thumbnail);
          state.contents.push(content);
        });
      }
    },
    SET_CURRENT_PAGE(state, page) {
      // SET 현재 페이지
      state.current_page = page;
    },
    SET_LAST_PAGE(state, page) {
      // SET 마지막 페이지
      state.last_page = page;
    },
    SET_PAGE_SIZE(state, size) {
      // SET 페이지별 컨텐츠 수
      state.page_size = size;
    },
    SET_SEARCH_PARAMETER(state, parameter) {
      // SET 검색 파라미터들
      state.gubun = parameter.gubun;
      state.use_yn = parameter.use_yn;
      state.title = parameter.title;
    },
    SET_DEFAULT_WALLPAPER_SIZES(state, size) {
      // SET default 배경화면 사이즈 리스트
      state.default_wallpaper_sizes = size;
    },
    SET_POP_CONTENT(state, content) {
      // SET 기존 컨텐츠 조회
      if (content != null && content.thumbnail != null) content.thumbnail = decodeBase64(content.thumbnail);
      state.pop_content = content;
    },
  },

  actions: {
    // 컨텐츠 리스트 조회
    GET_CONTENTS(context) {
      const getter = context.getters;
      const api_data = {
        page: getter.current_page,
        page_size: getter.page_size,
      };
      if (getter.gubun !== "") api_data.gubun = getter.gubun;
      if (getter.use_yn !== "") api_data.use_yn = getter.use_yn;
      if (getter.title.trim() !== "") api_data.title = getter.title.trim();

      callApi("GET", "/hitchhiker/content", api_data, function (data) {
        context.commit("SET_CONTENT_COUNT", data.total_count);
        context.commit("SET_CONTENTS", data.contents);
        context.commit("SET_LAST_PAGE", data.last_page);
      });
    },
    // GET default 배경화면 사이즈 리스트
    GET_DEFAULT_WALLPAPER_SIZES(context) {
      callApi("GET", "/hitchhiker/wallpaper/size", null, function (data) {
        context.commit("SET_DEFAULT_WALLPAPER_SIZES", data);
      });
    },
    // GET 기존 컨텐츠 조회
    GET_CONTENT_ONE(context, content_idx) {
      callApi("GET", "/hitchhiker/content/" + content_idx, null, function (data) {
        context.commit("SET_POP_CONTENT", data);
      });
    },
    // 사이즈관리 저장
    SAVE_WALLPAPER_SIZE(context, size_data) {
      console.log(size_data);

      const app = size_data.app;
      size_data.app = null;

      const set_contents = function (data) {
        console.log("SAVE_WALLPAPER_SIZE\n", data);
        context.commit("SET_DEFAULT_WALLPAPER_SIZES", data);
      };

      if (size_data.device_idx != null && size_data.device_idx !== 0) {
        callApi("PUT", "/hitchhiker/wallpaper/size", size_data, set_contents);
      } else {
        callApi("POST", "/hitchhiker/wallpaper/size", size_data, set_contents);
      }
    },
  },
});
