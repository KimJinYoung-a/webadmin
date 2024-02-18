var app = new Vue({
  el: "#app",
  store: store,
  template: `
        <div>
            <!-- 검색 테이블 -->
            <table class="table table-dark table-search">
                <colgroup>
                    <col style="width:150px;">
                    <col style="width:300px;">
                    <col style="width:150px;">
                    <col>
                    <col style="width:200px;">
                </colgroup>
                <thead class="thead-tenbyten">
                    <tr>
                        <th>검색조건</th>
                        <td style="text-align:left;">
                            <select id="gubun" class="form-control inline small">
                                <option value="">구분</option>
                                <option value="1">PC</option>
                                <option value="2">MOBILE</option>
                                <option value="3">MOVIE</option>
                                <option value="4">MOBILE배경</option>
                            </select>
                            <select id="use_yn" class="form-control inline small">
                                <option value="">사용여부</option>
                                <option>Y</option>
                                <option>N</option>
                            </select>
                        </td>
                        <th>타이틀</th>
                        <td style="text-align:left;">
                            <input id="title" type="text" class="form-control">
                        </td>
                        <td class="td-button align-right">
                            <button @click="reload" type="button" class="button secondary">검색조건Reset</button>
                            <button @click="do_search" type="button" class="button dark">검색</button>
                        </td>
                    </tr>
                </thead>
            </table>
            <p class="p-table">
                <span>검색결과 : <strong>{{content_count}}</strong></span>
                <i class='fas fa-sync' @click="reload"></i>
                <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                    <option v-for="n in 5" :value="n*10">{{n*10}}개씩 보기</option>
                </select>
                <button id="modify_wallpaper_size" @click="popup_manage_wallpaper_size" type="button" class="button dark">사이즈 관리</button>
                <button id="reg_new_content" @click="popup_manage_content" type="button" class="button dark">신규 등록</button>
            </p>

            <!-- 리스트 테이블 -->
            <table class="table table-dark">
                <colgroup>
                    <col style="width:10%;">
                    <col style="width:10%;">
                    <col style="width:15%;">
                    <col>
                    <col style="width:70px;">
                    <col style="width:15%;">
                    <col style="width:15%;">
                </colgroup>
                <thead>
                    <tr>
                        <th>번호</th>
                        <th>구분</th>
                        <th>썸네일</th>
                        <th>타이틀</th>
                        <th>사용여부</th>
                        <th>시작일</th>
                        <th>등록일</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="content in contents" :key="content.content_idx">
                        <td @click="popup_manage_content" class="link">{{content.content_idx}}</td>
                        <td>{{content.gubun}}</td>
                        <td>
                            <img @click="popup_thumbnail" :src="content.thumbnail" class="img-thumbnail link" style="width:50px;height:50px;">
                        </td>
                        <td>{{content.title}}</td>
                        <td>{{content.use_yn ? 'Y' : 'N'}}</td>
                        <td v-html="get_start_date(content.start_date, content.open_yn)"></td>
                        <td>{{content.reg_date}}</td>
                    </tr>
                </tbody>
            </table>

            <!-- 페이지 -->
            <Pagination
                @click_page="click_page"
                :current_page="current_page"
                :last_page="last_page"
            ></Pagination>

            <!-- 등록/수정 모달 -->
            <Modal v-show="show_write_modal"
                @save="save_content" @close="show_write_modal = false"
                modal_width="830px" header_title="히치하이커 컨텐츠 등록"
            >
                <Hitchhiker-Content-Write slot="body" :content="pop_content" :wallpaper_sizes="default_wallpaper_sizes"/>
            </Modal>

            <!-- 썸네일 모달 -->
            <Modal v-show="show_thumbnail_modal" @close="show_thumbnail_modal = false"
                modal_width="400px" :show_header_yn="false" :show_footer_yn="false"
                :close_background_click_yn="true"
            >
                <img width="100%" :src="popup_thumbnail_src" slot="body">
            </Modal>

            <!-- 사이즈관리 모달 -->
            <Modal v-show="show_size_modal" modal_width="870px" header_title="히치하이커 배경화면 사이즈 관리">
                <Hitchhiker-Wallpaper-Size-Write
                    slot="body" ref="wallpaper_size_modal_body" 
                    @save_size="save_size"
                    :wallpaper_sizes="default_wallpaper_sizes"/>
                <button slot="footer" @click="show_size_modal = false" class="button secondary">취소</button>
            </Modal>
        </div>
    `,
  data() {
    return {
      show_write_modal: false, // 등록 모달 노출여부
      show_size_modal: false, // 사이즈관리 모달 노출여부
      show_thumbnail_modal: false, // 썸네일 모달 노출여부
      popup_thumbnail_src: "", // 썸네일 모달 이미지 src
    };
  },
  created() {
    this.$store.dispatch("GET_CONTENTS"); // 컨텐츠 리스트 조회
    this.$store.dispatch("GET_DEFAULT_WALLPAPER_SIZES"); // default 사이즈 관리 정보 조회
  },
  computed: {
    content_count() {
      // 컨텐츠 총 개수
      return this.$store.getters.content_count;
    },
    contents() {
      // 컨텐츠 리스트
      return this.$store.getters.contents;
    },
    current_page() {
      // 현재 페이지
      return this.$store.getters.current_page;
    },
    last_page() {
      // 마지막 페이지
      return this.$store.getters.last_page;
    },
    default_wallpaper_sizes() {
      // default 배경화면 사이즈 리스트
      return this.$store.getters.default_wallpaper_sizes;
    },
    pop_content() {
      // 기존 컨텐츠
      return this.$store.getters.pop_content;
    },
  },
  methods: {
    get_start_date(start_date, open_yn) {
      // GET 시작일 오픈여부 노출 HTML
      return start_date + (open_yn ? '<span class="blue">(오픈)</span>' : '<span class="green">(오픈예정)</span>');
    },
    click_page(page) {
      // 페이지 클릭 이벤트
      this.$store.commit("SET_CURRENT_PAGE", page);
      this.$store.dispatch("GET_CONTENTS");
      window.scrollTo(0, 0);
    },
    do_search() {
      // 검색버튼 클릭 이벤트
      this.$store.commit("SET_SEARCH_PARAMETER", {
        gubun: document.getElementById("gubun").value,
        use_yn: document.getElementById("use_yn").value,
        title: document.getElementById("title").value,
      });
      this.$store.commit("SET_CURRENT_PAGE", 1);
      this.$store.dispatch("GET_CONTENTS");
    },
    change_page_size() {
      // 페이지당 컨텐츠 노출 수 변경
      this.$store.commit("SET_PAGE_SIZE", document.getElementById("page_size").value);
      this.$store.commit("SET_CURRENT_PAGE", 1);
      this.$store.dispatch("GET_CONTENTS");
    },
    popup_manage_content(e) {
      // popup 컨텐츠 관리
      if (e.target.tagName === "TD") {
        this.$store.dispatch("GET_CONTENT_ONE", e.target.innerText);
      } else if (e.target.id === "reg_new_content") {
        this.$store.commit("SET_POP_CONTENT", {});
      }
      this.show_write_modal = true;
    },
    popup_thumbnail(e) {
      // popup 썸네일 모달
      this.popup_thumbnail_src = e.target.src;
      this.show_thumbnail_modal = true;
    },
    save_content() {
      // 컨텐츠 저장
      const _this = this;
      const form_data = new FormData(document.hitchhiker_content);
      const api_data = {};
      form_data.forEach((value, key) => {
        if (key === "size_link_idxs" || key === "size_device_idxs" || key === "size_links") {
          if (api_data[key] === undefined) {
            api_data[key] = [value.trim()];
          } else {
            api_data[key].push(value.trim());
          }
        } else {
          api_data[key] = value;
        }
      });
      // PC, Mobile일 경우 배경화면 사이즈 링크 Array 변환
      if (
        api_data.size_link_idxs !== undefined &&
        api_data.size_device_idxs !== undefined &&
        api_data.size_links !== undefined &&
        (api_data.gubun === "1" || api_data.gubun === "2")
      ) {
        this.convert_size_array(api_data, api_data.size_link_idxs, api_data.size_device_idxs, api_data.size_links);
      }
      console.log(api_data);

      // 검증
      const validate = this.validate_content_data(api_data);
      if (!validate[0]) {
        alert(validate[1]);
        return false;
      }

      callApi("POST", "/hitchhiker/content", api_data, function (data) {
        console.log("callApi\n", data);
        alert("저장 되었습니다.");
        _this.$store.dispatch("GET_CONTENTS");
        _this.show_write_modal = false;
      });
    },
    validate_content_data(content) {
      // 컨텐츠 저장 검증
      // 구분값 검증
      if (content.gubun === "") return [false, "구분값을 선택 해 주세요."];

      // 시작일 검증
      let date_pattern = /^(19|20)\d{2}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[0-1])$/;
      if (content.start_date === "") return [false, "시작일을 입력 해 주세요."];
      if (!date_pattern.test(content.start_date)) return [false, "잘못된 시작일 형식 입니다."];

      // Mobile배경이 아닐 경우 타이틀 검증
      if (content.gubun !== "4" && content.title === "") return [false, "타이틀을 입력 해 주세요"];

      // Movie일 경우 상세 내용 검증
      if (content.gubun === "3" && content.detail_content === "") {
        return [false, "상세 내용을 입력 해 주세요"];
      }

      // Mobile배경일 경우 썸네일 검증
      if (content.gubun === "4" && content.thumbnail === "") return [false, "썸네일을 등록 해 주세요"];

      return [true, ""];
    },
    convert_size_array(api_data, link_idx_array, device_idx_array, link_array) {
      // 배경화면 사이즈 링크 Array 변환
      for (let i = 0; i < link_idx_array.length; i++) {
        api_data["wallpaper_sizes[" + i + "].link_idx"] = link_idx_array[i];
        api_data["wallpaper_sizes[" + i + "].device_idx"] = device_idx_array[i];
        api_data["wallpaper_sizes[" + i + "].link"] = link_array[i];
      }
    },
    popup_manage_wallpaper_size() {
      // 사이즈관리 모달 팝업
      this.show_size_modal = true;
    },
    save_size(size_data) {
      // 사이즈관리 저장
      this.$store.dispatch("SAVE_WALLPAPER_SIZE", size_data);
    },
    reload() {
      // 새로고침
      window.location.reload(true);
    },
  },
  watch: {
    default_wallpaper_sizes(sizes) {
      this.$refs.wallpaper_size_modal_body.set_default_data(sizes);
    },
  },
});
