var app = new Vue({
  el: "#app",
  store: store,
  template: `
        <div>
            <!-- �˻� ���̺� -->
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
                        <th>�˻�����</th>
                        <td style="text-align:left;">
                            <select id="gubun" class="form-control inline small">
                                <option value="">����</option>
                                <option value="1">PC</option>
                                <option value="2">MOBILE</option>
                                <option value="3">MOVIE</option>
                                <option value="4">MOBILE���</option>
                            </select>
                            <select id="use_yn" class="form-control inline small">
                                <option value="">��뿩��</option>
                                <option>Y</option>
                                <option>N</option>
                            </select>
                        </td>
                        <th>Ÿ��Ʋ</th>
                        <td style="text-align:left;">
                            <input id="title" type="text" class="form-control">
                        </td>
                        <td class="td-button align-right">
                            <button @click="reload" type="button" class="button secondary">�˻�����Reset</button>
                            <button @click="do_search" type="button" class="button dark">�˻�</button>
                        </td>
                    </tr>
                </thead>
            </table>
            <p class="p-table">
                <span>�˻���� : <strong>{{content_count}}</strong></span>
                <i class='fas fa-sync' @click="reload"></i>
                <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                    <option v-for="n in 5" :value="n*10">{{n*10}}���� ����</option>
                </select>
                <button id="modify_wallpaper_size" @click="popup_manage_wallpaper_size" type="button" class="button dark">������ ����</button>
                <button id="reg_new_content" @click="popup_manage_content" type="button" class="button dark">�ű� ���</button>
            </p>

            <!-- ����Ʈ ���̺� -->
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
                        <th>��ȣ</th>
                        <th>����</th>
                        <th>�����</th>
                        <th>Ÿ��Ʋ</th>
                        <th>��뿩��</th>
                        <th>������</th>
                        <th>�����</th>
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

            <!-- ������ -->
            <Pagination
                @click_page="click_page"
                :current_page="current_page"
                :last_page="last_page"
            ></Pagination>

            <!-- ���/���� ��� -->
            <Modal v-show="show_write_modal"
                @save="save_content" @close="show_write_modal = false"
                modal_width="830px" header_title="��ġ����Ŀ ������ ���"
            >
                <Hitchhiker-Content-Write slot="body" :content="pop_content" :wallpaper_sizes="default_wallpaper_sizes"/>
            </Modal>

            <!-- ����� ��� -->
            <Modal v-show="show_thumbnail_modal" @close="show_thumbnail_modal = false"
                modal_width="400px" :show_header_yn="false" :show_footer_yn="false"
                :close_background_click_yn="true"
            >
                <img width="100%" :src="popup_thumbnail_src" slot="body">
            </Modal>

            <!-- ��������� ��� -->
            <Modal v-show="show_size_modal" modal_width="870px" header_title="��ġ����Ŀ ���ȭ�� ������ ����">
                <Hitchhiker-Wallpaper-Size-Write
                    slot="body" ref="wallpaper_size_modal_body" 
                    @save_size="save_size"
                    :wallpaper_sizes="default_wallpaper_sizes"/>
                <button slot="footer" @click="show_size_modal = false" class="button secondary">���</button>
            </Modal>
        </div>
    `,
  data() {
    return {
      show_write_modal: false, // ��� ��� ���⿩��
      show_size_modal: false, // ��������� ��� ���⿩��
      show_thumbnail_modal: false, // ����� ��� ���⿩��
      popup_thumbnail_src: "", // ����� ��� �̹��� src
    };
  },
  created() {
    this.$store.dispatch("GET_CONTENTS"); // ������ ����Ʈ ��ȸ
    this.$store.dispatch("GET_DEFAULT_WALLPAPER_SIZES"); // default ������ ���� ���� ��ȸ
  },
  computed: {
    content_count() {
      // ������ �� ����
      return this.$store.getters.content_count;
    },
    contents() {
      // ������ ����Ʈ
      return this.$store.getters.contents;
    },
    current_page() {
      // ���� ������
      return this.$store.getters.current_page;
    },
    last_page() {
      // ������ ������
      return this.$store.getters.last_page;
    },
    default_wallpaper_sizes() {
      // default ���ȭ�� ������ ����Ʈ
      return this.$store.getters.default_wallpaper_sizes;
    },
    pop_content() {
      // ���� ������
      return this.$store.getters.pop_content;
    },
  },
  methods: {
    get_start_date(start_date, open_yn) {
      // GET ������ ���¿��� ���� HTML
      return start_date + (open_yn ? '<span class="blue">(����)</span>' : '<span class="green">(���¿���)</span>');
    },
    click_page(page) {
      // ������ Ŭ�� �̺�Ʈ
      this.$store.commit("SET_CURRENT_PAGE", page);
      this.$store.dispatch("GET_CONTENTS");
      window.scrollTo(0, 0);
    },
    do_search() {
      // �˻���ư Ŭ�� �̺�Ʈ
      this.$store.commit("SET_SEARCH_PARAMETER", {
        gubun: document.getElementById("gubun").value,
        use_yn: document.getElementById("use_yn").value,
        title: document.getElementById("title").value,
      });
      this.$store.commit("SET_CURRENT_PAGE", 1);
      this.$store.dispatch("GET_CONTENTS");
    },
    change_page_size() {
      // �������� ������ ���� �� ����
      this.$store.commit("SET_PAGE_SIZE", document.getElementById("page_size").value);
      this.$store.commit("SET_CURRENT_PAGE", 1);
      this.$store.dispatch("GET_CONTENTS");
    },
    popup_manage_content(e) {
      // popup ������ ����
      if (e.target.tagName === "TD") {
        this.$store.dispatch("GET_CONTENT_ONE", e.target.innerText);
      } else if (e.target.id === "reg_new_content") {
        this.$store.commit("SET_POP_CONTENT", {});
      }
      this.show_write_modal = true;
    },
    popup_thumbnail(e) {
      // popup ����� ���
      this.popup_thumbnail_src = e.target.src;
      this.show_thumbnail_modal = true;
    },
    save_content() {
      // ������ ����
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
      // PC, Mobile�� ��� ���ȭ�� ������ ��ũ Array ��ȯ
      if (
        api_data.size_link_idxs !== undefined &&
        api_data.size_device_idxs !== undefined &&
        api_data.size_links !== undefined &&
        (api_data.gubun === "1" || api_data.gubun === "2")
      ) {
        this.convert_size_array(api_data, api_data.size_link_idxs, api_data.size_device_idxs, api_data.size_links);
      }
      console.log(api_data);

      // ����
      const validate = this.validate_content_data(api_data);
      if (!validate[0]) {
        alert(validate[1]);
        return false;
      }

      callApi("POST", "/hitchhiker/content", api_data, function (data) {
        console.log("callApi\n", data);
        alert("���� �Ǿ����ϴ�.");
        _this.$store.dispatch("GET_CONTENTS");
        _this.show_write_modal = false;
      });
    },
    validate_content_data(content) {
      // ������ ���� ����
      // ���а� ����
      if (content.gubun === "") return [false, "���а��� ���� �� �ּ���."];

      // ������ ����
      let date_pattern = /^(19|20)\d{2}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[0-1])$/;
      if (content.start_date === "") return [false, "�������� �Է� �� �ּ���."];
      if (!date_pattern.test(content.start_date)) return [false, "�߸��� ������ ���� �Դϴ�."];

      // Mobile����� �ƴ� ��� Ÿ��Ʋ ����
      if (content.gubun !== "4" && content.title === "") return [false, "Ÿ��Ʋ�� �Է� �� �ּ���"];

      // Movie�� ��� �� ���� ����
      if (content.gubun === "3" && content.detail_content === "") {
        return [false, "�� ������ �Է� �� �ּ���"];
      }

      // Mobile����� ��� ����� ����
      if (content.gubun === "4" && content.thumbnail === "") return [false, "������� ��� �� �ּ���"];

      return [true, ""];
    },
    convert_size_array(api_data, link_idx_array, device_idx_array, link_array) {
      // ���ȭ�� ������ ��ũ Array ��ȯ
      for (let i = 0; i < link_idx_array.length; i++) {
        api_data["wallpaper_sizes[" + i + "].link_idx"] = link_idx_array[i];
        api_data["wallpaper_sizes[" + i + "].device_idx"] = device_idx_array[i];
        api_data["wallpaper_sizes[" + i + "].link"] = link_array[i];
      }
    },
    popup_manage_wallpaper_size() {
      // ��������� ��� �˾�
      this.show_size_modal = true;
    },
    save_size(size_data) {
      // ��������� ����
      this.$store.dispatch("SAVE_WALLPAPER_SIZE", size_data);
    },
    reload() {
      // ���ΰ�ħ
      window.location.reload(true);
    },
  },
  watch: {
    default_wallpaper_sizes(sizes) {
      this.$refs.wallpaper_size_modal_body.set_default_data(sizes);
    },
  },
});
