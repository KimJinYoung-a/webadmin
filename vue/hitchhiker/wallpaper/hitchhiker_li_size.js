/**
 * 히치하이커 사이즈 관리 Li
 */
Vue.component("Hitchhiker-Li-Size", {
  template: `
        <li class="wallpaper-size">
            <div style="margin-right: auto;">
                <span v-if="gubun === 'mobile'">대표기종: <input v-model="this_device_name" type="text" class="form-control inline small"></span>
                <span>사이즈: <input v-model="this_size" type="text" class="form-control inline small"></span>
                <span>우선순위: <input v-model="this_sort_number" type="text" class="form-control inline mini center"></span>
                <span>
                    사용 여부:
                    <div class="form-check">
                        <input v-model="this_use_yn" value="Y" :name="use_yn_name" :id="use_yn_true_id" class="form-check-input" type="radio">
                        <label class="form-check-label" :for="use_yn_true_id">Y</label>
                        <input v-model="this_use_yn" value="N" :name="use_yn_name" :id="use_yn_false_id" class="form-check-input" type="radio">
                        <label class="form-check-label" :for="use_yn_false_id">N</label>
                    </div>
                </span>
            </div>

            <button v-if="update_yn" @click="save_size" class="button secondary">수정</button>
            <button v-else @click="save_size" class="button dark">등록</button>
        </li>
    `,
  data() {
    return {
      this_device_name: this.device_name,
      this_size: this.size,
      this_sort_number: this.sort_number,
      this_use_yn: this.use_yn,
    };
  },
  props: {
    device_idx: { type: Number, default: 0 }, // 인덱스
    gubun: { type: String, default: "" }, // 구분(pc, mobile)
    device_name: { type: String, default: "" }, // 대표기종
    size: { type: String, default: "" }, // 사이즈
    sort_number: { type: [String, Number], default: "" }, // 우선순위
    use_yn: { type: String, default: "N" }, // 사용 여부
    update_yn: { type: Boolean, default: true }, // 수정여부
  },
  computed: {
    use_yn_true_id() {
      return this.gubun + "_use_yn_true_" + this.device_idx;
    },
    use_yn_false_id() {
      return this.gubun + "_use_yn_false_" + this.device_idx;
    },
    use_yn_name() {
      return this.gubun + "_use_yn_" + this.device_idx;
    },
  },
  methods: {
    // 사이즈 저장
    save_size() {
      this.$emit("save_size", {
        device_idx: this.device_idx,
        gubun: this.gubun,
        device_name: this.this_device_name,
        size: this.this_size,
        sort_number: this.this_sort_number,
        use_yn: this.this_use_yn,
      });
    },
    // SET 기본 data
    set_default_data() {
      this.this_device_name = "";
      this.this_size = "";
      this.this_sort_number = "";
      this.this_use_yn = "N";
    },
    // SET data
    set_size_data(size_data) {
      this.this_device_name = size_data.device_name;
      this.this_size = size_data.content_size;
      this.this_sort_number = size_data.sort_number;
      this.this_use_yn = size_data.use_yn;
      this.check_use_yn();
    },
    // 이상하게 radio버튼이 동적으로 체크가 안됨.. setTimeout을 걸어 강제로 체크함
    check_use_yn() {
      const _this = this;
      setTimeout(function () {
        if (_this.this_use_yn === "Y") {
          document.getElementById(_this.use_yn_true_id).checked = true;
        } else {
          document.getElementById(_this.use_yn_false_id).checked = true;
        }
      }, 100);
    },
  },
});
