/*
    히치하이커 배경화면 사이즈 관리 모달 Body
*/
Vue.component("Hitchhiker-Wallpaper-Size-Write", {
  template: `
    <div>
        <table class="table table-write table-dark">
            <colgroup>
                <col style="width:120px;">
                <col>
            </colgroup>
            <tbody>
                <tr>
                    <th>PC</th>
                    <td>
                        <ul class="ul-write-hitchhiker">
                            <Hitchhiker-Li-Size gubun="pc" @save_size="save_size"
                                :update_yn="false" ref="default_pc"/>
                            <li></li>
                            <Hitchhiker-Li-Size
                                v-for="(size, index) in pc_wallpaper_sizes" :key="index"
                                @save_size="save_size"
                                :device_idx="size.device_idx" gubun="pc" :size="size.content_size"
                                :sort_number="size.sort_number" :use_yn="size.use_yn"
                                ref="pc_size"
                            />
                        </ul>
                    </td>
                </tr>
                <tr>
                    <th>Mobile</th>
                    <td>
                        <ul class="ul-write-hitchhiker">
                            <Hitchhiker-Li-Size gubun="mobile" @save_size="save_size"
                                :update_yn="false" ref="default_mobile"/>
                            <li></li>
                            <Hitchhiker-Li-Size
                                v-for="(size, index) in mobile_wallpaper_sizes" :key="index"
                                @save_size="save_size"
                                :device_idx="size.device_idx" gubun="mobile" :size="size.content_size"
                                :device_name="size.device_name"
                                :sort_number="size.sort_number" :use_yn="size.use_yn"
                                ref="mobile_size"
                            />
                        </ul>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    `,
  data() {
    return {
      mobile_wallpaper_sizes: [], // Mobile
      pc_wallpaper_sizes: [], // PC
    };
  },
  methods: {
    set_default_data(sizes) {
      // props 사이즈 배열 -> data
      this.mobile_wallpaper_sizes = sizes.mobile_wallpaper_sizes;
      this.pc_wallpaper_sizes = sizes.pc_wallpaper_sizes;

      const _this = this;
      if (this.$refs.pc_size !== undefined) {
        this.$refs.pc_size.forEach((ref, index) => {
          ref.set_size_data(_this.pc_wallpaper_sizes[index]);
        });
      }
      if (this.$refs.mobile_size !== undefined) {
        this.$refs.mobile_size.forEach((ref, index) => {
          ref.set_size_data(_this.mobile_wallpaper_sizes[index]);
        });
      }

      this.$refs.default_pc.set_default_data();
      this.$refs.default_mobile.set_default_data();
    },
    save_size(size_data) {
      // 사이즈 등록/수정
      this.$emit("save_size", size_data);
    },
  },
});
