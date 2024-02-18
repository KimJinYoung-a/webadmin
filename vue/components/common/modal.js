Vue.component('Modal',{
    template: `
    <transition name="modal">
        <div class="modal-mask" @click="click_mask">
            <div class="modal-wrapper">
                <div class="modal-container" :style="container_style">

                    <div v-if="show_header_yn" class="modal-header">
                        <slot name="header">
                            <h5 class="modal-title">{{header_title}}</h5>
                        </slot>
                    </div>

                    <div class="modal-body">
                        <slot name="body">
                            default body
                        </slot>
                    </div>

                    <div v-if="show_footer_yn" class="modal-footer">
                        <slot name="footer">
                            <button @click="$emit('save')" class="button dark">저장</button>
                            <button @click="$emit('close')" class="button secondary">취소</button>
                        </slot>
                    </div>

                </div>
            </div>
        </div>
    </transition>
    `,
    props: {
        show_header_yn : {type:Boolean, default: true}, // 헤더 노출 여부
        show_footer_yn : {type:Boolean, default: true}, // 푸터 노출 여부
        close_background_click_yn : {type:Boolean, default: false}, // 배경 클릭 시 모달창 닫음 여부
        header_title : {type:String, default: 'Title'}, // 헤더 타이틀

        /* Style */
        modal_width : {type:String, default: '600px'}, // 모달 Width
    },
    computed : {
        container_style() { // .modal-container style
            return {
                width : this.modal_width
            }
        }
    },
    methods : {
        click_mask(e) { // 배경 클릭 시 닫음
            if( this.close_background_click_yn && e.target.classList.contains('modal-wrapper') ) {
                this.$emit('close');
            }
        }
    }
});