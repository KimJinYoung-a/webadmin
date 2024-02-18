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
                            <button @click="$emit('save')" class="button dark">����</button>
                            <button @click="$emit('close')" class="button secondary">���</button>
                        </slot>
                    </div>

                </div>
            </div>
        </div>
    </transition>
    `,
    props: {
        show_header_yn : {type:Boolean, default: true}, // ��� ���� ����
        show_footer_yn : {type:Boolean, default: true}, // Ǫ�� ���� ����
        close_background_click_yn : {type:Boolean, default: false}, // ��� Ŭ�� �� ���â ���� ����
        header_title : {type:String, default: 'Title'}, // ��� Ÿ��Ʋ

        /* Style */
        modal_width : {type:String, default: '600px'}, // ��� Width
    },
    computed : {
        container_style() { // .modal-container style
            return {
                width : this.modal_width
            }
        }
    },
    methods : {
        click_mask(e) { // ��� Ŭ�� �� ����
            if( this.close_background_click_yn && e.target.classList.contains('modal-wrapper') ) {
                this.$emit('close');
            }
        }
    }
});