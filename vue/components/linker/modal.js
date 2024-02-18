Vue.component('MODAL', {
    template : `
        <transition name="fade">
            <div v-if="show" class="modal">
                <div class="modal-overlay"></div>
                <div class="modal-wrap" :style="[{width: width + 'px'}]">
                    <button @click="closeModal" class="modal-close-btn"></button>
                    <div class="modal-title">
                        <h3>{{title}}</h3>
                    </div>
                    <div class="modal-container">
                        <slot name="body"></slot>
                    </div>
                </div>
            </div>
        </transition>
    `,
    data() {return {
        show : false, // 모달 노출 여부
    }},
    props : {
        title : { type:String, default: '' }, // 모달 제목
        width : { type:Number, default: 600 }, // 모달 Width
    },
    methods : {
        openModal() {
            this.show = true;
            this.$emit('openModal');
            document.body.style.overflow = 'hidden';
            document.body.style.height = '100%';
        },
        closeModal() {
            this.show = false;
            this.$emit('closeModal');
            document.body.style.overflow = '';
            document.body.style.height = '';
        },
    }
});