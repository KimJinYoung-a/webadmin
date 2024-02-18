Vue.component('Pagination',{
    template: `
        <ul class="pagination justify-content-center">
            <li :class="['page-item', {disabled: is_previous_disabled}]"><a class="page-link" @click="click_page(current_page-1)">Previous</a></li>
            <li v-for="page in page_list" :class="['page-item', {active: page === current_page}]">
                <a class="page-link" @click="click_page(page)">{{page}}</a>
            </li>
            <li :class="['page-item', {disabled: is_next_disabled}]"><a class="page-link" @click="click_page(current_page+1)">Next</a></li>
        </ul>
    `,
    data() { return {
        show_page : 5 // ������ ������ ��
    }},
    props : {
        current_page : {type: Number, default: 0}, // ���� ������
        last_page : {type: Number, default: 0}, // ������ ������
    },
    computed : {
        is_previous_disabled() { // ���� ������ disabled ����
            return this.current_page <= 1;
        },
        is_next_disabled() { // ���� ������ disabled ����
            return this.current_page >= this.last_page;
        },
        page_list() { // ������ ������ ����Ʈ
            const start_page = Math.floor((this.current_page-1)/this.show_page) * this.show_page + 1;
            const page_list = [];
            for( let i=0 ; i<this.show_page ; i++ ) {
                if( this.last_page < start_page + i )
                    break;
                page_list.push(start_page + i);
            }
            return page_list;
        }
    },
    methods : {
        click_page(page) {
            this.$emit('click_page', page);
        }
    }
});