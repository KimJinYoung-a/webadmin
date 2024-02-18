Vue.component('PAGINATION', {
    template : `
        <ul class="pagination">
            <li :class="{disabled:currentPage <= 1}">
                <a @click.self="clickPage(currentPage-1, $event)">&lt;</a>
            </li>
            <li v-for="page in pages" :class="{on:page === currentPage}">
                <a @click.self="clickPage(page, $event)">{{page}}</a>
            </li>
            <li :class="{disabled:currentPage >= lastPage}">
                <a @click.self="clickPage(currentPage+1, $event)">&gt;</a>
            </li>
        </ul>    
    `,
    props : {
        currentPage : { type:Number, default:0 }, // 현재 페이지
        lastPage : { type:Number, default:0 }, // 마지막 페이지
        showPageCount : { type:Number, default:5 }, // 보여줄 페이지 수
    },
    computed : {
        //region pages 페이지 리스트
        pages() {
            if( this.currentPage < 1 || this.lastPage <= 1 )
                return [1];

            const pageList = [];
            const startPage = Math.floor((this.currentPage-1)/this.showPageCount)*this.showPageCount + 1;
            let endPage = (startPage + this.showPageCount - 1);
            endPage = endPage < this.lastPage ? endPage : this.lastPage;

            for( let i=startPage ; i<=endPage ; i++ ) {
                pageList.push(i);
            }
            return pageList;
        },
        //endregion
    },
    methods : {
        //region clickPage 페이지 클릭
        clickPage(page, e) {
            if( e.target.parentElement.classList.contains('disabled') )
                return false;

            this.$emit('clickPage', page);
        },
        //endregion
    }
});