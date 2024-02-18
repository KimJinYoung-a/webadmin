Vue.component('LIST-FORUM', {
    template : `
        <li @click="clickForum" :class="[{on : currentForumIdx === forum.forumIdx}]">
            <p class="number">{{forumNumber}}</p>
            <div class="info">
                <strong v-html="forum.title"></strong>
                <span>{{period}} / {{isOpen}}</span>
            </div>
        </li>
    `,
    props : {
        currentForumIdx : {  type:Number, default:0  }, // 현재 활성화된 포럼 일련번호
        //region forum 포럼
        forum : {
            forumIdx : { type:Number, default:0 }, // 포럼 일련번호
            title : { type:String, default:'' }, // 포럼 제목
            subTitle : { type:String, default:'' }, // 포럼 부제목
            startDate : { type:String, default:'' }, // 포럼 시작일자
            endDate : { type:String, default:'' }, // 포럼 종료일자
            useYn : { type:Boolean, default:false }, // 포럼 사용여부
            sortNo : { type:Number, default:0 }, // 포럼 노출순서
        },
        //endregion
    },
    computed : {
        //region forumNumber 포럼 번호
        forumNumber() {
            return (this.forum.forumIdx < 10 ? '0' : '') + this.forum.forumIdx;
        },
        //endregion
        //region period 오픈 기간
        period() {
            return this.getLocalDateTimeFormat(this.forum.startDate, 'yyyy-MM-dd')
                + ' ~ ' + this.getLocalDateTimeFormat(this.forum.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region isOpen 오픈 여부
        isOpen() {
            return this.forum.useYn ? '오픈' : '오픈안함';
        },
        //endregion
    },
    methods : {
        //region clickForum 포럼 클릭
        clickForum() {
            this.$emit('clickForum', this.forum.forumIdx);
        },
        //endregion
    }
});