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
        currentForumIdx : {  type:Number, default:0  }, // ���� Ȱ��ȭ�� ���� �Ϸù�ȣ
        //region forum ����
        forum : {
            forumIdx : { type:Number, default:0 }, // ���� �Ϸù�ȣ
            title : { type:String, default:'' }, // ���� ����
            subTitle : { type:String, default:'' }, // ���� ������
            startDate : { type:String, default:'' }, // ���� ��������
            endDate : { type:String, default:'' }, // ���� ��������
            useYn : { type:Boolean, default:false }, // ���� ��뿩��
            sortNo : { type:Number, default:0 }, // ���� �������
        },
        //endregion
    },
    computed : {
        //region forumNumber ���� ��ȣ
        forumNumber() {
            return (this.forum.forumIdx < 10 ? '0' : '') + this.forum.forumIdx;
        },
        //endregion
        //region period ���� �Ⱓ
        period() {
            return this.getLocalDateTimeFormat(this.forum.startDate, 'yyyy-MM-dd')
                + ' ~ ' + this.getLocalDateTimeFormat(this.forum.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region isOpen ���� ����
        isOpen() {
            return this.forum.useYn ? '����' : '���¾���';
        },
        //endregion
    },
    methods : {
        //region clickForum ���� Ŭ��
        clickForum() {
            this.$emit('clickForum', this.forum.forumIdx);
        },
        //endregion
    }
});