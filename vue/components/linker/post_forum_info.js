Vue.component('POST-FORUM-INFO', {
    template : `
        <div>
            <div class="forum-info-tab">
                <button @click="changePlatform('A')" :class="{on:platform === 'A'}">App</button>
                <button @click="changePlatform('M')" :class="{on:platform === 'M'}">Mobile Web</button>
                <button @click="changePlatform('P')" :class="{on:platform === 'P'}">PC</button>
            </div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:120px;">
                    <col>
                </colgroup>
                <tbody>
                    <!--region ����/���� -->
                    <template v-for="(info, key) in tempInfo">
                        <tr v-show="key === platform">
                            <th>����</th>
                            <td>
                                <textarea v-model="info.title" placeholder="���� �ȳ������� �Է����ּ���"></textarea>
                            </td>
                        </tr>
                        <tr v-show="key === platform">
                            <th>����(�ڵ�)</th>
                            <td>
                                <textarea v-model="info.content" rows="14" placeholder="���� ������ �ڵ�� �Է����ּ���"></textarea>
                            </td>
                        </tr>
                    </template>
                    <!--endregion-->
                    <!--region �����ڵ�-->
                    <tr>
                        <th>�����ڵ�</th>
                        <td>
                            <textarea class="forum-descr-sample" rows="6" wrap="off" :value="sampleCode" readonly></textarea>
                        </td>
                    </tr>
                    <!--endregion-->
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveInfo" class="linker-btn">����</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.modifyInfo )
            this.setTempInfo(this.modifyInfo);
    },
    data() {return {
        //region tempInfo ���/������ �ӽ� �ȳ� ����
        tempInfo : {
            'A' : { title:'', content:'' },
            'M' : { title:'', content:'' },
            'P' : { title:'', content:'' }
        },
        //endregion
        //region sampleCode ���� �ڵ�
        sampleCode : ''
            + '<div class="ex_img">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_01.gif" alt="since 2001">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_02.jpg" alt="�ٹ�����, �ӱ��ſ� �ϳ� �����ΰ�?">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_03.gif" alt="�ӱ���">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_04.jpg" alt="20��° �ӱ��� ���� ����!">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_05.gif" alt="�ӱ��ŵ��� ������ �����?">\n'
            + '</div>',
        //endregion
        platform : 'A',
    }},
    props : {
        forumIdx : { type:Number, default:0 },
        //region modifyInfo ������ �ȳ� ����
        modifyInfo : {
            infoIdx : { type:Number, default:0 },
            appTitle : { type:String, default:'' },
            appContent : { type:String, default:'' },
            mobileTitle : { type:String, default:'' },
            mobileContent : { type:String, default:'' },
            pcTitle : { type:String, default:'' },
            pcContent : { type:String, default:'' }
        },
        //endregion
    },
    methods : {
        //region setTempInfo Set �ӽ� �ȳ� ����
        setTempInfo(info) {
            this.tempInfo = {
                'A' : { title:info.appTitle, content:info.appContent },
                'M' : { title:info.mobileTitle, content:info.mobileContent },
                'P' : { title:info.pcTitle, content:info.pcContent }
            }
        },
        //endregion
        //region changePlatform �÷��� ����
        changePlatform(platform) {
            this.platform = platform;
        },
        //endregion
        //region saveInfo �ȳ� ���� ����
        saveInfo() {
            if( !this.validateInfoData() ) {
                return false;
            }
            const infoIdx = this.modifyInfo ? this.modifyInfo.infoIdx : null;
            const url = infoIdx ? '/linker/forum/info/update' : '/linker/forum/info';
            this.callApi(2, 'POST', url, this.createPostInfoData(infoIdx), this.successSaveInfo);
        },
        //region validateInfoData ���� �ȳ� �Է°� ����
        validateInfoData() {
            if( this.tempInfo.A.title.trim() === '' ) {
                alert('APP ������ �Է� �� �ּ���');
                return false;
            }
            if( this.tempInfo.M.title.trim() === '' ) {
                alert('Mobile Web ������ �Է� �� �ּ���');
                return false;
            }
            if( this.tempInfo.P.title.trim() === '' ) {
                alert('PC ������ �Է� �� �ּ���');
                return false;
            }
            if( this.tempInfo.A.content.trim() === '' ) {
                alert('APP ������ �Է� �� �ּ���');
                return false;
            }
            if( this.tempInfo.M.content.trim() === '' ) {
                alert('Mobile Web ������ �Է� �� �ּ���');
                return false;
            }
            if( this.tempInfo.P.content.trim() === '' ) {
                alert('PC ������ �Է� �� �ּ���');
                return false;
            }

            return true;
        },
        //endregion
        createPostInfoData(infoIdx) {
            const data = {
                forumIndex : this.forumIdx,
                appTitle : this.tempInfo.A.title,
                appContent : this.tempInfo.A.content,
                mobileTitle : this.tempInfo.M.title,
                mobileContent : this.tempInfo.M.content,
                pcTitle : this.tempInfo.P.title,
                pcContent : this.tempInfo.P.content
            };
            if( infoIdx )
                data.infoIndex = infoIdx;

            return data;
        },
        successSaveInfo(data) {
            alert('���� �Ǿ����ϴ�.');
            const modifyInfoIdx = this.modifyInfo ? this.modifyInfo.infoIdx : null;
            this.$emit('saveInfo', data, modifyInfoIdx);
        },
        //endregion
    }
})