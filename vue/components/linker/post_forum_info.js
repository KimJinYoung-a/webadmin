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
                    <!--region 제목/설명 -->
                    <template v-for="(info, key) in tempInfo">
                        <tr v-show="key === platform">
                            <th>제목</th>
                            <td>
                                <textarea v-model="info.title" placeholder="포럼 안내제목을 입력해주세요"></textarea>
                            </td>
                        </tr>
                        <tr v-show="key === platform">
                            <th>설명(코드)</th>
                            <td>
                                <textarea v-model="info.content" rows="14" placeholder="설명 내용을 코드로 입력해주세요"></textarea>
                            </td>
                        </tr>
                    </template>
                    <!--endregion-->
                    <!--region 샘플코드-->
                    <tr>
                        <th>샘플코드</th>
                        <td>
                            <textarea class="forum-descr-sample" rows="6" wrap="off" :value="sampleCode" readonly></textarea>
                        </td>
                    </tr>
                    <!--endregion-->
                </tbody>
            </table>

            <div class="modal-btn-area">
                <button @click="saveInfo" class="linker-btn">저장</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.modifyInfo )
            this.setTempInfo(this.modifyInfo);
    },
    data() {return {
        //region tempInfo 등록/수정용 임시 안내 정보
        tempInfo : {
            'A' : { title:'', content:'' },
            'M' : { title:'', content:'' },
            'P' : { title:'', content:'' }
        },
        //endregion
        //region sampleCode 샘플 코드
        sampleCode : ''
            + '<div class="ex_img">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_01.gif" alt="since 2001">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_02.jpg" alt="텐바이텐, 머그컵에 꽤나 진심인걸?">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_03.gif" alt="머그컵">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_04.jpg" alt="20번째 머그컵 드디어 공개!">\n'
            +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_05.gif" alt="머그컵들을 구경해 볼까요?">\n'
            + '</div>',
        //endregion
        platform : 'A',
    }},
    props : {
        forumIdx : { type:Number, default:0 },
        //region modifyInfo 수정할 안내 정보
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
        //region setTempInfo Set 임시 안내 정보
        setTempInfo(info) {
            this.tempInfo = {
                'A' : { title:info.appTitle, content:info.appContent },
                'M' : { title:info.mobileTitle, content:info.mobileContent },
                'P' : { title:info.pcTitle, content:info.pcContent }
            }
        },
        //endregion
        //region changePlatform 플랫폼 변경
        changePlatform(platform) {
            this.platform = platform;
        },
        //endregion
        //region saveInfo 안내 정보 저장
        saveInfo() {
            if( !this.validateInfoData() ) {
                return false;
            }
            const infoIdx = this.modifyInfo ? this.modifyInfo.infoIdx : null;
            const url = infoIdx ? '/linker/forum/info/update' : '/linker/forum/info';
            this.callApi(2, 'POST', url, this.createPostInfoData(infoIdx), this.successSaveInfo);
        },
        //region validateInfoData 포럼 안내 입력값 검증
        validateInfoData() {
            if( this.tempInfo.A.title.trim() === '' ) {
                alert('APP 제목을 입력 해 주세요');
                return false;
            }
            if( this.tempInfo.M.title.trim() === '' ) {
                alert('Mobile Web 제목을 입력 해 주세요');
                return false;
            }
            if( this.tempInfo.P.title.trim() === '' ) {
                alert('PC 제목을 입력 해 주세요');
                return false;
            }
            if( this.tempInfo.A.content.trim() === '' ) {
                alert('APP 내용을 입력 해 주세요');
                return false;
            }
            if( this.tempInfo.M.content.trim() === '' ) {
                alert('Mobile Web 내용을 입력 해 주세요');
                return false;
            }
            if( this.tempInfo.P.content.trim() === '' ) {
                alert('PC 내용을 입력 해 주세요');
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
            alert('저장 되었습니다.');
            const modifyInfoIdx = this.modifyInfo ? this.modifyInfo.infoIdx : null;
            this.$emit('saveInfo', data, modifyInfoIdx);
        },
        //endregion
    }
})