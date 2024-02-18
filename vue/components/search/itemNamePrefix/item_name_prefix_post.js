Vue.component('ITEM-NAME-PREFIX-POST', {
    template : `
        <div>
            <table class="modal-write-tbl">
                <colgroup>
                    <col style="width:100px;">
                    <col>
                </colgroup>
                <tbody>
                    <tr>
                        <th>말머리</th>
                        <td><input v-model="prefixWord" type="text" placeholder="말머리를 입력해주세요"></td>
                    </tr>
                    <tr>
                        <th>기간</th>
                        <td>
                            <span class="datepicker">
                                <label for="startDate">
                                    <strong>시작일</strong>
                                </label>
                                <DATE-PICKER @updateDate="updateStartDate" :date="startDate" id="startDate"/>
                                <input v-model="startTime" type="time">
                            </span>
                            <span class="datepicker">
                                <label for="endDate">
                                    <strong>종료일</strong>
                                </label>
                                <DATE-PICKER @updateDate="updateEndDate" :date="endDate" id="endDate"/>
                                <input v-model="endTime" type="time">
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <th>사용여부</th>
                        <td>
                            <p class="radio-area">
                                <input v-model="use" :value="true" id="itemNamePrefixUseY" type="radio">
                                <label for="itemNamePrefixUseY">Y</label>
                                <input v-model="use" :value="false" id="itemNamePrefixUseN" type="radio">
                                <label for="itemNamePrefixUseN">N</label>
                            </p>
                        </td>
                    </tr>
                </tbody>
            </table>
            
            <!--region 등록/수정 실패 사유-->
            <div v-if="failProducts.length > 0" class="modal-alert">
                <a @click="failProducts = []" class="close">x</a>
                <strong>실패 사유</strong>
                <p v-for="product in failProducts">
                    - 상품코드 : {{product.duplicatedProductId}}, 
                    상품명 : "{{product.duplicatedProductName}}" 
                    > 말머리 [{{product.duplicatedPrefixWord}}]과 이벤트기간 겹침
                </p>
            </div>
            <!--endregion-->
            
            <div class="modal-btn-area">
                <button @click="postPrefix" class="btn">저장</button>
            </div>
        </div>
    `,
    mounted() {
        if( this.updatePrefix ) {
            this.setUpdatePrefixData();
        }
    },
    data() {return {
        prefixWord : '', // 말머리
        startDate : '', // 시작 일자
        startTime : '', // 시작 시간
        endDate : '', // 종료 일자
        endTime : '', // 종료 시간
        use : false, // 사용 여부
        failProducts : [], // 실패한 상품 리스트
    }},
    props : {
        //region updatePrefix 수정할 말머리
        updatePrefix : {
            prefixIdx : { type:Number, default:0 },
            prefixWord : { type:String, default:'' },
            startDate : { type:String, default:'' },
            endDate : { type:String, default:'' },
            use : { type:String, default:'' },
            state : { type:String, default:'' },
            itemCount : { type:Number, default:0 },
            regDate : { type:String, default:'' },
            regAdminId : { type:String, default:'' },
            regAdminName : { type:String, default:'' }
        },
        //endregion
    },
    computed : {
        //region postPrefixApiSendData 말머리 저장 API 전송 데이터
        postPrefixApiSendData() {
            const data = {
                prefixWord : this.prefixWord,
                startDate : this.startDate + ' ' + this.startTime,
                endDate : this.endDate + ' ' + this.endTime,
                use : this.use
            };

            if( this.updatePrefix )
                data.prefixIdx = this.updatePrefix.prefixIdx;

            return data;
        },
        //endregion
    },
    methods : {
        //region postPrefix 말머리 저장
        postPrefix() {
            if( !confirm('저장 하시겠습니까?') )
                return false;

            this.callApi(1, 'POST', '/search/prefix', this.postPrefixApiSendData, this.successPostPrefix);
        },
        successPostPrefix(failProducts) {
            if( failProducts.length > 0 ) {
                this.failProducts = failProducts;
            } else {
                this.$emit('postPrefix');
            }
        },
        //endregion
        //region setUpdatePrefixData 수정 말머리 데이터 설정
        setUpdatePrefixData() {
            this.prefixWord = this.updatePrefix.prefixWord;
            this.use = this.updatePrefix.use;
            this.setUpdatePrefixDateAndTime('start', this.updatePrefix.startDate);
            this.setUpdatePrefixDateAndTime('end', this.updatePrefix.endDate);
        },
        setUpdatePrefixDateAndTime(flag, date) {
            const dateArr = date.split('T');
            this[`${flag}Date`] = dateArr[0];
            this[`${flag}Time`] = dateArr[1];
        },
        //endregion
        //region updateStartDate 시작일 수정
        updateStartDate(date) {
            this.startDate = date;
        },
        //endregion
        //region updateEndDate 종료일 수정
        updateEndDate(date) {
            this.endDate = date;
        },
        //endregion
    }
});