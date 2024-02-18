Vue.component('List-Write',{
    template: `
        <div style="height: 600px; overflow-y:auto;">
            <form name="play_content" id="play_content" enctype="multipart/form-data">
                <input v-if="current_content.pidx != 0" name="pidx" type="hidden" v-model="current_content.pidx">
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>레이아웃 *</th>
                            <td>
                                <select name="uinumber" class="form-control inline small must" v-model="current_content.uinumber" @change="changeUiNumber">
                                    <option value="0">전체</option>
                                    <option value="1">리스트형</option>
                                    <option value="2">상세형</option>
                                    <option value="4">동영상형</option>
                                    <option value="5">이벤트형</option>
                                </select>
                            </td>
                        </tr>
                        <tr v-show="showF.cidxF">
                            <th>컨텐츠 *</th>
                            <td>
                                <select name="cidx" class="form-control inline small must" v-model="current_content.cidx">
                                    <option value="0">전체</option>
                                    <option value="1">마스터피스</option>
                                    <option value="2">탐구생활</option>
                                    <option value="3">DAY.FILM</option>
                                    <option value="4">THING.배지</option>
                                    <option value="5">PLAY.GOODS</option>
                                    <option value="7">WEEKLY WALLPAPER</option>
                                </select>
                            </td>
                        </tr>                        
                    </tbody>
                </table>
                
                <br/>   
                <h2>리스트 노출 구성</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th id="tt">제목 *</th>
                            <td>
                                <textarea name="titlename" style="width: 100%;" rows="3" v-model="current_content.titlename" class="must"></textarea>
                                - 줄바꿈으로 두 줄까지 입력 가능하며, 한 줄당 한글기준 8자까지만 입력해 주세요.
                            </td>
                        </tr>
                        <tr>
                            <th>리스트 이미지</th>
                            <td>
                                <label for="addListimage">등록</label> <br/> * jpg, gif 포맷의 파일만 등록 가능합니다.
                                <div v-show="current_content.listimage" class="thumbnail-area" @click="delete_list_image">
                                    <img id="showListImage" :src="current_content.listimage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">삭제</div>
                                </div>
                                <input type="text" name="listimage" v-model="current_content.listimage" style="display: none;" />                                
                                <input type="text" name="listimageChangeF" value="N" style="display: none" />
                            </td>
                        </tr>
                        <tr v-show="showF.badgeF">
                            <th>뱃지 *</th>
                            <td>
                                <input type="radio" name="bedgeflag" value="0" checked/> 안함
                                <input type="radio" name="bedgeflag" value="1"/> 신상품
                                <input type="radio" name="bedgeflag" value="2"/> 한정상품
                            </td>
                        </tr>
                        <tr v-show="showF.contentsF">
                            <th>내용 <p v-show="current_content.uinumber == 2 || current_content.uinumber == 4" style="display: contents;">*</p></th>
                            <td>
                                <textarea v-model="current_content.contents" @input="change_contents_text" name="contents" style="width: 100%;" rows="5" :class="{must : checkMustClass('contents')}"></textarea>
                                <p style="text-align: right;">{{contents_text_length}}/800</p>
                                - 한글기준 최대 800자까지 입력 가능합니다.
                            </td>
                        </tr>
                        <tr v-show="showF.tagF">
                            <th>태그 *</th>
                            <td>
                                <p>- 백스페이스로 태그 삭제가 가능합니다.</p>
                                <div id="tagDiv">
                                    <p v-for="(item, index) in pop_content_tag" @click="delete_tag(index)" name="tagP" style="float:left; margin-right: 5px;">{{item}}</p>
                                </div>
                                <input type="text" name="tagInput" @keyup.enter.space="tagInsert" @keydown.delete="checkTagInputEmpty" @keyup.delete="tagDelete"/>
                                <input type="text" name="tagP" id="tagP" v-model="tagp" style="display: none" :class="{must : checkMustClass('tagP')}"/>        
                                <input type="text" name="tagPChangeF" value="N" style="display: none" />
                            </td>
                        </tr>
                        <tr v-show="showF.itemidF">
                            <th>연관상품 <p v-show="current_content.uinumber == 1" style="display: contents;">*</p></th>
                            <td>
                                <input type="button" @click="reg_itemid" value="등록"/><br/>
                                <div id="thumnailDiv" v-show="pop_content_items">
                                    <div class="thumbnail-area" v-for="(item, index) in pop_content_items" style="display: inline-block;" @click="delete_itemid(item.itemid, index)">
                                        <img name="itemidImg" :src="item.itemimage" />
                                        <div class="overlay">삭제</div>
                                    </div>
                                </div>                                
                                <p>- 연관상품은 최대 10건까지 등록 가능합니다.</p>
                                <input type="text" name="itemid" id="itemid" :value="itemid" style="display: none" :class="{must : checkMustClass('itemid')}" />
                                <input type="text" name="itemidChangeF" value="N" style="display: none" />
                            </td>
                        </tr>
                        <tr v-show="showF.winBadgeF">
                            <th>당첨자 안내 뱃지</th>
                            <td>
                                <input type="checkbox" name="winbadge" />사용
                                <input type="text" name="winbadgestdate" /> 부터 <input type="text" name="winbedgeeddate" /> 까지 이벤트 중
                                <input type="text" name="winnerissue" />
                            </td>
                        </tr>
                        <tr v-show="showF.linkUrlF">
                            <th>링크 *</th>
                            <td>
                                <input type="text" name="linkurl" v-model="current_content.linkurl" :class="{must : checkMustClass('linkurl')}"/>
                            </td>
                        </tr>
                    </tbody>
                </table>
                
                <br/> 
                <h2 v-show="showF.headF">내용 구성</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr v-show="showF.detailCaseF">
                            <th>Excute File Path</th>
                            <td>
                                <textarea name="excutepath" style="width: 100%;" rows="5" v-model="current_content.excutepath"></textarea>
                                - 개발자 전용. 있을 경우에만 입력해 주세요.
                            </td>
                        </tr>
                        <tr v-show="showF.detailCaseF">
                            <th>HTML</th>
                            <td>
                                <textarea name="htmlcode" style="width: 100%;" rows="5" v-model="current_content.htmlcode"></textarea>
                                - 한글 기준 최대 800자 까지 입력 가능합니다.
                            </td>
                        </tr>    
                        <tr v-show="showF.videoUrlF">
                            <th>동영상 코드 *</th>
                            <td>
                                <input type="text" name="videourl" v-model="current_content.videourl" :class="{must : checkMustClass('videourl')}" />
                                - Youtube URL을 입력해주세요.
                            </td>
                        </tr>                     
                    </tbody>
                </table>
                
                <br/> 
                <h2>게시 설정</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>오프닝 *</th>
                            <td>
                                <input type="radio" name="openingflag" value="0" checked/>안함
                                <input type="radio" name="openingflag" value="1" />오프닝 1
                                <input type="radio" name="openingflag" value="2" />오프닝 2
                                <input type="radio" name="openingflag" value="3" />오프닝 3
                            </td>
                        </tr>
                        <tr>
                            <th>게시일 *</th>
                            <td>
                                <input type="checkbox" name="isaod" v-model="current_content.isaod" true-value="1" fasle-value="0" value="1" />시작일부터 상시 노출
                                <input type="text" name="startdate" id="start_date" v-model="current_content.startdate" class="must"/> 부터 <input type="text" name="enddate" id="end_date" v-model="current_content.enddate" class="must"/> 까지
                            </td>
                        </tr>                        
                    </tbody>
                </table>
                
                <br/> 
                <h2>운영 정보</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>메모</th>
                            <td>
                                <textarea name="ordertext" style="width: 100%;" rows="5" v-model="current_content.ordertext"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>진행상황 *</th>
                            <td>
                                <select name="stateflag" class="form-control inline small must" v-model="current_content.stateflag">
                                    <option value="0">선택</option>
                                    <option value="1">등록대기</option>
                                    <option value="2">디자인요청</option>
                                    <option value="3">퍼블리싱요청</option>
                                    <option value="4">개발요청</option>
                                    <option value="5">오픈요청</option>
                                    <option value="7">오픈</option>
                                    <option value="8">보류</option>
                                    <option value="9">종료</option>
                                </select>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </form>
            
            <input name="addListimage" id="addListimage" @change="change_list_image($event)" type="file" class="form-control inline middle" style="display: none;" />
            <input type="text" name="additemid" id="additemid" v-model="additemid" style="display: none"/>
            <input type="text" name="additemimage" id="additemimage" v-model="additemimage" style="display: none"/>
        </div>
    `,
    mounted() {
        const _this = this;

        const arrDayMin = ["일","월","화","수","목","금","토"];
        const arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
        $("#start_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                const min_date = $(this).datepicker("getDate");
                $("#end_date").datepicker('setDate', min_date);
                $("#end_date").datepicker('option', "minDate", min_date);

                _this.current_content.startdate = document.getElementById("start_date").value;
                _this.current_content.enddate = document.getElementById("end_date").value;
            }
        });
        $("#end_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '이전달', nextText: '다음달', yearSuffix: '년',
            dayNamesMin: arrDayMin,
            monthNames: arrMonth,
            showMonthAfterYear: true
            , onSelect : function (date){
                _this.current_content.enddate = document.getElementById("end_date").value;
            }
        });
    },
    data() {
        return {
            current_content : {
                pidx : 0
                , uinumber : 0
                , cidx: 0
                , stateflag: 0
                , itemid: ""
                , changeflag : false
                , listimage : ""
                , contents : ""
            }
            , showF: {
                cidxF: false
                , badgeF: false
                , contentsF: false
                , tagF: false
                , itemidF: false
                , winBadgeF: false
                , detailCaseF: false
                , videoUrlF: false
                , linkUrlF: false
                , headF: false
            }
            , postTagEmpty: false
            , additemid: ""
            , additemimage: ""
            , tagp: ""
            , itemid: ""
            , contents_text_length : 0
        }
    },
    props: {
        pop_content : {
            pidx : {type:Number, default:0} // 컨텐츠 idx
            , uinumber : {type:Number, default:0}
        }
        , pop_content_items : {type:Array, default: null}
        , pop_content_tag : {type:Array, default: null}
    },
    methods : {
        changeUiNumber(event){
            this.initShowF();
            var uiNumber = event.target.value;
            this.changeShowF(uiNumber);
        }
        , initShowF(){
            this.showF.cidxF = false;
            this.showF.badgeF = false;
            this.showF.contentsF = false;
            this.showF.tagF = false;
            this.showF.itemidF = false;
            this.showF.winBadgeF = false;
            this.showF.detailCaseF = false;
            this.showF.videoUrlF = false;
            this.showF.linkUrlF = false;
            this.showF.headF = false;
        }
        , changeShowF(uiNumber){
            if(uiNumber == 0){

            }else if(uiNumber == 1){ //리스트형
                this.showF.cidxF = true;
                this.showF.badgeF = true;
                this.showF.contentsF = true;
                this.showF.tagF = true;
                this.showF.itemidF = true;
            }else if(uiNumber == 2){
                this.showF.cidxF = true;
                this.showF.badgeF = true;
                this.showF.contentsF = true;
                this.showF.winBadgeF = true;
                this.showF.detailCaseF = true;
                this.showF.headF = true;
            }else if(uiNumber == 4){
                this.showF.cidxF = true;
                this.showF.badgeF = true;
                this.showF.contentsF = true;
                this.showF.itemidF = true;
                this.showF.videoUrlF = true;
                this.showF.headF = true;
            }else if(uiNumber == 5){
                this.showF.linkUrlF = true;
            }
        }
        , checkMustClass(name){
            var uinumber = this.current_content.uinumber;

            if(uinumber == 1){ //리스트형
                if(name == 'tagP' || name == 'itemid'){
                    return true;
                }else return false
            }else if(uinumber == 2){
                if(name == 'contents') return true;
                else return false;
            }else if(uinumber == 4){ // 동영상형
                if(name == 'videourl' || name == 'contents'){
                    return true;
                }else return false;
            }else if(uinumber == 5){ // 이벤트형
                if(name == 'linkurl') return true;
                else return false;
            }
        }
        , tagInsert(tag){
            var _this = this;
            $(function(){
                if(tag.target.value.trim() != ""){
                    return new Promise(function(resolve, reject){
                        _this.pop_content_tag.push(tag.target.value.trim());
                        $("input[name=tagInput]").val("");
                        resolve("");
                    }).then(function (data){
                        this.tagp += this.tagp + tag.target.value.trim();
                        $("input[name=tagPChangeF]").val("Y");
                    });
                }else{
                    $("input[name=tagInput]").val("");
                }
            });
        }
        , tagDelete(tag){
            if(this.postTagEmpty){
                this.pop_content_tag.pop();
                this.change_tag_string();
                $("input[name=tagPChangeF]").val("Y");
            }
        }
        , delete_tag(index){
            const _this = this;
            let content_tag = this.pop_content_tag.slice(0, index).concat(this.pop_content_tag.slice(index+1, _this.pop_content_tag.length));
            this.$emit("change_content_tag", content_tag);

            this.change_tag_string();
            $("input[name=tagPChangeF]").val("Y");
        }
        , checkTagInputEmpty(tag){
            if(tag.target.value.trim() == ""){
                this.postTagEmpty = true;
            }else{
                this.postTagEmpty = false;
            }
        }
        , change_tag_string(){
            this.tagp = "," + this.pop_content_tag.toString();
        }
        , change_list_image(image){
            const _this = this;
            const file = image.target.files[0];

            if (!file.type.match("image.*")) {
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = function(e){
                _this.current_content.listimage = e.target.result;
            }

            $("input[name=listimageChangeF]").val("Y");
            console.log("check", this.current_content);
        }
        , delete_list_image(){
            if(confirm("제거하시겠습니까?")){
                this.current_content.listimage = "";
                $("#addListimage").val("");
            }
        }
        , reg_itemid(){
            window.open("/admin/sitemaster/piece/pop_singleItemSelect_V2.asp?target=play_content&ptype=piece&itemarr=", "연관상품", "width:300px, height:200px");
        }
        , delete_itemid(itemid, index){
            var preItemid = this.itemid.replace("," + itemid, "");
            this.itemid = preItemid;
            this.pop_content_items.splice(index, 1);
            $("input[name=itemidChangeF]").val("Y");

            this.pop_content.changeflag = true;
        }
        , change_contents_text(e){
            //this.contents_text_length = e.target.value.length;
            this.contents_text_length = this.current_content.contents.length;

            if(this.contents_text_length > 800){
                alert("800자를 초과했습니다.");
            }
        }
    }
    , watch:{
        pop_content(popcontent) { // 컨텐츠 변경 시
            const _this = this;

            //console.log("check alter");
            $("input[name=tagPChangeF]").val("N");
            $("input[name=itemidChangeF]").val("N");
            this.postTagEmpty = true;
            $("input[name=listimageChangeF]").val("N");

            if(popcontent.pidx != null){
                this.current_content = popcontent;
                this.current_content.changeflag = false;
                this.initShowF();
                this.changeShowF(this.current_content.uinumber);

                var bedgflag = this.current_content.bedgeflag;
                var openingflag = this.current_content.openingflag;

                $("input[name=bedgeflag]").each(function(){
                    if(this.value == bedgflag){
                        this.checked = true;
                    }
                });

                $("input[name=openingflag]").each(function(){
                    if(this.value == openingflag){
                        this.checked = true;
                    }
                });

                this.change_tag_string();
            }else{
                this.current_content = {
                    pidx : 0
                    , uinumber : 0
                    , cidx: 0
                    , stateflag: 0
                    , itemid: ""
                    , changeflag : false
                    , listimage : ""
                };
                this.itemid = "";
                this.tagp = "";
                this.initShowF();

                $("input[name=bedgeflag]:first").prop("checked", true);
                $("input[name=openingflag]:first").prop("checked", true);

                $("#thumnailDiv").empty();
                $("p[name=tagP]").remove();
            }

            this.contents_text_length = this.current_content.contents.length;
        }
        , pop_content_tag(pop_content_tag){
            console.log("watch pop_content_tag");
            this.change_tag_string();
        }
        , pop_content_items(pop_content_items){
            console.log("watch pop_content_items");
            var itemidString = "";
            pop_content_items.forEach(function (item){
                itemidString += "," + item.itemid;
            });
            this.itemid = itemidString;
        }
        , additemid(additemid){
            var object = {"itemid": additemid, "itemimage": this.additemimage};
            this.pop_content_items.push(object);

            this.itemid += additemid;
            $("input[name=itemidChangeF]").val("Y");
            this.pop_content.changeflag = true;
        }
    }
});