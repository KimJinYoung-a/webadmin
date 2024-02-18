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
                            <th>���̾ƿ� *</th>
                            <td>
                                <select name="uinumber" class="form-control inline small must" v-model="current_content.uinumber" @change="changeUiNumber">
                                    <option value="0">��ü</option>
                                    <option value="1">����Ʈ��</option>
                                    <option value="2">����</option>
                                    <option value="4">��������</option>
                                    <option value="5">�̺�Ʈ��</option>
                                </select>
                            </td>
                        </tr>
                        <tr v-show="showF.cidxF">
                            <th>������ *</th>
                            <td>
                                <select name="cidx" class="form-control inline small must" v-model="current_content.cidx">
                                    <option value="0">��ü</option>
                                    <option value="1">�������ǽ�</option>
                                    <option value="2">Ž����Ȱ</option>
                                    <option value="3">DAY.FILM</option>
                                    <option value="4">THING.����</option>
                                    <option value="5">PLAY.GOODS</option>
                                    <option value="7">WEEKLY WALLPAPER</option>
                                </select>
                            </td>
                        </tr>                        
                    </tbody>
                </table>
                
                <br/>   
                <h2>����Ʈ ���� ����</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th id="tt">���� *</th>
                            <td>
                                <textarea name="titlename" style="width: 100%;" rows="3" v-model="current_content.titlename" class="must"></textarea>
                                - �ٹٲ����� �� �ٱ��� �Է� �����ϸ�, �� �ٴ� �ѱ۱��� 8�ڱ����� �Է��� �ּ���.
                            </td>
                        </tr>
                        <tr>
                            <th>����Ʈ �̹���</th>
                            <td>
                                <label for="addListimage">���</label> <br/> * jpg, gif ������ ���ϸ� ��� �����մϴ�.
                                <div v-show="current_content.listimage" class="thumbnail-area" @click="delete_list_image">
                                    <img id="showListImage" :src="current_content.listimage" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                                    <div class="overlay">����</div>
                                </div>
                                <input type="text" name="listimage" v-model="current_content.listimage" style="display: none;" />                                
                                <input type="text" name="listimageChangeF" value="N" style="display: none" />
                            </td>
                        </tr>
                        <tr v-show="showF.badgeF">
                            <th>���� *</th>
                            <td>
                                <input type="radio" name="bedgeflag" value="0" checked/> ����
                                <input type="radio" name="bedgeflag" value="1"/> �Ż�ǰ
                                <input type="radio" name="bedgeflag" value="2"/> ������ǰ
                            </td>
                        </tr>
                        <tr v-show="showF.contentsF">
                            <th>���� <p v-show="current_content.uinumber == 2 || current_content.uinumber == 4" style="display: contents;">*</p></th>
                            <td>
                                <textarea v-model="current_content.contents" @input="change_contents_text" name="contents" style="width: 100%;" rows="5" :class="{must : checkMustClass('contents')}"></textarea>
                                <p style="text-align: right;">{{contents_text_length}}/800</p>
                                - �ѱ۱��� �ִ� 800�ڱ��� �Է� �����մϴ�.
                            </td>
                        </tr>
                        <tr v-show="showF.tagF">
                            <th>�±� *</th>
                            <td>
                                <p>- �齺���̽��� �±� ������ �����մϴ�.</p>
                                <div id="tagDiv">
                                    <p v-for="(item, index) in pop_content_tag" @click="delete_tag(index)" name="tagP" style="float:left; margin-right: 5px;">{{item}}</p>
                                </div>
                                <input type="text" name="tagInput" @keyup.enter.space="tagInsert" @keydown.delete="checkTagInputEmpty" @keyup.delete="tagDelete"/>
                                <input type="text" name="tagP" id="tagP" v-model="tagp" style="display: none" :class="{must : checkMustClass('tagP')}"/>        
                                <input type="text" name="tagPChangeF" value="N" style="display: none" />
                            </td>
                        </tr>
                        <tr v-show="showF.itemidF">
                            <th>������ǰ <p v-show="current_content.uinumber == 1" style="display: contents;">*</p></th>
                            <td>
                                <input type="button" @click="reg_itemid" value="���"/><br/>
                                <div id="thumnailDiv" v-show="pop_content_items">
                                    <div class="thumbnail-area" v-for="(item, index) in pop_content_items" style="display: inline-block;" @click="delete_itemid(item.itemid, index)">
                                        <img name="itemidImg" :src="item.itemimage" />
                                        <div class="overlay">����</div>
                                    </div>
                                </div>                                
                                <p>- ������ǰ�� �ִ� 10�Ǳ��� ��� �����մϴ�.</p>
                                <input type="text" name="itemid" id="itemid" :value="itemid" style="display: none" :class="{must : checkMustClass('itemid')}" />
                                <input type="text" name="itemidChangeF" value="N" style="display: none" />
                            </td>
                        </tr>
                        <tr v-show="showF.winBadgeF">
                            <th>��÷�� �ȳ� ����</th>
                            <td>
                                <input type="checkbox" name="winbadge" />���
                                <input type="text" name="winbadgestdate" /> ���� <input type="text" name="winbedgeeddate" /> ���� �̺�Ʈ ��
                                <input type="text" name="winnerissue" />
                            </td>
                        </tr>
                        <tr v-show="showF.linkUrlF">
                            <th>��ũ *</th>
                            <td>
                                <input type="text" name="linkurl" v-model="current_content.linkurl" :class="{must : checkMustClass('linkurl')}"/>
                            </td>
                        </tr>
                    </tbody>
                </table>
                
                <br/> 
                <h2 v-show="showF.headF">���� ����</h2>
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
                                - ������ ����. ���� ��쿡�� �Է��� �ּ���.
                            </td>
                        </tr>
                        <tr v-show="showF.detailCaseF">
                            <th>HTML</th>
                            <td>
                                <textarea name="htmlcode" style="width: 100%;" rows="5" v-model="current_content.htmlcode"></textarea>
                                - �ѱ� ���� �ִ� 800�� ���� �Է� �����մϴ�.
                            </td>
                        </tr>    
                        <tr v-show="showF.videoUrlF">
                            <th>������ �ڵ� *</th>
                            <td>
                                <input type="text" name="videourl" v-model="current_content.videourl" :class="{must : checkMustClass('videourl')}" />
                                - Youtube URL�� �Է����ּ���.
                            </td>
                        </tr>                     
                    </tbody>
                </table>
                
                <br/> 
                <h2>�Խ� ����</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>������ *</th>
                            <td>
                                <input type="radio" name="openingflag" value="0" checked/>����
                                <input type="radio" name="openingflag" value="1" />������ 1
                                <input type="radio" name="openingflag" value="2" />������ 2
                                <input type="radio" name="openingflag" value="3" />������ 3
                            </td>
                        </tr>
                        <tr>
                            <th>�Խ��� *</th>
                            <td>
                                <input type="checkbox" name="isaod" v-model="current_content.isaod" true-value="1" fasle-value="0" value="1" />�����Ϻ��� ��� ����
                                <input type="text" name="startdate" id="start_date" v-model="current_content.startdate" class="must"/> ���� <input type="text" name="enddate" id="end_date" v-model="current_content.enddate" class="must"/> ����
                            </td>
                        </tr>                        
                    </tbody>
                </table>
                
                <br/> 
                <h2>� ����</h2>
                <table class="table table-write table-dark">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    <tbody>
                        <tr>
                            <th>�޸�</th>
                            <td>
                                <textarea name="ordertext" style="width: 100%;" rows="5" v-model="current_content.ordertext"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>�����Ȳ *</th>
                            <td>
                                <select name="stateflag" class="form-control inline small must" v-model="current_content.stateflag">
                                    <option value="0">����</option>
                                    <option value="1">��ϴ��</option>
                                    <option value="2">�����ο�û</option>
                                    <option value="3">�ۺ��̿�û</option>
                                    <option value="4">���߿�û</option>
                                    <option value="5">���¿�û</option>
                                    <option value="7">����</option>
                                    <option value="8">����</option>
                                    <option value="9">����</option>
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

        const arrDayMin = ["��","��","ȭ","��","��","��","��"];
        const arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
        $("#start_date").datepicker({
            dateFormat: "yy-mm-dd",
            prevText: '������', nextText: '������', yearSuffix: '��',
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
            prevText: '������', nextText: '������', yearSuffix: '��',
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
            pidx : {type:Number, default:0} // ������ idx
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

            }else if(uiNumber == 1){ //����Ʈ��
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

            if(uinumber == 1){ //����Ʈ��
                if(name == 'tagP' || name == 'itemid'){
                    return true;
                }else return false
            }else if(uinumber == 2){
                if(name == 'contents') return true;
                else return false;
            }else if(uinumber == 4){ // ��������
                if(name == 'videourl' || name == 'contents'){
                    return true;
                }else return false;
            }else if(uinumber == 5){ // �̺�Ʈ��
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
            if(confirm("�����Ͻðڽ��ϱ�?")){
                this.current_content.listimage = "";
                $("#addListimage").val("");
            }
        }
        , reg_itemid(){
            window.open("/admin/sitemaster/piece/pop_singleItemSelect_V2.asp?target=play_content&ptype=piece&itemarr=", "������ǰ", "width:300px, height:200px");
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
                alert("800�ڸ� �ʰ��߽��ϴ�.");
            }
        }
    }
    , watch:{
        pop_content(popcontent) { // ������ ���� ��
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