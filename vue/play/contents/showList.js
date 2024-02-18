Vue.component('Show-List',{
    template: `
        <input type="button" @click="change_show_type('list')" value="����Ʈ"/>
        <input type="button" @click="change_show_type('contents')" value="������"/>
        <!-- �˻� ���̺� -->
        <table class="table table-dark table-search">
            <thead class="thead-tenbyten">
                <tr>
                    <th>�Ⱓ</th>
                    <td style="text-align:left;display: flex;" colspan="4">
                        <select id="period" class="form-control inline small">
                            <option value="1">������ ����</option>
                            <option value="2">������ ����</option>
                            <option value="3">����� ����</option>
                        </select>
                        
                        <input id="startDate" class="text" size="10" maxlength="10" />
                         ~
                        <input id="endDate" class="text" size="10" maxlength="10" />
                    </td>
                </tr>
                <tr>
                    <th>���̾ƿ�</th>
                    <td style="text-align:left;">
                        <select id="uiNumber" class="form-control inline small">
                            <option value="0">��ü</option>
                            <option value="1">����Ʈ��</option>
                            <option value="2">����</option>
                            <option value="3">��������</option>
                            <option value="4">�̺�Ʈ��</option>
                        </select>
                    </td>
                    
                    <th>������</th>
                    <td style="text-align:left;">
                        <select id="contentsNumber" class="form-control inline small">
                            <option value="0">��ü</option>
                            <option value="1">�������ǽ�</option>
                            <option value="2">Ž����Ȱ</option>
                            <option value="3">DAY.FILM</option>
                            <option value="4">THING.����</option>
                            <option value="5">PLAY.GOODS</option>
                            <option value="7">WEEKLY WALLPAPER</option>
                        </select>
                    </td>
                    
                    <th>�������</th>
                    <td style="text-align:left;">
                        <select id="stateFlag" class="form-control inline small">
                            <option value="0" selected="selected">����</option>
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
                <tr>
                    <th>Ű���� �˻�</th>
                    <td style="text-align:left;">
                        <select id="searchKey" class="form-control inline small">
                            <option value="1">��ȣ</option>
                            <option value="2">��������</option>
                            <option value="3">�ۼ���</option>
                        </select>
                    </td>
                  </tr>
                  <tr>
                    <td class="td-button align-right">
                        <button @click="reload" type="button" class="button secondary">�˻�����Reset</button>
                        <button @click="do_search" type="button" class="button dark">�˻�</button>
                    </td>
                  </tr>
            </thead>
        </table>
        
        <p class="p-table">
            <span>�˻���� : <strong>{{content_count}}</strong></span>
            <i class='fas fa-sync' @click="reload"></i>
            <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                <option v-for="n in 5" :value="n*10">{{n*10}}���� ����</option>
            </select>
            <button id="reg_new_content" @click="popup_content('')" type="button" class="button dark">�ű� ���</button>
        </p>
        
        <p>
            <table class="table table-dark">
                <colgroup>
                    <col style="width:33%"/>
                    <col style="width:33%"/>
                    <col style="width:33%"/>
                </colgroup>
                <thead>
                  <tr>
                      <th>������ 1</th>
                      <th>������ 2</th>
                      <th>������ 3</th>
                  </tr>
                </thead>
                <tbody>
                    <tr>
                        <td v-for="item in opening_list">
                            <div v-if="item.pidx">
                                {{item.pidx}} {{item.contentTitleName}} <br/>
                              <b>{{item.titlename}}</b>
                              <p>{{item.startdate}} ~ {{item.enddate}} ���� </p>
                              <input type="button" @click="deleteOpening(item.pidx)" value="����"/>     
                            </div>                                                           
                        </td>
                    </tr>
                </tbody>
            </table>
        </p>

        <!-- ����Ʈ ���̺� -->
        <table class="table table-dark">
            <colgroup>
                <col style="width:20%"/>
                <col style="width:40%"/>
                <col style="width:10%"/>
                <col style="width:10%"/>
                <col style="width:20%"/>
            </colgroup>
            <thead>
                <tr>
                    <th>�����</th>
                    <th>��</th>
                    <th>��ȸ��</th>
                    <th>���� ����</th>
                    <th>��Ÿ</th>
                </tr>
            </thead>
            <tbody>
                <tr v-for="content in contents" :key="content.content_idx">
                    <td>
                        <p v-if="content.openingflag !== 0">{{content.openingflag}}</p>
                        <img @click="popup_thumbnail" :src="content.listimage" class="img-thumbnail link" style="width:50px;height:50px;" />
                    </td>
                    <td @click="popup_content(content.pidx)" class="link">
                        {{content.pidx}} {{content.contentTitleName}} <br/>
                        <b>{{content.titlename}}</b>
                        <p>{{content.occupation}} {{content.nickname}} {{content.regdate}} ���</p>
                        <p v-if="content.lastupdate">{{content.lastOccupation}} {{content.lastNickname}} {{content.lastupdate}} ��������</p> 
                        <p>{{content.startdate}} ~ {{content.enddate}} ����</p>
                    </td>
                    <td>{{content.viewcount}}</td>
                    <td>{{content.stateflag == 7 ? '����' : '�̿���'}}</td>
                    <td><input type="button" @click="popup_content(content.pidx)" value="����"/> <input type="button" @click="delete_content(content.pidx)" value="����"/></td>
                </tr>
            </tbody>
        </table>

        <!-- ������ -->
        <Pagination
            @click_page="click_page"
            :current_page="current_page"
            :last_page="last_page"
        ></Pagination>

        <!-- ���/���� ��� -->
        <Modal v-show="show_write_modal"
            @save="save_content" @close="show_write_modal = false"
            modal_width="830px" header_title="PLAY ������ ���/����"
        >
            <Content-Write slot="body" :pop_content="pop_content" :pop_content_items="pop_content_items" :pop_content_tag="pop_content_tag" ref="write"/>
        </Modal>

        <!-- ����� ��� -->
        <Modal v-show="show_thumbnail_modal" @close="show_thumbnail_modal = false"
            modal_width="400px" :show_header_yn="false" :show_footer_yn="false"
            :close_background_click_yn="true"
        >
            <img width="100%" :src="popup_thumbnail_src" slot="body" />
        </Modal>
    `
    ,
});