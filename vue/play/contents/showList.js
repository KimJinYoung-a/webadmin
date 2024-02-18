Vue.component('Show-List',{
    template: `
        <input type="button" @click="change_show_type('list')" value="리스트"/>
        <input type="button" @click="change_show_type('contents')" value="컨텐츠"/>
        <!-- 검색 테이블 -->
        <table class="table table-dark table-search">
            <thead class="thead-tenbyten">
                <tr>
                    <th>기간</th>
                    <td style="text-align:left;display: flex;" colspan="4">
                        <select id="period" class="form-control inline small">
                            <option value="1">시작일 기준</option>
                            <option value="2">종료일 기준</option>
                            <option value="3">등록일 기준</option>
                        </select>
                        
                        <input id="startDate" class="text" size="10" maxlength="10" />
                         ~
                        <input id="endDate" class="text" size="10" maxlength="10" />
                    </td>
                </tr>
                <tr>
                    <th>레이아웃</th>
                    <td style="text-align:left;">
                        <select id="uiNumber" class="form-control inline small">
                            <option value="0">전체</option>
                            <option value="1">리스트형</option>
                            <option value="2">상세형</option>
                            <option value="3">동영상형</option>
                            <option value="4">이벤트형</option>
                        </select>
                    </td>
                    
                    <th>컨텐츠</th>
                    <td style="text-align:left;">
                        <select id="contentsNumber" class="form-control inline small">
                            <option value="0">전체</option>
                            <option value="1">마스터피스</option>
                            <option value="2">탐구생활</option>
                            <option value="3">DAY.FILM</option>
                            <option value="4">THING.배지</option>
                            <option value="5">PLAY.GOODS</option>
                            <option value="7">WEEKLY WALLPAPER</option>
                        </select>
                    </td>
                    
                    <th>진행상태</th>
                    <td style="text-align:left;">
                        <select id="stateFlag" class="form-control inline small">
                            <option value="0" selected="selected">선택</option>
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
                <tr>
                    <th>키워드 검색</th>
                    <td style="text-align:left;">
                        <select id="searchKey" class="form-control inline small">
                            <option value="1">번호</option>
                            <option value="2">컨텐츠명</option>
                            <option value="3">작성자</option>
                        </select>
                    </td>
                  </tr>
                  <tr>
                    <td class="td-button align-right">
                        <button @click="reload" type="button" class="button secondary">검색조건Reset</button>
                        <button @click="do_search" type="button" class="button dark">검색</button>
                    </td>
                  </tr>
            </thead>
        </table>
        
        <p class="p-table">
            <span>검색결과 : <strong>{{content_count}}</strong></span>
            <i class='fas fa-sync' @click="reload"></i>
            <select @change="change_page_size" id="page_size" class="form-control form-control-sm">
                <option v-for="n in 5" :value="n*10">{{n*10}}개씩 보기</option>
            </select>
            <button id="reg_new_content" @click="popup_content('')" type="button" class="button dark">신규 등록</button>
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
                      <th>오프닝 1</th>
                      <th>오프닝 2</th>
                      <th>오프닝 3</th>
                  </tr>
                </thead>
                <tbody>
                    <tr>
                        <td v-for="item in opening_list">
                            <div v-if="item.pidx">
                                {{item.pidx}} {{item.contentTitleName}} <br/>
                              <b>{{item.titlename}}</b>
                              <p>{{item.startdate}} ~ {{item.enddate}} 오픈 </p>
                              <input type="button" @click="deleteOpening(item.pidx)" value="제거"/>     
                            </div>                                                           
                        </td>
                    </tr>
                </tbody>
            </table>
        </p>

        <!-- 리스트 테이블 -->
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
                    <th>썸네일</th>
                    <th>상세</th>
                    <th>조회수</th>
                    <th>오픈 여부</th>
                    <th>기타</th>
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
                        <p>{{content.occupation}} {{content.nickname}} {{content.regdate}} 등록</p>
                        <p v-if="content.lastupdate">{{content.lastOccupation}} {{content.lastNickname}} {{content.lastupdate}} 최종수정</p> 
                        <p>{{content.startdate}} ~ {{content.enddate}} 오픈</p>
                    </td>
                    <td>{{content.viewcount}}</td>
                    <td>{{content.stateflag == 7 ? '오픈' : '미오픈'}}</td>
                    <td><input type="button" @click="popup_content(content.pidx)" value="수정"/> <input type="button" @click="delete_content(content.pidx)" value="보류"/></td>
                </tr>
            </tbody>
        </table>

        <!-- 페이지 -->
        <Pagination
            @click_page="click_page"
            :current_page="current_page"
            :last_page="last_page"
        ></Pagination>

        <!-- 등록/수정 모달 -->
        <Modal v-show="show_write_modal"
            @save="save_content" @close="show_write_modal = false"
            modal_width="830px" header_title="PLAY 컨텐츠 등록/수정"
        >
            <Content-Write slot="body" :pop_content="pop_content" :pop_content_items="pop_content_items" :pop_content_tag="pop_content_tag" ref="write"/>
        </Modal>

        <!-- 썸네일 모달 -->
        <Modal v-show="show_thumbnail_modal" @close="show_thumbnail_modal = false"
            modal_width="400px" :show_header_yn="false" :show_footer_yn="false"
            :close_background_click_yn="true"
        >
            <img width="100%" :src="popup_thumbnail_src" slot="body" />
        </Modal>
    `
    ,
});