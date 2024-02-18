const app = new Vue({
    el : '#app',
    mixin : [api_mixin],
    store : store,
    template : `
        <div class="container">
            
            <header class="linker-title">
                <h3>LINKER 관리</h3>
                <div class="nickname-btn-area">
                    <a @click="$refs.nicknameDictionaryModal.openModal()">닉네임 사전</a>
                    <a @click="$refs.manageNicknameSlangModal.openModal()">닉네임 비속어 관리</a>
                </div>
            </header>
            
            <div class="linker-content">
                
                <!-- region 포럼 리스트 -->
                <div class="forum-list-container">
                    <div class="title">
                        <h5>포럼</h5>
                        <div class="btn-area">
                            <button @click="$refs.postForumModal.openModal()" class="linker-btn">포럼 등록</button>
                            <button @click="$refs.manageForumSortModal.openModal()" class="linker-btn">노출 순서관리</button>
                        </div>
                    </div>
                    
                    <ul class="forum-list">
                        <LIST-FORUM v-for="forum in forums" :key="forum.forumIdx" :forum="forum" :currentForumIdx="currentForumIdx"
                            @clickForum="changeCurrentForumIdx"/>
                    </ul>
                    
                </div>
                <!-- endregion -->
                
                <div v-if="currentForum" class="forum-content">
                    <!-- region 상단 포럼 정보 -->
                    <div class="forum-content-top">
                        <div class="title-control">
                            <div class="title">
                                <p v-html="currentForum.subTitle"></p>
                                <h5 v-html="currentForum.title"></h5>
                            </div>
                            <div class="btn-area">
                                <button @click="openModifyForumModal(currentForum)" class="linker-btn">포럼 수정</button>
                                <button @click="deleteForum(currentForumIdx)" class="linker-btn">포럼 삭제</button>
                            </div>
                        </div>
                        <div class="title-info">
                            <span>운영기간 : {{currentForumPeriod}}</span>
                            <span>프론트 오픈여부 : {{currentForum.useYn ? '오픈' : '오픈안함'}}</span>
                            <span>노출 순서 : {{currentForum.sortNo}}</span>
                        </div>
                    </div>
                    <!-- endregion -->
                    
                    <FORUM-INFO ref="forumInfo" :infos="currentForum.infos" 
                        @postForumInfo="openPostForumInfoModal" @deleteInfos="deleteForumInfos" 
                        @modifySort="modifyForumInfoSort"/>
                
                    <!--region 포스팅 리스트-->
                    <div class="forum-content-bottom">
                        <div class="title">
                            <h3>포스팅 리스트</h3>
                            <div>
                                <a @click="openManageReportPostingsModal">신고 포스팅 관리</a>
                                <a @click="openManageReportCommentsModal">신고 댓글 관리</a>
                            </div>
                        </div>
                        
                        <!-- region 검색 -->
                        <div class="search">
                            <div>
                                <!--region 회원구분-->
                                <div class="search-group">
                                    <label>회원구분:</label>
                                    <select v-model="postingSearch.creatorType">
                                        <option value="">전체</option>
                                        <option value="H">Host</option>
                                        <option value="G">Guest</option>
                                        <option value="N">User</option>
                                    </select>
                                </div>
                                <!--endregion-->
                                <!--region 회원등급-->
                                <div class="search-group">
                                    <label>회원등급:</label>
                                    <select v-model="postingSearch.creatorLevelName">
                                        <option value="">전체</option>
                                        <option>WHITE</option>
                                        <option>RED</option>
                                        <option>VIP</option>
                                        <option>VIP GOLD</option>
                                        <option>VVIP</option>
                                        <option>STAFF</option>
                                        <option>BIZ</option>
                                    </select>
                                </div>
                                <!--endregion-->
                                <!--region 검색어-->
                                <div class="search-group">
                                    <select v-model="postingSearch.searchTypes">
                                        <option value="creator_nickname">닉네임</option>
                                        <option value="creator_descr">회원설명</option>
                                        <option value="posting_cotents">작성내용</option>
                                    </select>
                                    :
                                    <input v-model="postingSearch.keyword" type="text">
                                </div>
                                <!--endregion-->
                                <!--region 등록일자-->
                                <div class="search-group">
                                    <label>등록일자:</label>
                                    <DATE-PICKER @updateDate="setPostingSearchStartDate" :date="postingSearch.startDate" id="searchStartDate"/>
                                    ~
                                    <DATE-PICKER @updateDate="setPostingSearchEndDate" :date="postingSearch.endDate" id="searchEndDate"/>
                                </div>
                                <!--endregion-->
                            </div>
                            <button @click="searchPostings" class="linker-btn">검색</button>
                        </div>
                        <!-- endregion -->
                        
                        <!--region 검색결과-->
                        <div class="forum-posting-result">
                            <div class="forum-posting-top">
                                <p>검색결과 : <span>{{numberFormat(postingCount)}}</span></p>
                                <div>
                                    <select v-model="pageSize" class="page-size">
                                        <option value="10">10개</option>
                                        <option value="20">20개</option>
                                        <option value="50">50개</option>
                                    </select>
                                    <button @click="openFixPostingModal" class="linker-btn">선택 항목 고정</button>
                                    <button @click="openManageFixPostingsModal" class="linker-btn">고정 포스팅 관리</button>
                                </div>
                            </div>
                            
                            <POSTING-LIST :postings="postings" :checkedPostingIdx="checkedPostingIdx" 
                                @clickPosting="openManagePostingModal" @checkPosting="checkPosting"
                                @deletePosting="deletePosting"/>
        
                            <PAGINATION :currentPage="page" :lastPage="lastPage" @clickPage="changePostingPage"/>
                        </div>
                        <!--endregion-->
                    </div>
                    <!--endregion-->
                </div>
            </div>
            
            <!-- region 포럼 신규등록/수정 모달 -->
            <MODAL ref="postForumModal" :title="postForumTitle" @closeModal="clearModifyForum">
                <POST-FORUM slot="body" :modifyForum="modifyForum" @saveForum="completeSaveForum"/>
            </MODAL>
            <!-- endregion -->
            
            <!-- region 포럼 노출 순서관리 모달 -->
            <MODAL ref="manageForumSortModal" title="포럼 노출 순서관리" :width="1100">
                <MANAGE-FORUM-SORT slot="body" ref="manageForumSort" :forums="forums" 
                    @modifySortNo="modifyForumSortNo" @modifyForum="openModifyForumModal"
                    @deleteForum="deleteForum"/>
            </MODAL>
            <!-- endregion -->
            
            <!-- region 포럼 안내 등록/수정 모달 -->
            <MODAL ref="postForumInfoModal" title="포럼 안내" :width="750" @closeModal="modifyForumInfo = null">
                <POST-FORUM-INFO slot="body" ref="postForumInfo" :forumIdx="currentForumIdx" 
                    :modifyInfo="modifyForumInfo" @saveInfo="saveForumInfo"/>
            </MODAL>
            <!-- endregion -->
            
            <!--region 포스팅 관리 모달-->
            <MODAL ref="managePostingModal" title="포스팅 관리" :width="700">
                <MANAGE-POSTING slot="body" ref="managePosting" :posting="managedPosting" 
                    @savePosting="updatePosting" @openFixPostingModal="openManagedFixPostingModal"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 포스팅 고정 등록/수정 모달-->
            <MODAL ref="postFixPostingModal" title="포스팅 고정" @closeModal="callbackClosePostFixPostingModal">
                <MANAGE-FIX-POSTING slot="body" :orgFix="managedFixPosting" @saveFixPosting="saveFixPosting"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 신고 포스팅 관리 모달-->
            <MODAL ref="manageReportPostingsModal" title="신고 포스팅 관리" :width="1150">
                <MANAGE-REPORT-POSTINGS slot="body" ref="manageReportPostings" 
                    :postingCount="reportPostingCount" :postings="reportPostings"
                    @deletePosting="deletePosting" @deletePostings="deletePostings"
                    @unBlockPostings="unBlockPostings"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 신고 댓글 관리 모달-->
            <MODAL ref="manageReportCommentsModal" title="신고 댓글 관리" :width="1150">
                <MANAGE-REPORT-COMMENTS slot="body" ref="manageReportComments" :comments="reportComments"
                    @deleteComments="deleteComments" @unBlockComments="unBlockComments"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 고정 포스팅 관리 모달-->
            <MODAL ref="manageFixPostingsModal" title="고정 포스팅 관리" :width="1100">
                <MANAGE-FIX-POSTINGS slot="body" ref="manageFixPostings" :postings="fixPostings"
                    @modifyPosting="openManagePostingModalInManageFixPosting"
                    @clearPostings="clearPostings" @modifyFixPositions="modifyFixPositions"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 닉네임 사전 모달-->
            <MODAL ref="nicknameDictionaryModal" title="닉네임 사전" :width="900">
                <MANAGE-NICKNAME-DICTIONARY slot="body" ref="nicknameDictionary" 
                    :words1="words1" :words2="words2" @modifyWord="openModifyNicknameModal"
                    @openPostModal="openPostNicknameModal" @deleteWords="deleteWords"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 닉네임 비속어 관리 모달-->
            <MODAL ref="manageNicknameSlangModal" title="닉네임 비속어 관리" :width="700">
                <MANAGE-NICKNAME-SLANG slot="body" ref="manageNicknameSlang" :words="slangWords" @deleteWords="deleteWords"
                    @openPostModal="openPostNicknameModal" @modifyWord="openModifyNicknameModal"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 닉네임 등록 모달-->
            <MODAL ref="postNicknameModal" :title="postNicknameTitle">
                <POST-WORDS slot="body" :wordNumber="postNicknameNumber" @saveNicknames="postNicknames"/>
            </MODAL>
            <!--endregion-->
            
            <!--region 닉네임 수정 모달-->
            <MODAL ref="modifyNicknameModal" :title="'단어' + postNicknameNumber" @closeModal="lockBodyScroll">
                <MODIFY-WORD slot="body" :wordNumber="postNicknameNumber" :word="modifyWord" @save="modifyNickname"/>
            </MODAL>
            <!--endregion-->
        </div>
    `,
    created() {
        this.$store.dispatch('GET_FORUMS', this);
        this.$store.dispatch('GET_NICKNAMES', this);
        this.$store.dispatch('GET_NICKNAMES_SLANG', this);
    },
    data() {return {
        currentForumIdx : 0,
        modifyForum : null,
        postForumTitle : '포럼 신규등록', // 포럼 등록/수정 모달 타이틀
        modifyForumInfo : null, // 수정중 포럼 안내
        managedPosting : null, // 수정중 포스팅
        managedFixPosting : null, // 수정중 고정 포스팅
        checkedPostingIdx : 0, // 선택한 포스팅 일련번호
        postNicknameNumber : 1, // 닉네임 등록/수정 번호
        modifyWord : null, // 수정 할 닉네임

        page : 1, // 현재 페이지
        pageSize : 10, // 페이지별 갯수

        // region postingSearch 포스팅 검색
        postingSearch : {
            creatorType : '', // 회원구분
            creatorLevelName : '', // 회원구분
            searchTypes : 'creator_nickname', // 검색조건
            keyword : '', // 검색어
            startDate : '', // 검색 시작일자
            endDate : '', // 검색 종료일자
        },
        // endregion
    }},
    computed : {
        forums() { return this.$store.getters.forums; }, // 포럼 리스트
        postingCount() { return this.$store.getters.postingCount; }, // 포스팅 전체 갯수
        postings() { return this.$store.getters.postings; }, // 포스팅 리스트
        lastPage() { return this.$store.getters.lastPage; }, // 마지막 페이지
        reportPostingCount() { return this.$store.getters.reportPostingCount; }, // 신고 포스팅 갯수
        reportPostings() { return this.$store.getters.reportPostings; }, // 신고 포스팅 리스트
        reportComments() { return this.$store.getters.reportComments; }, // 신고 댓글 리스트
        fixPostings() { return this.$store.getters.fixPostings; }, // 고정 포스팅 리스트
        words1() { return this.$store.getters.words1; }, // 단어1 리스트
        words2() { return this.$store.getters.words2; }, // 단어2 리스트
        slangWords() { return this.$store.getters.slangWords; }, // 닉네임 비속어 리스트
        //region currentForum 현재 활성화된 포럼
        currentForum() {
            return this.forums.find(f => this.currentForumIdx === f.forumIdx);
        },
        //endregion
        //region currentForumPeriod 현재 포럼 오픈 기간
        currentForumPeriod() {
            return this.getLocalDateTimeFormat(this.currentForum.startDate, 'yyyy-MM-dd')
                + ' ~ ' + this.getLocalDateTimeFormat(this.currentForum.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region getPostingsApiData 포스팅 목록 조회 API 호출 데이터
        getPostingsApiData() {
            return {
                forumIdx : this.currentForumIdx,
                currentPage : this.page,
                pageSize : this.pageSize,
                creatorType : this.postingSearch.creatorType,
                creatorLevelName : this.postingSearch.creatorLevelName,
                searchTypes : this.postingSearch.searchTypes,
                keyword : this.postingSearch.keyword,
                searchStartDate : this.postingSearch.startDate,
                searchEndDate : this.postingSearch.endDate
            }
        },
        //endregion
        //region postNicknameTitle 닉네임 등록 모달 타이틀
        postNicknameTitle() {
            if( this.postNicknameNumber === 3 )
                return '비속어 등록';
            else
                return '단어' + this.postNicknameNumber + ' 등록';
        },
        //endregion
    },
    methods : {
        //region changeCurrentForumIdx 현재 활성화 포럼 일련번호 변경
        changeCurrentForumIdx(forumIdx) {
            this.currentForumIdx = forumIdx;
        },
        //endregion
        //region completeSaveForum 포럼 저장 완료
        completeSaveForum() {
            this.$refs.postForumModal.closeModal();
            this.$store.dispatch('GET_FORUMS', this);
        },
        //endregion
        //region deleteForum 포럼 삭제
        deleteForum(forumIdx) {
            if( forumIdx && confirm('해당 포럼을 정말 삭제하시겠습니까?') ) {
                this.callApi(2, 'POST', `/linker/forum/delete/${forumIdx}`, null, this.successDeleteForum);
            }
        },
        successDeleteForum() {
            this.$refs.manageForumSortModal.closeModal();
            this.$store.dispatch('GET_FORUMS', this);
            this.changeCurrentForumIdx(this.forums[0].forumIdx);
        },
        //endregion
        //region modifyForumSortNo 포럼 정렬순서 수정
        modifyForumSortNo(sortedForumsBySortNo) {
            this.$store.commit('MODIFY_FORUMS_SORT', sortedForumsBySortNo);
            alert('수정 되었습니다.');
            this.$refs.manageForumSort.syncSortedForumsBySortNo();
        },
        //endregion
        //region openModifyForumModal 수정 포럼 모달 열기
        openModifyForumModal(forum) {
            this.modifyForum = forum;
            this.$refs.manageForumSortModal.closeModal();
            this.postForumTitle = '포럼 수정';
            this.modifyForum = forum;
            this.$refs.postForumModal.openModal();
        },
        //endregion
        //region clearModifyForum 수정 포럼 초기화
        clearModifyForum() {
            this.postForumTitle = '포럼 신규등록';
            this.modifyForum = null;
        },
        //endregion
        //region openPostForumInfoModal 포럼 안내 등록/수정 모달 열기
        openPostForumInfoModal(info) {
            if( info )
                this.modifyForumInfo = info;
            this.$refs.postForumInfoModal.openModal();
        },
        //endregion
        //region setPostingSearchDate 포스팅 검색 일자 Set
        setPostingSearchStartDate(date) {
            this.postingSearch.startDate = date;
        },
        setPostingSearchEndDate(date) {
            this.postingSearch.endDate = date;
        },
        //endregion
        //region resetPostingSearch 포스팅 검색조건 초기화
        resetPostingSearch() {
            this.postingSearch = {
                creatorType : '', // 회원구분
                creatorLevelName : '', // 회원구분
                searchTypes : 'creator_nickname', // 검색조건
                keyword : '', // 검색어
                startDate : '', // 검색 시작일자
                endDate : '', // 검색 종료일자
            };
        },
        //endregion
        //region saveForumInfo 포럼 안내 저장
        saveForumInfo(info, modifyInfoIdx) {
            this.$refs.postForumInfoModal.closeModal();
            if( modifyInfoIdx ) {
                this.$store.commit('MODIFY_FORUM_INFO', {
                    forumIdx : this.currentForumIdx,
                    modifyInfoIdx,
                    info
                });
            } else {
                this.$store.commit('ADD_FORUM_INFO', {
                    forumIdx : this.currentForumIdx,
                    info : info
                });
            }
            this.$refs.forumInfo.setTempInfos(this.currentForum.infos);
        },
        //endregion
        //region deleteForumInfos 포럼 안내 삭제
        deleteForumInfos(infoIdxs) {
            this.$store.dispatch('DELETE_FORUM_INFOS', {
                forumIdx : this.currentForumIdx,
                idxs : infoIdxs,
                app : this
            });
        },
        //endregion
        //region modifyForumInfoSort 포럼 안내 정렬 수정
        modifyForumInfoSort(idxs) {
            this.$store.dispatch('MODIFY_FORUM_INFO_SORT', {
                app : this,
                forumIdx : this.currentForumIdx,
                idxs : idxs
            });
        },
        //endregion
        //region changePostingPage 포스팅 페이지 변경
        changePostingPage(page) {
            this.page = page;
        },
        //endregion
        //region getPostings 포스팅 리스트 조회
        getPostings() {
            this.$store.dispatch('GET_POSTINGS', {
                app : this,
                apiData : this.getPostingsApiData
            });
        },
        //endregion
        //region searchPostings 포스팅 검색
        searchPostings() {
            this.page = 1;
            this.getPostings();
        },
        //endregion
        //region openManagePostingModal 포스팅 관리 모달 열기
        openManagePostingModal(posting) {
            this.managedPosting = posting;
            this.$refs.managePostingModal.openModal();
        },
        //endregion
        //region openManagePostingModalInManageFixPosting 고정 포스팅 관리에서 포스팅 수정 모달 열기
        openManagePostingModalInManageFixPosting(posting) {
            this.$refs.manageFixPostingsModal.closeModal();
            this.openManagePostingModal(posting);
        },
        //endregion
        //region updatePosting 포스팅 수정
        updatePosting(posting) {
            this.$store.dispatch('UPDATE_POSTING', {
                app : this,
                data : posting
            });
        },
        //endregion
        //region openManagedFixPostingModal 포스팅 관리에서 고정포스팅 등록/수정 모달 열기
        openManagedFixPostingModal(fix) {
            this.managedFixPosting = fix;
            this.$refs.postFixPostingModal.openModal();
        },
        //endregion
        //region saveFixPosting 고정 포스팅 등록/수정 모달 저장
        saveFixPosting(fix) {
            if( this.managedPosting ) { // 포스팅 수정 중
                this.$refs.managePosting.setFixPosting(fix);
            } else { // 선택 항목 고정
                this.$store.dispatch('UPDATE_FIX_POSTING', {
                    app : this,
                    postingIdx : this.checkedPostingIdx,
                    fix
                });
            }
            this.$refs.postFixPostingModal.closeModal();
        },
        //endregion
        //region closeManageFixPostingModal 고정 포스팅 모달 닫은 후 callback
        callbackClosePostFixPostingModal() {
            this.managedFixPosting = null;
        },
        //endregion
        //region checkPosting 포스팅 체크
        checkPosting(e, postingIdx) {
            if( e.target.checked )
                this.checkedPostingIdx = postingIdx;
            else
                this.checkedPostingIdx = 0;
        },
        //endregion
        //region openFixPostingModal 고정 포스팅 모달 열기
        openFixPostingModal() {
            if( this.checkedPostingIdx === 0 ) {
                alert('고정할 포스팅을 선택 해 주세요');
                return false;
            }
            this.$refs.postFixPostingModal.openModal();
        },
        //endregion
        //region clearPostings 고정 포스팅 여러개 해제
        clearPostings(postingIdxs) {
            this.$store.dispatch('CLEAR_FIX_POSTINGS', {
                app : this,
                postingIdxs
            });
        },
        //endregion
        //region modifyFixPositions 고정 포스팅 리스트 노출 위치 수정
        modifyFixPositions(data) {
            this.$store.dispatch('MODIFY_FIX_POSTING_POSITIONS', {
                app : this,
                data
            });
        },
        //endregion
        //region deletePosting 포스팅 삭제
        deletePosting(postingIdx, isReportPosting) {
            if( confirm('이 포스팅을 삭제 하시겠습니까?') ) {
                this.$store.dispatch('DELETE_POSTING', {
                    app : this,
                    postingIdx,
                    isReportPosting : isReportPosting
                });
            }
        },
        //endregion
        //region deletePostings 포스팅 여러개 삭제
        deletePostings(postingIdxs) {
            this.$store.dispatch('DELETE_POSTINGS', {
                app : this,
                postingIdxs
            });
        },
        //endregion
        //region openManageFixPostingsModal 고정 포스팅 관리 모달 열기
        openManageFixPostingsModal() {
            this.$store.dispatch('GET_FIX_POSTINGS', {
                app : this,
                forumIdx : this.currentForumIdx
            });
            this.$refs.manageFixPostingsModal.openModal();
        },
        //endregion
        //region openManageReportPostingsModal 신고 포스팅 관리 모달 열기
        openManageReportPostingsModal() {
            this.$store.dispatch('GET_REPORT_POSTINGS', {
                app : this,
                forumIdx : this.currentForumIdx
            });
            this.$refs.manageReportPostingsModal.openModal();
        },
        //endregion
        //region unBlockPostings 포스팅 블락 해제
        unBlockPostings(postingIdxs) {
            this.$store.dispatch('UNBLOCK_POSTINGS', {
                app : this, postingIdxs
            });
        },
        //endregion
        //region openManageReportCommentsModal 신고 댓글 관리 모달 열기
        openManageReportCommentsModal() {
            this.$store.dispatch('GET_REPORT_COMMENTS', {
                app : this,
                forumIdx : this.currentForumIdx
            });
            this.$refs.manageReportCommentsModal.openModal();
        },
        //endregion
        //region deleteComments 신고 댓글 여러개 삭제
        deleteComments(commentIdxs) {
            this.$store.dispatch('DELETE_REPORT_COMMENTS', {
                app : this,
                commentIdxs
            });
        },
        //endregion
        //region unBlockComments 신고 댓글 여러개 블락 해제
        unBlockComments(commentIdxs) {
            this.$store.dispatch('CLEAR_BLOCK_REPORT_COMMENTS', {
                app : this,
                commentIdxs
            });
        },
        //endregion
        //region openPostNicknameModal 닉네임 등록 모달 열기
        openPostNicknameModal(num) {
            if( this.$refs.nicknameDictionaryModal.show )
                this.$refs.nicknameDictionaryModal.closeModal();
            else if( this.$refs.manageNicknameSlangModal.show )
                this.$refs.manageNicknameSlangModal.closeModal();

            this.postNicknameNumber = num;
            this.$refs.postNicknameModal.openModal();
        },
        //endregion
        //region postNicknames 닉네임 등록
        postNicknames(words) {
            this.$store.dispatch('POST_NICKNAMES', {
                app : this,
                data : {
                    type : this.getNicknameTypeStr(this.postNicknameNumber),
                    words : words.join(',')
                }
            });
        },
        //endregion
        //region deleteWords 단어 여러개 삭제
        deleteWords(typeNum, wordIdxs) {
            this.$store.dispatch('DELETE_NICKNAMES', {
                app : this,
                data : {
                    type : this.getNicknameTypeStr(typeNum),
                    wordIndexs : wordIdxs.join(',')
                }
            });
        },
        //endregion
        //region openModifyNicknameModal 단어 수정 모달 열기
        openModifyNicknameModal(num, word) {
            this.postNicknameNumber = num;
            this.modifyWord = word;
            this.$refs.modifyNicknameModal.openModal();
        },
        //endregion
        //region modifyNickname 단어 수정
        modifyNickname(wordIdx, word) {
            this.$store.dispatch('MODIFY_NICKNAME', {
                app : this,
                data : {
                    type : this.getNicknameTypeStr(this.postNicknameNumber),
                    wordIndex : wordIdx,
                    word : word
                }
            });
        },
        //endregion
        //region getNicknameTypeStr 닉네임 유형 string 값
        getNicknameTypeStr(num) {
            if( num === 1 )
                return 'adj';
            else if( num === 2 )
                return 'noun';
            else if( num === 3 )
                return 'slang';
            else
                return '';
        },
        //endregion
        //region numberFormat 숫자 천자리 (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
        //region lockBodyScroll Body 스크롤 잠금
        lockBodyScroll() {
            document.body.style.overflow = 'hidden';
            document.body.style.height = '100%';
        },
        //endregion
    },
    watch : {
        //region currentForumIdx 현재 포럼번호 변경
        currentForumIdx() {
            this.page = 1;
            this.resetPostingSearch();
            if( this.$refs.forumInfo )
                this.$refs.forumInfo.setTempInfos(this.currentForum.infos);

            this.getPostings();
        },
        //endregion
        //region pageSize 페이지 사이즈 변경
        pageSize() {
            this.getPostings();
        },
        //endregion
        //region page 포스팅 페이지 변경
        page() {
            this.getPostings();
        },
        //endregion
    }
});