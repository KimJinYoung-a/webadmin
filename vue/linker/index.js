const app = new Vue({
    el : '#app',
    mixin : [api_mixin],
    store : store,
    template : `
        <div class="container">
            
            <header class="linker-title">
                <h3>LINKER ����</h3>
                <div class="nickname-btn-area">
                    <a @click="$refs.nicknameDictionaryModal.openModal()">�г��� ����</a>
                    <a @click="$refs.manageNicknameSlangModal.openModal()">�г��� ��Ӿ� ����</a>
                </div>
            </header>
            
            <div class="linker-content">
                
                <!-- region ���� ����Ʈ -->
                <div class="forum-list-container">
                    <div class="title">
                        <h5>����</h5>
                        <div class="btn-area">
                            <button @click="$refs.postForumModal.openModal()" class="linker-btn">���� ���</button>
                            <button @click="$refs.manageForumSortModal.openModal()" class="linker-btn">���� ��������</button>
                        </div>
                    </div>
                    
                    <ul class="forum-list">
                        <LIST-FORUM v-for="forum in forums" :key="forum.forumIdx" :forum="forum" :currentForumIdx="currentForumIdx"
                            @clickForum="changeCurrentForumIdx"/>
                    </ul>
                    
                </div>
                <!-- endregion -->
                
                <div v-if="currentForum" class="forum-content">
                    <!-- region ��� ���� ���� -->
                    <div class="forum-content-top">
                        <div class="title-control">
                            <div class="title">
                                <p v-html="currentForum.subTitle"></p>
                                <h5 v-html="currentForum.title"></h5>
                            </div>
                            <div class="btn-area">
                                <button @click="openModifyForumModal(currentForum)" class="linker-btn">���� ����</button>
                                <button @click="deleteForum(currentForumIdx)" class="linker-btn">���� ����</button>
                            </div>
                        </div>
                        <div class="title-info">
                            <span>��Ⱓ : {{currentForumPeriod}}</span>
                            <span>����Ʈ ���¿��� : {{currentForum.useYn ? '����' : '���¾���'}}</span>
                            <span>���� ���� : {{currentForum.sortNo}}</span>
                        </div>
                    </div>
                    <!-- endregion -->
                    
                    <FORUM-INFO ref="forumInfo" :infos="currentForum.infos" 
                        @postForumInfo="openPostForumInfoModal" @deleteInfos="deleteForumInfos" 
                        @modifySort="modifyForumInfoSort"/>
                
                    <!--region ������ ����Ʈ-->
                    <div class="forum-content-bottom">
                        <div class="title">
                            <h3>������ ����Ʈ</h3>
                            <div>
                                <a @click="openManageReportPostingsModal">�Ű� ������ ����</a>
                                <a @click="openManageReportCommentsModal">�Ű� ��� ����</a>
                            </div>
                        </div>
                        
                        <!-- region �˻� -->
                        <div class="search">
                            <div>
                                <!--region ȸ������-->
                                <div class="search-group">
                                    <label>ȸ������:</label>
                                    <select v-model="postingSearch.creatorType">
                                        <option value="">��ü</option>
                                        <option value="H">Host</option>
                                        <option value="G">Guest</option>
                                        <option value="N">User</option>
                                    </select>
                                </div>
                                <!--endregion-->
                                <!--region ȸ�����-->
                                <div class="search-group">
                                    <label>ȸ�����:</label>
                                    <select v-model="postingSearch.creatorLevelName">
                                        <option value="">��ü</option>
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
                                <!--region �˻���-->
                                <div class="search-group">
                                    <select v-model="postingSearch.searchTypes">
                                        <option value="creator_nickname">�г���</option>
                                        <option value="creator_descr">ȸ������</option>
                                        <option value="posting_cotents">�ۼ�����</option>
                                    </select>
                                    :
                                    <input v-model="postingSearch.keyword" type="text">
                                </div>
                                <!--endregion-->
                                <!--region �������-->
                                <div class="search-group">
                                    <label>�������:</label>
                                    <DATE-PICKER @updateDate="setPostingSearchStartDate" :date="postingSearch.startDate" id="searchStartDate"/>
                                    ~
                                    <DATE-PICKER @updateDate="setPostingSearchEndDate" :date="postingSearch.endDate" id="searchEndDate"/>
                                </div>
                                <!--endregion-->
                            </div>
                            <button @click="searchPostings" class="linker-btn">�˻�</button>
                        </div>
                        <!-- endregion -->
                        
                        <!--region �˻����-->
                        <div class="forum-posting-result">
                            <div class="forum-posting-top">
                                <p>�˻���� : <span>{{numberFormat(postingCount)}}</span></p>
                                <div>
                                    <select v-model="pageSize" class="page-size">
                                        <option value="10">10��</option>
                                        <option value="20">20��</option>
                                        <option value="50">50��</option>
                                    </select>
                                    <button @click="openFixPostingModal" class="linker-btn">���� �׸� ����</button>
                                    <button @click="openManageFixPostingsModal" class="linker-btn">���� ������ ����</button>
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
            
            <!-- region ���� �űԵ��/���� ��� -->
            <MODAL ref="postForumModal" :title="postForumTitle" @closeModal="clearModifyForum">
                <POST-FORUM slot="body" :modifyForum="modifyForum" @saveForum="completeSaveForum"/>
            </MODAL>
            <!-- endregion -->
            
            <!-- region ���� ���� �������� ��� -->
            <MODAL ref="manageForumSortModal" title="���� ���� ��������" :width="1100">
                <MANAGE-FORUM-SORT slot="body" ref="manageForumSort" :forums="forums" 
                    @modifySortNo="modifyForumSortNo" @modifyForum="openModifyForumModal"
                    @deleteForum="deleteForum"/>
            </MODAL>
            <!-- endregion -->
            
            <!-- region ���� �ȳ� ���/���� ��� -->
            <MODAL ref="postForumInfoModal" title="���� �ȳ�" :width="750" @closeModal="modifyForumInfo = null">
                <POST-FORUM-INFO slot="body" ref="postForumInfo" :forumIdx="currentForumIdx" 
                    :modifyInfo="modifyForumInfo" @saveInfo="saveForumInfo"/>
            </MODAL>
            <!-- endregion -->
            
            <!--region ������ ���� ���-->
            <MODAL ref="managePostingModal" title="������ ����" :width="700">
                <MANAGE-POSTING slot="body" ref="managePosting" :posting="managedPosting" 
                    @savePosting="updatePosting" @openFixPostingModal="openManagedFixPostingModal"/>
            </MODAL>
            <!--endregion-->
            
            <!--region ������ ���� ���/���� ���-->
            <MODAL ref="postFixPostingModal" title="������ ����" @closeModal="callbackClosePostFixPostingModal">
                <MANAGE-FIX-POSTING slot="body" :orgFix="managedFixPosting" @saveFixPosting="saveFixPosting"/>
            </MODAL>
            <!--endregion-->
            
            <!--region �Ű� ������ ���� ���-->
            <MODAL ref="manageReportPostingsModal" title="�Ű� ������ ����" :width="1150">
                <MANAGE-REPORT-POSTINGS slot="body" ref="manageReportPostings" 
                    :postingCount="reportPostingCount" :postings="reportPostings"
                    @deletePosting="deletePosting" @deletePostings="deletePostings"
                    @unBlockPostings="unBlockPostings"/>
            </MODAL>
            <!--endregion-->
            
            <!--region �Ű� ��� ���� ���-->
            <MODAL ref="manageReportCommentsModal" title="�Ű� ��� ����" :width="1150">
                <MANAGE-REPORT-COMMENTS slot="body" ref="manageReportComments" :comments="reportComments"
                    @deleteComments="deleteComments" @unBlockComments="unBlockComments"/>
            </MODAL>
            <!--endregion-->
            
            <!--region ���� ������ ���� ���-->
            <MODAL ref="manageFixPostingsModal" title="���� ������ ����" :width="1100">
                <MANAGE-FIX-POSTINGS slot="body" ref="manageFixPostings" :postings="fixPostings"
                    @modifyPosting="openManagePostingModalInManageFixPosting"
                    @clearPostings="clearPostings" @modifyFixPositions="modifyFixPositions"/>
            </MODAL>
            <!--endregion-->
            
            <!--region �г��� ���� ���-->
            <MODAL ref="nicknameDictionaryModal" title="�г��� ����" :width="900">
                <MANAGE-NICKNAME-DICTIONARY slot="body" ref="nicknameDictionary" 
                    :words1="words1" :words2="words2" @modifyWord="openModifyNicknameModal"
                    @openPostModal="openPostNicknameModal" @deleteWords="deleteWords"/>
            </MODAL>
            <!--endregion-->
            
            <!--region �г��� ��Ӿ� ���� ���-->
            <MODAL ref="manageNicknameSlangModal" title="�г��� ��Ӿ� ����" :width="700">
                <MANAGE-NICKNAME-SLANG slot="body" ref="manageNicknameSlang" :words="slangWords" @deleteWords="deleteWords"
                    @openPostModal="openPostNicknameModal" @modifyWord="openModifyNicknameModal"/>
            </MODAL>
            <!--endregion-->
            
            <!--region �г��� ��� ���-->
            <MODAL ref="postNicknameModal" :title="postNicknameTitle">
                <POST-WORDS slot="body" :wordNumber="postNicknameNumber" @saveNicknames="postNicknames"/>
            </MODAL>
            <!--endregion-->
            
            <!--region �г��� ���� ���-->
            <MODAL ref="modifyNicknameModal" :title="'�ܾ�' + postNicknameNumber" @closeModal="lockBodyScroll">
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
        postForumTitle : '���� �űԵ��', // ���� ���/���� ��� Ÿ��Ʋ
        modifyForumInfo : null, // ������ ���� �ȳ�
        managedPosting : null, // ������ ������
        managedFixPosting : null, // ������ ���� ������
        checkedPostingIdx : 0, // ������ ������ �Ϸù�ȣ
        postNicknameNumber : 1, // �г��� ���/���� ��ȣ
        modifyWord : null, // ���� �� �г���

        page : 1, // ���� ������
        pageSize : 10, // �������� ����

        // region postingSearch ������ �˻�
        postingSearch : {
            creatorType : '', // ȸ������
            creatorLevelName : '', // ȸ������
            searchTypes : 'creator_nickname', // �˻�����
            keyword : '', // �˻���
            startDate : '', // �˻� ��������
            endDate : '', // �˻� ��������
        },
        // endregion
    }},
    computed : {
        forums() { return this.$store.getters.forums; }, // ���� ����Ʈ
        postingCount() { return this.$store.getters.postingCount; }, // ������ ��ü ����
        postings() { return this.$store.getters.postings; }, // ������ ����Ʈ
        lastPage() { return this.$store.getters.lastPage; }, // ������ ������
        reportPostingCount() { return this.$store.getters.reportPostingCount; }, // �Ű� ������ ����
        reportPostings() { return this.$store.getters.reportPostings; }, // �Ű� ������ ����Ʈ
        reportComments() { return this.$store.getters.reportComments; }, // �Ű� ��� ����Ʈ
        fixPostings() { return this.$store.getters.fixPostings; }, // ���� ������ ����Ʈ
        words1() { return this.$store.getters.words1; }, // �ܾ�1 ����Ʈ
        words2() { return this.$store.getters.words2; }, // �ܾ�2 ����Ʈ
        slangWords() { return this.$store.getters.slangWords; }, // �г��� ��Ӿ� ����Ʈ
        //region currentForum ���� Ȱ��ȭ�� ����
        currentForum() {
            return this.forums.find(f => this.currentForumIdx === f.forumIdx);
        },
        //endregion
        //region currentForumPeriod ���� ���� ���� �Ⱓ
        currentForumPeriod() {
            return this.getLocalDateTimeFormat(this.currentForum.startDate, 'yyyy-MM-dd')
                + ' ~ ' + this.getLocalDateTimeFormat(this.currentForum.endDate, 'yyyy-MM-dd');
        },
        //endregion
        //region getPostingsApiData ������ ��� ��ȸ API ȣ�� ������
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
        //region postNicknameTitle �г��� ��� ��� Ÿ��Ʋ
        postNicknameTitle() {
            if( this.postNicknameNumber === 3 )
                return '��Ӿ� ���';
            else
                return '�ܾ�' + this.postNicknameNumber + ' ���';
        },
        //endregion
    },
    methods : {
        //region changeCurrentForumIdx ���� Ȱ��ȭ ���� �Ϸù�ȣ ����
        changeCurrentForumIdx(forumIdx) {
            this.currentForumIdx = forumIdx;
        },
        //endregion
        //region completeSaveForum ���� ���� �Ϸ�
        completeSaveForum() {
            this.$refs.postForumModal.closeModal();
            this.$store.dispatch('GET_FORUMS', this);
        },
        //endregion
        //region deleteForum ���� ����
        deleteForum(forumIdx) {
            if( forumIdx && confirm('�ش� ������ ���� �����Ͻðڽ��ϱ�?') ) {
                this.callApi(2, 'POST', `/linker/forum/delete/${forumIdx}`, null, this.successDeleteForum);
            }
        },
        successDeleteForum() {
            this.$refs.manageForumSortModal.closeModal();
            this.$store.dispatch('GET_FORUMS', this);
            this.changeCurrentForumIdx(this.forums[0].forumIdx);
        },
        //endregion
        //region modifyForumSortNo ���� ���ļ��� ����
        modifyForumSortNo(sortedForumsBySortNo) {
            this.$store.commit('MODIFY_FORUMS_SORT', sortedForumsBySortNo);
            alert('���� �Ǿ����ϴ�.');
            this.$refs.manageForumSort.syncSortedForumsBySortNo();
        },
        //endregion
        //region openModifyForumModal ���� ���� ��� ����
        openModifyForumModal(forum) {
            this.modifyForum = forum;
            this.$refs.manageForumSortModal.closeModal();
            this.postForumTitle = '���� ����';
            this.modifyForum = forum;
            this.$refs.postForumModal.openModal();
        },
        //endregion
        //region clearModifyForum ���� ���� �ʱ�ȭ
        clearModifyForum() {
            this.postForumTitle = '���� �űԵ��';
            this.modifyForum = null;
        },
        //endregion
        //region openPostForumInfoModal ���� �ȳ� ���/���� ��� ����
        openPostForumInfoModal(info) {
            if( info )
                this.modifyForumInfo = info;
            this.$refs.postForumInfoModal.openModal();
        },
        //endregion
        //region setPostingSearchDate ������ �˻� ���� Set
        setPostingSearchStartDate(date) {
            this.postingSearch.startDate = date;
        },
        setPostingSearchEndDate(date) {
            this.postingSearch.endDate = date;
        },
        //endregion
        //region resetPostingSearch ������ �˻����� �ʱ�ȭ
        resetPostingSearch() {
            this.postingSearch = {
                creatorType : '', // ȸ������
                creatorLevelName : '', // ȸ������
                searchTypes : 'creator_nickname', // �˻�����
                keyword : '', // �˻���
                startDate : '', // �˻� ��������
                endDate : '', // �˻� ��������
            };
        },
        //endregion
        //region saveForumInfo ���� �ȳ� ����
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
        //region deleteForumInfos ���� �ȳ� ����
        deleteForumInfos(infoIdxs) {
            this.$store.dispatch('DELETE_FORUM_INFOS', {
                forumIdx : this.currentForumIdx,
                idxs : infoIdxs,
                app : this
            });
        },
        //endregion
        //region modifyForumInfoSort ���� �ȳ� ���� ����
        modifyForumInfoSort(idxs) {
            this.$store.dispatch('MODIFY_FORUM_INFO_SORT', {
                app : this,
                forumIdx : this.currentForumIdx,
                idxs : idxs
            });
        },
        //endregion
        //region changePostingPage ������ ������ ����
        changePostingPage(page) {
            this.page = page;
        },
        //endregion
        //region getPostings ������ ����Ʈ ��ȸ
        getPostings() {
            this.$store.dispatch('GET_POSTINGS', {
                app : this,
                apiData : this.getPostingsApiData
            });
        },
        //endregion
        //region searchPostings ������ �˻�
        searchPostings() {
            this.page = 1;
            this.getPostings();
        },
        //endregion
        //region openManagePostingModal ������ ���� ��� ����
        openManagePostingModal(posting) {
            this.managedPosting = posting;
            this.$refs.managePostingModal.openModal();
        },
        //endregion
        //region openManagePostingModalInManageFixPosting ���� ������ �������� ������ ���� ��� ����
        openManagePostingModalInManageFixPosting(posting) {
            this.$refs.manageFixPostingsModal.closeModal();
            this.openManagePostingModal(posting);
        },
        //endregion
        //region updatePosting ������ ����
        updatePosting(posting) {
            this.$store.dispatch('UPDATE_POSTING', {
                app : this,
                data : posting
            });
        },
        //endregion
        //region openManagedFixPostingModal ������ �������� ���������� ���/���� ��� ����
        openManagedFixPostingModal(fix) {
            this.managedFixPosting = fix;
            this.$refs.postFixPostingModal.openModal();
        },
        //endregion
        //region saveFixPosting ���� ������ ���/���� ��� ����
        saveFixPosting(fix) {
            if( this.managedPosting ) { // ������ ���� ��
                this.$refs.managePosting.setFixPosting(fix);
            } else { // ���� �׸� ����
                this.$store.dispatch('UPDATE_FIX_POSTING', {
                    app : this,
                    postingIdx : this.checkedPostingIdx,
                    fix
                });
            }
            this.$refs.postFixPostingModal.closeModal();
        },
        //endregion
        //region closeManageFixPostingModal ���� ������ ��� ���� �� callback
        callbackClosePostFixPostingModal() {
            this.managedFixPosting = null;
        },
        //endregion
        //region checkPosting ������ üũ
        checkPosting(e, postingIdx) {
            if( e.target.checked )
                this.checkedPostingIdx = postingIdx;
            else
                this.checkedPostingIdx = 0;
        },
        //endregion
        //region openFixPostingModal ���� ������ ��� ����
        openFixPostingModal() {
            if( this.checkedPostingIdx === 0 ) {
                alert('������ �������� ���� �� �ּ���');
                return false;
            }
            this.$refs.postFixPostingModal.openModal();
        },
        //endregion
        //region clearPostings ���� ������ ������ ����
        clearPostings(postingIdxs) {
            this.$store.dispatch('CLEAR_FIX_POSTINGS', {
                app : this,
                postingIdxs
            });
        },
        //endregion
        //region modifyFixPositions ���� ������ ����Ʈ ���� ��ġ ����
        modifyFixPositions(data) {
            this.$store.dispatch('MODIFY_FIX_POSTING_POSITIONS', {
                app : this,
                data
            });
        },
        //endregion
        //region deletePosting ������ ����
        deletePosting(postingIdx, isReportPosting) {
            if( confirm('�� �������� ���� �Ͻðڽ��ϱ�?') ) {
                this.$store.dispatch('DELETE_POSTING', {
                    app : this,
                    postingIdx,
                    isReportPosting : isReportPosting
                });
            }
        },
        //endregion
        //region deletePostings ������ ������ ����
        deletePostings(postingIdxs) {
            this.$store.dispatch('DELETE_POSTINGS', {
                app : this,
                postingIdxs
            });
        },
        //endregion
        //region openManageFixPostingsModal ���� ������ ���� ��� ����
        openManageFixPostingsModal() {
            this.$store.dispatch('GET_FIX_POSTINGS', {
                app : this,
                forumIdx : this.currentForumIdx
            });
            this.$refs.manageFixPostingsModal.openModal();
        },
        //endregion
        //region openManageReportPostingsModal �Ű� ������ ���� ��� ����
        openManageReportPostingsModal() {
            this.$store.dispatch('GET_REPORT_POSTINGS', {
                app : this,
                forumIdx : this.currentForumIdx
            });
            this.$refs.manageReportPostingsModal.openModal();
        },
        //endregion
        //region unBlockPostings ������ ��� ����
        unBlockPostings(postingIdxs) {
            this.$store.dispatch('UNBLOCK_POSTINGS', {
                app : this, postingIdxs
            });
        },
        //endregion
        //region openManageReportCommentsModal �Ű� ��� ���� ��� ����
        openManageReportCommentsModal() {
            this.$store.dispatch('GET_REPORT_COMMENTS', {
                app : this,
                forumIdx : this.currentForumIdx
            });
            this.$refs.manageReportCommentsModal.openModal();
        },
        //endregion
        //region deleteComments �Ű� ��� ������ ����
        deleteComments(commentIdxs) {
            this.$store.dispatch('DELETE_REPORT_COMMENTS', {
                app : this,
                commentIdxs
            });
        },
        //endregion
        //region unBlockComments �Ű� ��� ������ ��� ����
        unBlockComments(commentIdxs) {
            this.$store.dispatch('CLEAR_BLOCK_REPORT_COMMENTS', {
                app : this,
                commentIdxs
            });
        },
        //endregion
        //region openPostNicknameModal �г��� ��� ��� ����
        openPostNicknameModal(num) {
            if( this.$refs.nicknameDictionaryModal.show )
                this.$refs.nicknameDictionaryModal.closeModal();
            else if( this.$refs.manageNicknameSlangModal.show )
                this.$refs.manageNicknameSlangModal.closeModal();

            this.postNicknameNumber = num;
            this.$refs.postNicknameModal.openModal();
        },
        //endregion
        //region postNicknames �г��� ���
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
        //region deleteWords �ܾ� ������ ����
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
        //region openModifyNicknameModal �ܾ� ���� ��� ����
        openModifyNicknameModal(num, word) {
            this.postNicknameNumber = num;
            this.modifyWord = word;
            this.$refs.modifyNicknameModal.openModal();
        },
        //endregion
        //region modifyNickname �ܾ� ����
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
        //region getNicknameTypeStr �г��� ���� string ��
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
        //region numberFormat ���� õ�ڸ� (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
        //region lockBodyScroll Body ��ũ�� ���
        lockBodyScroll() {
            document.body.style.overflow = 'hidden';
            document.body.style.height = '100%';
        },
        //endregion
    },
    watch : {
        //region currentForumIdx ���� ������ȣ ����
        currentForumIdx() {
            this.page = 1;
            this.resetPostingSearch();
            if( this.$refs.forumInfo )
                this.$refs.forumInfo.setTempInfos(this.currentForum.infos);

            this.getPostings();
        },
        //endregion
        //region pageSize ������ ������ ����
        pageSize() {
            this.getPostings();
        },
        //endregion
        //region page ������ ������ ����
        page() {
            this.getPostings();
        },
        //endregion
    }
});