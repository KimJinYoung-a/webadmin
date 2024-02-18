const store = new Vuex.Store({
    state : {
        forums : [], // 포럼 리스트

        postingCount : 0, // 포스팅 갯수
        postings : [], // 포스팅 리스트
        lastPage : 0, // 마지막 페이지

        reportPostingCount : 0, // 신고 포스팅 갯수
        reportPostings : [], // 신고 포스팅 리스트
        reportComments : [], // 신고 댓글 리스트
        fixPostings : [], // 고정 포스팅 리스트

        words1 : [], // 닉네임 단어1 리스트
        words2 : [], // 닉네임 단어2 리스트
        slangWords : [], // 닉네임 비속어 리스트
    },

    getters : {
        forums(state) {return state.forums;},
        postingCount(state) {return state.postingCount;},
        postings(state) {return state.postings;},
        lastPage(state) {return state.lastPage;},
        reportPostingCount(state) {return state.reportPostingCount;},
        reportPostings(state) {return state.reportPostings;},
        reportComments(state) {return state.reportComments;},
        fixPostings(state) {return state.fixPostings;},
        words1(state) {return state.words1;},
        words2(state) {return state.words2;},
        slangWords(state) {return state.slangWords;},
    },

    mutations : {
        //region SET_FORUMS Set 포럼 리스트
        SET_FORUMS(state, forums) {
            state.forums = forums;
        },
        //endregion
        //region MODIFY_FORUMS_SORT 포럼 리스트 정렬순서 수정
        MODIFY_FORUMS_SORT(state, sortForums) {
            sortForums.forEach(s => {
                const forum = state.forums.find(f => f.forumIdx === s.forumIdx);
                forum.sortNo = s.sortNo;
            });
        },
        //endregion
        //region ADD_FORUM_INFO 포럼 안내 추가
        ADD_FORUM_INFO(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            forum.infos.push(payload.info);
        },
        //endregion
        //region MODIFY_FORUM_INFO 포럼 안내 수정
        MODIFY_FORUM_INFO(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            const info = forum.infos.find(i => i.infoIdx === payload.modifyInfoIdx);
            info.appTitle = payload.info.appTitle;
            info.mobileTitle = payload.info.mobileTitle;
            info.pcTitle = payload.info.pcTitle;
            info.appContent = payload.info.appContent;
            info.mobileContent = payload.info.mobileContent;
            info.pcContent = payload.info.pcContent;
        },
        //endregion
        //region DELETE_FORUM_INFOS 포럼 안내 삭제
        DELETE_FORUM_INFOS(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            forum.infos = forum.infos.filter(i => payload.idxs.indexOf(i.infoIdx) === -1);
        },
        //endregion
        //region MODIFY_FORUM_INFO_SORT 포럼 안내 정렬 수정
        MODIFY_FORUM_INFO_SORT(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            payload.idxs.forEach((infoIdx, index) => {
                const info = forum.infos.find(info => info.infoIdx === infoIdx);
                info.sortNo = index + 1;
            });
            forum.infos.sort((i1, i2) => i1.sortNo - i2.sortNo);
        },
        //endregion
        //region SET_POSTING_COUNT Set 포스팅 전체 갯수
        SET_POSTING_COUNT(state, count) {
            state.postingCount = count;
        },
        //endregion
        //region SET_POSTINGS Set 포스팅 리스트
        SET_POSTINGS(state, postings) {
            state.postings = postings;
        },
        //endregion
        //region SET_LAST_PAGE Set 마지막 페이지
        SET_LAST_PAGE(state, page) {
            state.lastPage = page;
        },
        //endregion
        //region MODIFY_POSTING 포스팅 수정
        MODIFY_POSTING(state, payload) {
            const posting = state.postings.find(p => p.postingIdx === payload.postingIndex);
            posting.postingContents = payload.postingCotents;
            posting.fixed = payload.useYn;
            posting.startDate = payload.startDate;
            posting.endDate = payload.endDate;
            posting.positionNo = payload.positionNo;
        },
        //endregion
        //region SET_REPORT_POSTINGS Set 신고 포스팅 리스트&갯수
        SET_REPORT_POSTINGS(state, payload) {
            state.reportPostingCount = payload.totalCount;
            state.reportPostings = payload.postings;
        },
        //endregion
        //region SET_REPORT_COMMENTS Set 신고 댓글 리스트
        SET_REPORT_COMMENTS(state, payload) {
            state.reportComments = payload;
        },
        //endregion
        //region DELETE_REPORT_POSTING 신고 포스팅 삭제
        DELETE_REPORT_POSTING(state, postingIdxs) {
            state.reportPostings = state.reportPostings
                                            .filter(posting => postingIdxs.indexOf(posting.postingIdx) === -1);
        },
        //endregion
        //region SET_FIX_POSTINGS Set 고정 포스팅 리스트
        SET_FIX_POSTINGS(state, payload) {
            state.fixPostings = payload;
        },
        //endregion
        //region DELETE_REPORT_COMMENTS 신고 댓글 삭제
        DELETE_REPORT_COMMENTS(state, payload) {
            state.reportComments = state.reportComments.filter(c => payload.indexOf(c.commentIdx) === -1);
        },
        //endregion
        //region SET_WORDS Set 닉네임 단어 리스트
        SET_WORDS(state, payload) {
            state.words1 = payload.words1;
            state.words2 = payload.words2;
        },
        //endregion
        //region DELETE_WORDS 닉네임 단어 삭제
        DELETE_WORDS(state, payload) {
            const wordIndexs = payload.wordIndexs.split(',').map(i => Number(i));
            if( payload.type === 'adj' ) {
                state.words1 = state.words1.filter(w => wordIndexs.indexOf(w.wordIdx) === -1);
            } else if( payload.type === 'noun' ) {
                state.words2 = state.words2.filter(w => wordIndexs.indexOf(w.wordIdx) === -1);
            } else {
                state.slangWords = state.slangWords.filter(w => wordIndexs.indexOf(w.wordIdx) === -1);
            }
        },
        //endregion
        //region MODIFY_WORD 닉네임 단어 수정
        MODIFY_WORD(state, payload) {
            let words;
            if( payload.type === 'slang' )
                words = state.slangWords;
            else if( payload.type === 'adj' )
                words = state.words1;
            else
                words = state.words2;

            const word = words.find(w => w.wordIdx === payload.wordIndex);
            word.word = payload.word;
        },
        //endregion
        //region DELETE_FIX_POSTINGS 고정 포스팅 리스트 여러개 삭제
        DELETE_FIX_POSTINGS(state, postingIdxs) {
            state.fixPostings = state.fixPostings.filter(p => postingIdxs.indexOf(p.postingIdx) === -1);
        },
        //endregion
        //region SET_SLANG_WORDS Set 닉네임 비속어 리스트
        SET_SLANG_WORDS(state, payload) {
            state.slangWords = payload;
        },
        //endregion
    },

    actions : {
        //region GET_FORUMS Get 포럼 리스트
        GET_FORUMS(context, app) {
            app.callApi(2, 'GET', '/linker/forums', null,
                data => {
                    context.commit('SET_FORUMS', data);
                    app.changeCurrentForumIdx(data[0].forumIdx);
                },
                e => {
                    alert('포럼 정보를 가져오는 중 에러가 발생했습니다.');
                    console.log(e);
                });
        },
        //endregion
        //region GET_POSTINGS Get 포스팅 리스트
        GET_POSTINGS(context, payload) {
            const app = payload.app;
            const apiData = payload.apiData;

            app.callApi(2, 'GET', '/linker/postings', apiData,
                data => {
                    context.commit('SET_POSTING_COUNT', data.totalCount);
                    context.commit('SET_POSTINGS', data.postings);
                    context.commit('SET_LAST_PAGE', data.lastPage);
                });
        },
        //endregion
        //region DELETE_FORUM_INFOS 포럼 안내 리스트 삭제
        DELETE_FORUM_INFOS(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/forum/infos/delete', {infoIndexs:payload.idxs.join(',')}, () => {
                alert('삭제 되었습니다.');
                context.commit('DELETE_FORUM_INFOS', payload);
                document.getElementById('forumInfoAll').checked = false;
                app.$refs.forumInfo.checkedInfos = [];
                app.$refs.forumInfo.setTempInfos(app.currentForum.infos);
            });
        },
        //endregion
        //region MODIFY_FORUM_INFO_SORT 포럼 안내 정렬 수정
        MODIFY_FORUM_INFO_SORT(context, payload) {
            const app = payload.app;
            const apiData = {
                forumIndex : payload.forumIdx
            }
            payload.idxs.forEach((infoIdx, index) => {
                apiData[`infos[${index}].infoIndex`] = infoIdx;
                apiData[`infos[${index}].sortNo`] = index + 1;
            });
            app.callApi(2, 'POST', '/linker/forum/info/modify/sort', apiData, result => {
                if( result ) {
                    context.commit('MODIFY_FORUM_INFO_SORT', payload);
                    alert('수정되었습니다.');
                    app.$refs.forumInfo.setTempInfos();
                } else {
                    alert('오류가 발생했습니다');
                }
            });
        },
        //endregion
        //region UPDATE_POSTING 포스팅 수정
        UPDATE_POSTING(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/posting/update',
                payload.data, data => {
                    alert('저장되었습니다.');
                    app.managedPosting = null;
                    app.$refs.managePostingModal.closeModal();
                    context.commit('MODIFY_POSTING', payload.data);
                });
        },
        //endregion
        //region UPDATE_FIX_POSTING 고정 포스팅 수정
        UPDATE_FIX_POSTING(context, payload) {
            const app = payload.app;
            const data = {
                postingIndex : payload.postingIdx,
                positionNo : payload.fix.positionNo,
                startDate : payload.fix.startDate,
                endDate : payload.fix.endDate,
                useYn : true
            };

            app.callApi(2, 'POST', '/linker/fix/posting/update', data,
                () => {
                    alert('고정 되었습니다.');
                    app.checkedPostingIdx = 0;
                    app.getPostings();
                });
        },
        //endregion
        //region DELETE_POSTING 포스팅 삭제
        DELETE_POSTING(context, payload) {
            const app = payload.app;
            const success = function() {
                alert('삭제 되었습니다.');
                if( payload.isReportPosting ) {
                    context.commit('DELETE_REPORT_POSTING', [payload.postingIdx]);
                }
                app.getPostings();
            };

            app.callApi(2, 'POST', '/linker/posting/delete/' + payload.postingIdx, null, success);
        },
        //endregion
        //region DELETE_POSTINGS 포스팅 여러개 삭제
        DELETE_POSTINGS(context, payload) {
            const app = payload.app;
            const data = {
                postingsIndex : payload.postingIdxs.join(',')
            };
            const success = function() {
                alert('삭제 되었습니다.');
                context.commit('DELETE_REPORT_POSTING', payload.postingIdxs);
                app.getPostings();
            };

            app.callApi(2, 'POST', '/linker/postings/delete', data, success);
        },
        //endregion
        //region GET_REPORT_POSTINGS Get 신고 포스팅 리스트
        GET_REPORT_POSTINGS(context, payload) {
            payload.app.callApi(2, 'GET', '/linker/postings/report', {forumIndex:payload.forumIdx},
                data => {
                    context.commit('SET_REPORT_POSTINGS', data);
                });
        },
        //endregion
        //region UNBLOCK_POSTINGS 신고 포스팅 블락 해제
        UNBLOCK_POSTINGS(context, payload) {
            const data = {postingIndexes : payload.postingIdxs.join(',')};
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/blocked/posting/update', data,
                () => {
                    alert('해제 되었습니다.');
                    context.commit('DELETE_REPORT_POSTING', payload.postingIdxs);
                    app.$refs.manageReportPostings.checkedPostingIdxs = [];
                });
        },
        //endregion
        //region GET_FIX_POSTINGS 고정 포스팅 리스트 조회
        GET_FIX_POSTINGS(context, payload) {
            const app = payload.app;
            const url = `/linker/fix/postings/${payload.forumIdx}`;
            app.callApi(2, 'GET', url, null, data => {
                context.commit('SET_FIX_POSTINGS', data);
            })
        },
        //endregion
        //region GET_NICKNAMES 닉네임 리스트 조회
        GET_NICKNAMES(context, app) {
            app.callApi(2, 'GET', '/linker/nicknames', null, data => {
                context.commit('SET_WORDS', data);
            });
        },
        //endregion
        //region POST_NICKNAMES 닉네임 여러개 등록
        POST_NICKNAMES(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/nickname', payload.data, () => {
                alert('저장 되었습니다.');
                app.$refs.postNicknameModal.closeModal();
                if( payload.data.type === 'slang' ) {
                    context.dispatch('GET_NICKNAMES_SLANG', app);
                    app.$refs.manageNicknameSlangModal.openModal();
                } else {
                    context.dispatch('GET_NICKNAMES', app);
                    app.$refs.nicknameDictionaryModal.openModal();
                }
            });
        },
        //endregion
        //region DELETE_NICKNAMES 닉네임 여러개 삭제
        DELETE_NICKNAMES(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/nickname/delete', payload.data, () => {
                alert('삭제 되었습니다.');
                context.commit('DELETE_WORDS', payload.data);

                if( payload.data.type === 'slang' )
                    app.$refs.manageNicknameSlang.checkedWordIdxs = [];
                else
                    app.$refs.nicknameDictionary.clearCheck();
            });
        },
        //endregion
        //region MODIFY_NICKNAME 닉네임 수정
        MODIFY_NICKNAME(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/nickname/update', payload.data, () => {
                alert('수정 되었습니다.');
                context.commit('MODIFY_WORD', payload.data);
                app.$refs.modifyNicknameModal.closeModal();
            });
        },
        //endregion
        //region CLEAR_FIX_POSTINGS 포스팅 여러개 고정 해제
        CLEAR_FIX_POSTINGS(context, payload) {
            const app = payload.app;
            const data = {
                postingIndexs : payload.postingIdxs.join(',')
            }
            app.callApi(2, 'POST', '/linker/fix/postings/clear', data, () => {
                alert('고정 해제 되었습니다.');
                context.commit('DELETE_FIX_POSTINGS', payload.postingIdxs);
                app.getPostings();
                app.$refs.manageFixPostings.checkedPostingIdxs = [];
            });
        },
        //endregion
        //region MODIFY_FIX_POSTING_POSITIONS 고정 포스팅 위치 리스트 수정
        MODIFY_FIX_POSTING_POSITIONS(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/fix/postings/update', payload.data, () => {
                alert('수정 되었습니다.');
                context.dispatch('GET_FIX_POSTINGS', {
                    app,
                    forumIdx : app.currentForumIdx
                });
                app.getPostings();
            });
        },
        //endregion
        //region GET_NICKNAMES_SLANG 닉네임 비속어 리스트 조회
        GET_NICKNAMES_SLANG(context, app) {
            app.callApi(2, 'GET', '/linker/nicknames/slang', null, data => {
                context.commit('SET_SLANG_WORDS', data);
            });
        },
        //endregion
        //region GET_REPORT_COMMENTS 신고 댓글 리스트 조회
        GET_REPORT_COMMENTS(context, payload) {
            payload.app.callApi(2, 'GET', '/linker/comments/report', {forumIndex:payload.forumIdx},
                data => {
                    context.commit('SET_REPORT_COMMENTS', data);
                });
        },
        //endregion
        //region DELETE_REPORT_COMMENTS 신고 댓글 여러개 삭제
        DELETE_REPORT_COMMENTS(context, payload) {
            const app = payload.app;
            const data = { commentIndexs : payload.commentIdxs.join(',') };
            app.callApi(2, 'POST', '/linker/comments/delete', data, () => {
                alert('삭제 되었습니다.');
                context.commit('DELETE_REPORT_COMMENTS', payload.commentIdxs);
            });
        },
        //endregion
        //region CLEAR_BLOCK_REPORT_COMMENTS 신고 댓글 여러개 블락 해제
        CLEAR_BLOCK_REPORT_COMMENTS(context, payload) {
            const app = payload.app;
            const data = { commentIndexs : payload.commentIdxs.join(',') };
            app.callApi(2, 'POST', '/linker/comments/clear/block', data, () => {
                alert('해제 되었습니다.');
                context.commit('DELETE_REPORT_COMMENTS', payload.commentIdxs);
            });
        },
        //endregion
    }
});