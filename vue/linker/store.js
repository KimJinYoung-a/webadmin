const store = new Vuex.Store({
    state : {
        forums : [], // ���� ����Ʈ

        postingCount : 0, // ������ ����
        postings : [], // ������ ����Ʈ
        lastPage : 0, // ������ ������

        reportPostingCount : 0, // �Ű� ������ ����
        reportPostings : [], // �Ű� ������ ����Ʈ
        reportComments : [], // �Ű� ��� ����Ʈ
        fixPostings : [], // ���� ������ ����Ʈ

        words1 : [], // �г��� �ܾ�1 ����Ʈ
        words2 : [], // �г��� �ܾ�2 ����Ʈ
        slangWords : [], // �г��� ��Ӿ� ����Ʈ
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
        //region SET_FORUMS Set ���� ����Ʈ
        SET_FORUMS(state, forums) {
            state.forums = forums;
        },
        //endregion
        //region MODIFY_FORUMS_SORT ���� ����Ʈ ���ļ��� ����
        MODIFY_FORUMS_SORT(state, sortForums) {
            sortForums.forEach(s => {
                const forum = state.forums.find(f => f.forumIdx === s.forumIdx);
                forum.sortNo = s.sortNo;
            });
        },
        //endregion
        //region ADD_FORUM_INFO ���� �ȳ� �߰�
        ADD_FORUM_INFO(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            forum.infos.push(payload.info);
        },
        //endregion
        //region MODIFY_FORUM_INFO ���� �ȳ� ����
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
        //region DELETE_FORUM_INFOS ���� �ȳ� ����
        DELETE_FORUM_INFOS(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            forum.infos = forum.infos.filter(i => payload.idxs.indexOf(i.infoIdx) === -1);
        },
        //endregion
        //region MODIFY_FORUM_INFO_SORT ���� �ȳ� ���� ����
        MODIFY_FORUM_INFO_SORT(state, payload) {
            const forum = state.forums.find(f => f.forumIdx === payload.forumIdx);
            payload.idxs.forEach((infoIdx, index) => {
                const info = forum.infos.find(info => info.infoIdx === infoIdx);
                info.sortNo = index + 1;
            });
            forum.infos.sort((i1, i2) => i1.sortNo - i2.sortNo);
        },
        //endregion
        //region SET_POSTING_COUNT Set ������ ��ü ����
        SET_POSTING_COUNT(state, count) {
            state.postingCount = count;
        },
        //endregion
        //region SET_POSTINGS Set ������ ����Ʈ
        SET_POSTINGS(state, postings) {
            state.postings = postings;
        },
        //endregion
        //region SET_LAST_PAGE Set ������ ������
        SET_LAST_PAGE(state, page) {
            state.lastPage = page;
        },
        //endregion
        //region MODIFY_POSTING ������ ����
        MODIFY_POSTING(state, payload) {
            const posting = state.postings.find(p => p.postingIdx === payload.postingIndex);
            posting.postingContents = payload.postingCotents;
            posting.fixed = payload.useYn;
            posting.startDate = payload.startDate;
            posting.endDate = payload.endDate;
            posting.positionNo = payload.positionNo;
        },
        //endregion
        //region SET_REPORT_POSTINGS Set �Ű� ������ ����Ʈ&����
        SET_REPORT_POSTINGS(state, payload) {
            state.reportPostingCount = payload.totalCount;
            state.reportPostings = payload.postings;
        },
        //endregion
        //region SET_REPORT_COMMENTS Set �Ű� ��� ����Ʈ
        SET_REPORT_COMMENTS(state, payload) {
            state.reportComments = payload;
        },
        //endregion
        //region DELETE_REPORT_POSTING �Ű� ������ ����
        DELETE_REPORT_POSTING(state, postingIdxs) {
            state.reportPostings = state.reportPostings
                                            .filter(posting => postingIdxs.indexOf(posting.postingIdx) === -1);
        },
        //endregion
        //region SET_FIX_POSTINGS Set ���� ������ ����Ʈ
        SET_FIX_POSTINGS(state, payload) {
            state.fixPostings = payload;
        },
        //endregion
        //region DELETE_REPORT_COMMENTS �Ű� ��� ����
        DELETE_REPORT_COMMENTS(state, payload) {
            state.reportComments = state.reportComments.filter(c => payload.indexOf(c.commentIdx) === -1);
        },
        //endregion
        //region SET_WORDS Set �г��� �ܾ� ����Ʈ
        SET_WORDS(state, payload) {
            state.words1 = payload.words1;
            state.words2 = payload.words2;
        },
        //endregion
        //region DELETE_WORDS �г��� �ܾ� ����
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
        //region MODIFY_WORD �г��� �ܾ� ����
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
        //region DELETE_FIX_POSTINGS ���� ������ ����Ʈ ������ ����
        DELETE_FIX_POSTINGS(state, postingIdxs) {
            state.fixPostings = state.fixPostings.filter(p => postingIdxs.indexOf(p.postingIdx) === -1);
        },
        //endregion
        //region SET_SLANG_WORDS Set �г��� ��Ӿ� ����Ʈ
        SET_SLANG_WORDS(state, payload) {
            state.slangWords = payload;
        },
        //endregion
    },

    actions : {
        //region GET_FORUMS Get ���� ����Ʈ
        GET_FORUMS(context, app) {
            app.callApi(2, 'GET', '/linker/forums', null,
                data => {
                    context.commit('SET_FORUMS', data);
                    app.changeCurrentForumIdx(data[0].forumIdx);
                },
                e => {
                    alert('���� ������ �������� �� ������ �߻��߽��ϴ�.');
                    console.log(e);
                });
        },
        //endregion
        //region GET_POSTINGS Get ������ ����Ʈ
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
        //region DELETE_FORUM_INFOS ���� �ȳ� ����Ʈ ����
        DELETE_FORUM_INFOS(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/forum/infos/delete', {infoIndexs:payload.idxs.join(',')}, () => {
                alert('���� �Ǿ����ϴ�.');
                context.commit('DELETE_FORUM_INFOS', payload);
                document.getElementById('forumInfoAll').checked = false;
                app.$refs.forumInfo.checkedInfos = [];
                app.$refs.forumInfo.setTempInfos(app.currentForum.infos);
            });
        },
        //endregion
        //region MODIFY_FORUM_INFO_SORT ���� �ȳ� ���� ����
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
                    alert('�����Ǿ����ϴ�.');
                    app.$refs.forumInfo.setTempInfos();
                } else {
                    alert('������ �߻��߽��ϴ�');
                }
            });
        },
        //endregion
        //region UPDATE_POSTING ������ ����
        UPDATE_POSTING(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/posting/update',
                payload.data, data => {
                    alert('����Ǿ����ϴ�.');
                    app.managedPosting = null;
                    app.$refs.managePostingModal.closeModal();
                    context.commit('MODIFY_POSTING', payload.data);
                });
        },
        //endregion
        //region UPDATE_FIX_POSTING ���� ������ ����
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
                    alert('���� �Ǿ����ϴ�.');
                    app.checkedPostingIdx = 0;
                    app.getPostings();
                });
        },
        //endregion
        //region DELETE_POSTING ������ ����
        DELETE_POSTING(context, payload) {
            const app = payload.app;
            const success = function() {
                alert('���� �Ǿ����ϴ�.');
                if( payload.isReportPosting ) {
                    context.commit('DELETE_REPORT_POSTING', [payload.postingIdx]);
                }
                app.getPostings();
            };

            app.callApi(2, 'POST', '/linker/posting/delete/' + payload.postingIdx, null, success);
        },
        //endregion
        //region DELETE_POSTINGS ������ ������ ����
        DELETE_POSTINGS(context, payload) {
            const app = payload.app;
            const data = {
                postingsIndex : payload.postingIdxs.join(',')
            };
            const success = function() {
                alert('���� �Ǿ����ϴ�.');
                context.commit('DELETE_REPORT_POSTING', payload.postingIdxs);
                app.getPostings();
            };

            app.callApi(2, 'POST', '/linker/postings/delete', data, success);
        },
        //endregion
        //region GET_REPORT_POSTINGS Get �Ű� ������ ����Ʈ
        GET_REPORT_POSTINGS(context, payload) {
            payload.app.callApi(2, 'GET', '/linker/postings/report', {forumIndex:payload.forumIdx},
                data => {
                    context.commit('SET_REPORT_POSTINGS', data);
                });
        },
        //endregion
        //region UNBLOCK_POSTINGS �Ű� ������ ��� ����
        UNBLOCK_POSTINGS(context, payload) {
            const data = {postingIndexes : payload.postingIdxs.join(',')};
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/blocked/posting/update', data,
                () => {
                    alert('���� �Ǿ����ϴ�.');
                    context.commit('DELETE_REPORT_POSTING', payload.postingIdxs);
                    app.$refs.manageReportPostings.checkedPostingIdxs = [];
                });
        },
        //endregion
        //region GET_FIX_POSTINGS ���� ������ ����Ʈ ��ȸ
        GET_FIX_POSTINGS(context, payload) {
            const app = payload.app;
            const url = `/linker/fix/postings/${payload.forumIdx}`;
            app.callApi(2, 'GET', url, null, data => {
                context.commit('SET_FIX_POSTINGS', data);
            })
        },
        //endregion
        //region GET_NICKNAMES �г��� ����Ʈ ��ȸ
        GET_NICKNAMES(context, app) {
            app.callApi(2, 'GET', '/linker/nicknames', null, data => {
                context.commit('SET_WORDS', data);
            });
        },
        //endregion
        //region POST_NICKNAMES �г��� ������ ���
        POST_NICKNAMES(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/nickname', payload.data, () => {
                alert('���� �Ǿ����ϴ�.');
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
        //region DELETE_NICKNAMES �г��� ������ ����
        DELETE_NICKNAMES(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/nickname/delete', payload.data, () => {
                alert('���� �Ǿ����ϴ�.');
                context.commit('DELETE_WORDS', payload.data);

                if( payload.data.type === 'slang' )
                    app.$refs.manageNicknameSlang.checkedWordIdxs = [];
                else
                    app.$refs.nicknameDictionary.clearCheck();
            });
        },
        //endregion
        //region MODIFY_NICKNAME �г��� ����
        MODIFY_NICKNAME(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/nickname/update', payload.data, () => {
                alert('���� �Ǿ����ϴ�.');
                context.commit('MODIFY_WORD', payload.data);
                app.$refs.modifyNicknameModal.closeModal();
            });
        },
        //endregion
        //region CLEAR_FIX_POSTINGS ������ ������ ���� ����
        CLEAR_FIX_POSTINGS(context, payload) {
            const app = payload.app;
            const data = {
                postingIndexs : payload.postingIdxs.join(',')
            }
            app.callApi(2, 'POST', '/linker/fix/postings/clear', data, () => {
                alert('���� ���� �Ǿ����ϴ�.');
                context.commit('DELETE_FIX_POSTINGS', payload.postingIdxs);
                app.getPostings();
                app.$refs.manageFixPostings.checkedPostingIdxs = [];
            });
        },
        //endregion
        //region MODIFY_FIX_POSTING_POSITIONS ���� ������ ��ġ ����Ʈ ����
        MODIFY_FIX_POSTING_POSITIONS(context, payload) {
            const app = payload.app;
            app.callApi(2, 'POST', '/linker/fix/postings/update', payload.data, () => {
                alert('���� �Ǿ����ϴ�.');
                context.dispatch('GET_FIX_POSTINGS', {
                    app,
                    forumIdx : app.currentForumIdx
                });
                app.getPostings();
            });
        },
        //endregion
        //region GET_NICKNAMES_SLANG �г��� ��Ӿ� ����Ʈ ��ȸ
        GET_NICKNAMES_SLANG(context, app) {
            app.callApi(2, 'GET', '/linker/nicknames/slang', null, data => {
                context.commit('SET_SLANG_WORDS', data);
            });
        },
        //endregion
        //region GET_REPORT_COMMENTS �Ű� ��� ����Ʈ ��ȸ
        GET_REPORT_COMMENTS(context, payload) {
            payload.app.callApi(2, 'GET', '/linker/comments/report', {forumIndex:payload.forumIdx},
                data => {
                    context.commit('SET_REPORT_COMMENTS', data);
                });
        },
        //endregion
        //region DELETE_REPORT_COMMENTS �Ű� ��� ������ ����
        DELETE_REPORT_COMMENTS(context, payload) {
            const app = payload.app;
            const data = { commentIndexs : payload.commentIdxs.join(',') };
            app.callApi(2, 'POST', '/linker/comments/delete', data, () => {
                alert('���� �Ǿ����ϴ�.');
                context.commit('DELETE_REPORT_COMMENTS', payload.commentIdxs);
            });
        },
        //endregion
        //region CLEAR_BLOCK_REPORT_COMMENTS �Ű� ��� ������ ��� ����
        CLEAR_BLOCK_REPORT_COMMENTS(context, payload) {
            const app = payload.app;
            const data = { commentIndexs : payload.commentIdxs.join(',') };
            app.callApi(2, 'POST', '/linker/comments/clear/block', data, () => {
                alert('���� �Ǿ����ϴ�.');
                context.commit('DELETE_REPORT_COMMENTS', payload.commentIdxs);
            });
        },
        //endregion
    }
});