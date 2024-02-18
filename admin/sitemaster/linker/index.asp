<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : LINKER 包府
'	History		: 2021.10.14 捞傈档 积己
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
</p>
<link rel="stylesheet" type="text/css" href="/css/linker.css">
<link rel="stylesheet" href="https://cdn.materialdesignicons.com/3.6.95/css/materialdesignicons.min.css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<% IF application("Svr_Info") = "Dev" THEN %>
<script src="https://unpkg.com/vue"></script>
<script src="https://unpkg.com/vuex"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<% Else %>
<script src="/vue/2.5/vue.min.js"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<script src="/vue/vuex.min.js"></script>
<% End If %>
<script src="//cdn.jsdelivr.net/npm/sortablejs@1.8.4/Sortable.min.js"></script>
<script src="//cdnjs.cloudflare.com/ajax/libs/Vue.Draggable/2.20.0/vuedraggable.umd.min.js"></script>

<div id="app"></div>

<script src="/vue/common/api_mixins.js"></script>
<script src="/vue/components/common/date_picker.js"></script>
<script src="/vue/components/linker/list_forum.js"></script>
<script src="/vue/components/linker/post_forum.js"></script>
<script src="/vue/components/linker/forum_info.js"></script>
<script src="/vue/components/linker/manage_forum_sort.js"></script>
<script src="/vue/components/linker/post_forum_info.js"></script>
<script src="/vue/components/linker/posting_list.js"></script>
<script src="/vue/components/linker/posting.js"></script>
<script src="/vue/components/linker/manage_posting.js"></script>
<script src="/vue/components/linker/manage_fix_posting.js"></script>
<script src="/vue/components/linker/manage_fix_postings.js"></script>
<script src="/vue/components/linker/manage_report_postings.js"></script>
<script src="/vue/components/linker/manage_report_comments.js"></script>
<script src="/vue/components/linker/manage_nickname_dictionary.js"></script>
<script src="/vue/components/linker/manage_nickname_slang.js"></script>
<script src="/vue/components/linker/post_words.js"></script>
<script src="/vue/components/linker/modify_word.js"></script>
<script src="/vue/components/linker/modal.js"></script>
<script src="/vue/components/linker/pagination.js"></script>
<script src="/vue/linker/store.js"></script>
<script src="/vue/linker/index.js"></script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->