<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : [ON] 八祸 包府 > 惑前 富赣府 包府
'	History		: 2021.11.11 捞傈档 积己
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
</p>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/vue/search/itemNamePrefix/index.css" rel="stylesheet">
<% IF application("Svr_Info") = "Dev" THEN %>
<script src="https://unpkg.com/vue"></script>
<script src="https://unpkg.com/vuex"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<% Else %>
<script src="/vue/2.5/vue.min.js"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<script src="/vue/vuex.min.js"></script>
<% End If %>

<div id="app"></div>

<script src="/vue/common/api_mixins.js"></script>
<script src="/vue/components/common/date_picker.js"></script>
<script src="/vue/components/linker/pagination.js"></script>
<script src="/vue/components/linker/modal.js"></script>
<script src="/vue/components/search/itemNamePrefix/item_name_prefix_search.js"></script>
<script src="/vue/components/search/itemNamePrefix/item_name_prefix_result_item.js"></script>
<script src="/vue/components/search/itemNamePrefix/item_name_prefix_result.js"></script>
<script src="/vue/components/search/itemNamePrefix/item_name_prefix_post.js"></script>
<script src="/vue/components/search/itemNamePrefix/item_name_prefix_manage_item.js"></script>
<script src="/vue/components/search/itemNamePrefix/item_name_prefix_search_item.js"></script>
<script src="/vue/search/itemNamePrefix/store.js"></script>
<script src="/vue/search/itemNamePrefix/index.js"></script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->