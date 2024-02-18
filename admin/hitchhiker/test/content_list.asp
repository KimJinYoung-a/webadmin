<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER CONTENT LIST TEST
'	History		: 2021.01.14 이전도 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminRenewal.css?v=1.009"/>
<link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous">
<link rel="stylesheet" href="/js/jqueryui/css/jquery-ui.css">
<script src="https://unpkg.com/lodash@4.13.1/lodash.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.auto.min.js"></script>
<% IF application("Svr_Info") = "Dev" THEN %>
<script src="https://unpkg.com/vue"></script>
<script src="https://unpkg.com/vuex"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<% Else %>
<script src="/vue/2.5/vue.min.js"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<script src="/vue/vuex.min.js"></script>
<% End If %>
<script src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<div id="app"></div>
<script src="/vue/common/common.js?v=1.0"></script>
<script src="/vue/components/common/pagination.js?v=1.0"></script>
<script src="/vue/components/common/modal.js?v=1.001"></script>
<script src="/vue/hitchhiker/contents/write.js?v=1.012"></script>
<script src="/vue/hitchhiker/wallpaper/write.js?v=1.000"></script>
<script src="/vue/hitchhiker/wallpaper/hitchhiker_li_size.js?v=1.000"></script>
<script src="/vue/hitchhiker/contents/store.js?v=1.005"></script>
<script src="/vue/hitchhiker/contents/index.js?v=1.0107"></script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->