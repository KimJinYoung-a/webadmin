<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<style>
    #tagDiv p:hover {background: #33FF33}
</style>

<script>
    document.domain = "10x10.co.kr";
</script>
<link rel="stylesheet" type="text/css" href="/css/adminRenewal.css?v=1.009"/>
<link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous">
<link rel="stylesheet" href="/js/jqueryui/css/jquery-ui.css">
<script src="https://unpkg.com/lodash@4.13.1/lodash.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.auto.min.js"></script>

<script src="/vue/2.5/vue.min.js"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<script src="/vue/vuex.min.js"></script>

<script src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>

<div id="app"></div>

<script src="/vue/common/common.js?v=1.01"></script>
<script src="/vue/components/common/pagination.js?v=1.00"></script>
<script src="/vue/components/common/modal.js?v=1.00"></script>
<script src="/vue/play/contents/listWrite.js?v=1.31"></script>
<script src="/vue/play/contents/contentWrite.js?v=1.03"></script>
<script src="/vue/play/contents/nicknameWrite.js?v=1.00"></script>

<script src="/vue/play/contents/store.js?v=1.26"></script>
<script src="/vue/play/contents/index.js?v=1.55"></script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->