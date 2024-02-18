<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : TIMESALE
'	History		: 2021.05.31 ±èÇüÅÂ
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

<script>
    //document.domain = "10x10.co.kr";
</script>

<!-- ´Þ·Â -->
<link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/timepicker/1.3.5/jquery.timepicker.min.css">
<script src="//cdnjs.cloudflare.com/ajax/libs/timepicker/1.3.5/jquery.timepicker.min.js"></script>

<script src="/vue/common/common.js?v=1.00"></script>
<script src="/vue/components/common/pagination.js?v=1.0"></script>
<script src="/vue/components/common/scrollModal.js?v=1.00"></script>

<script src="/vue/tenten_exclusive/write.js?v=1.00"></script>
<script src="/vue/tenten_exclusive/write_main.js?v=1.00"></script>
<script src="/vue/tenten_exclusive/write_detail.js?v=1.00"></script>

<script src="/vue/tenten_exclusive/store.js?v=1.00"></script>
<script src="/vue/tenten_exclusive/index.js?v=1.00"></script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->