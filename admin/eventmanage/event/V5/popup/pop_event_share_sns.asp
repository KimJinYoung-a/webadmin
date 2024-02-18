<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_share_sns.asp
' Discription : 이벤트 SNS 공유 설정
' History : 2021.11.16 이전도
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
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
<script src="/vue/event/popup/shareSns/index.js"></script>