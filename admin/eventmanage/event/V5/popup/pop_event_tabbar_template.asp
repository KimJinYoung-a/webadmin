<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_img_text_template.asp
' Discription : I형(통합형) 이벤트 기획전 탭바 설정 팝업
' History : 2021.11.05 이전도
' History : 2022.04.21 김형태
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->

<script src="/vue/2.5/vue.min.js"></script>
<script src="/vue/vue.lazyimg.min.js"></script>
<script src="/vue/vuex.min.js"></script>

<script src="//cdn.jsdelivr.net/npm/sortablejs@1.8.4/Sortable.min.js"></script>
<script src="//cdnjs.cloudflare.com/ajax/libs/Vue.Draggable/2.20.0/vuedraggable.umd.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>

<%
Dim eCode, menuidx, device
eCode = Request("eC")
menuidx = Request("menuidx")
device = Request("device")
%>
<script>
const masterIndex = '<%=menuidx%>';
const device = '<%=device%>';
</script>

<style>
    .tabV19 a {cursor:pointer;}
    .tabbarTemplate td {margin: 2px 20px 2px 0;font-size: 1.25rem;}
    .tabbarTemplate td .contentGroup:not(:first-child) {margin-top:10px;}
    .tabbarTemplate td .backTypeArea {margin-bottom: 10px;}
    .tabbarTemplate td .chedkLabel {cursor: pointer;display: inline-block;margin-right: 10px;}
    .tabbarTemplate td .chedkLabel input[type=radio] {position: relative;top: -2px;}
    .tabbarTemplate td .settingArea strong {display: inline-block;width: 70px;}
    .tabbarTemplate td .settingArea:not(:first-child) {margin-top: 10px;}
    .tabbarTemplate td input[type=text] {width: 100%;padding: 5px 10px;font-size: 15px;}
    .tabbarTemplate td input[type=text].short {width: 100px;}
    .tabbarTemplate td input[type=text].shortest {width: 50px;}
    .tabbarTemplate td input[type=file]{display: none;}

    .tabbarTemplate td.preview {background-color: #e6e8ee;}
    .tabbarTemplate td.preview.mobile .sliderArea {width: 375px;background-color: #fff;position: relative;display: inline-block;}
    .tabbarTemplate td.preview.mobile .etcArea {min-height: 300px;max-height: 500px;background-color: #fff;width: 375px;overflow: hidden;}
    .tabbarTemplate td.preview.pc .sliderArea {width: 1024px;background-color: #fff;position: relative;display: inline-block;}
    .tabbarTemplate td.preview.pc .etcArea {min-height: 200px;max-height: 300px;background-color: #fff;width: 1024px;overflow: hidden;}
    .tabbarTemplate td.preview .etcArea img {width: 100%;}
    .tabbarTemplate td.preview .inputArea {margin-top: 25px;}
    .tabbarTemplate td.preview .inputArea input {width: 160px;padding: 5px 10px;display: inline-block;margin-right: 15px;}

    .tabbarTemplate .back-image {display: block;max-height: 200px;margin-top: 10px;box-shadow: 4px 4px 4px #aaa;}

    .preview .swiper-container {overflow: hidden;position: relative;width: 100%;margin: 0 auto;z-index: 1;}
    .preview .swiper-container button {position:absolute; top:0; z-index:100; width:37px; height:100%; background:#fff url(//webimage.10x10.co.kr/eventIMG/2020/102974/btn_nav_m.png) 50%/100% no-repeat; font-size:0;border: none;}
    .preview .swiper-container button.btn-prev {left:0;}
    .preview .swiper-container button.btn-next {right:0; transform:rotate(180deg);}
    .preview .swiper-wrapper {position: relative;width: 100%;z-index: 1;display: flex;}
    .preview .swiper-slide {position: relative;flex-shrink: 0;height: 100%;text-align: center;color: #c3c3c3;line-height: 50px;}

    .modal {height: 100%;position: relative;z-index: 150;}
    .modal .modal-overlay {position: fixed;top: -1px;bottom: -1px;left: 0;width: 100vw;background: #000;opacity: 0.5;}
    .modal .modal-wrap {position: fixed;min-height: 150px;min-width: 600px;left: 50%;top: 50%;transform: translate(-50%,-50%);background: #fff;border-radius: 10px;padding: 15px 20px 50px 20px;}
    .modal .manage-button-area {display: flex;justify-content: right;margin-bottom: 10px;}
    .modal-close-btn {position: absolute;right: -28px;top: -24px;border: none;background: url(//fiximage.10x10.co.kr/web2021/anniv2021/icon_pop_close.png) 0 0 no-repeat;background-size: 100%;width: 24px;height: 24px;cursor: pointer;}
    .modal-title {margin-bottom: 30px;}
    .modal-title h3 {font-size: 17px;}
    .modal-title .add-descr {color: #ff3333;font-size: 11px;font-weight: bold;margin-top: 3px;}
    .modal .manage-area button {display: inline-block;padding: 7px 10px;line-height: 1;margin-left: 5px;border: 1px solid #4075ff;background-color: #4075ff;color: #fff;cursor: pointer;}
    .modal .manage-area button.add {border: 1px solid #444;background-color: #444;}
    .modal .manage-area .post-table button.add {width: 50px;}
    .modal .manage-area input[type=text] {width: 95%;padding: 5px 10px;font-size: 13px;}
    .modal .manage-area table {width: 100%;}
    .modal .manage-area table th, .modal .manage-area table td {padding: 10px 0;text-align: center;}
    .modal .manage-area table th {background-color: #e6e8ee;font-size: 13px;}
    .modal .manage-area table.post-table {margin-top: 10px;}
    .modal .manage-area table.post-table button {width: 80px;}
    .modal .manage-area table.post-table button.imageButton {width: 110px;}

    .post-item .tableV19A > tbody > tr > th {font-size: 13px; padding: 12px 15px;}
    .post-item .tableV19A > tbody > tr > td {font-size: 13px; padding: 5px 15px;}
    .post-item .tableV19A > tbody > tr > td input[type=text] {font-size: 13px;padding: 7px 10px;}
    .post-item .popBtnWrapV19 {padding: 0;}
    .post-item .popBtnWrapV19 button {padding: 16px 41px;}
</style>
<div id="app"></div>

<script src="/vue/common/api_mixins.js"></script>
<script src="/vue/components/linker/modal.js"></script>
<script src="/vue/components/event/modal_tabbar_manage_items.js"></script>
<script src="/vue/components/event/modal_tabbar_post.js"></script>
<script src="/vue/event/tabbar/app.js"></script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
