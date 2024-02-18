<%@ language=vbscript %>
<% option explicit %>
<% Response.charset = "euc-kr"
'###############################################
' PageName : pop_mobile_slide_ajax.asp
' Discription : 모바일 slide ajax
' History : 2016-02-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/common/lib/pop_slide/classes/slidemanageCls.asp"-->
<%
Dim menu , mastercode , detailcode , prevDate , device
Dim oSlideManage
dim i 

mastercode  = request("mastercode")
detailcode  = request("detailcode")
prevDate    = request("prevDate")
menu        = request("menu")
device      = request("device")

if prevDate = "" then prevDate = date()
if device = "" then device = "P"

set oSlideManage = new SlideListCls
    oSlideManage.FPageSize = 10
	oSlideManage.FCurrPage = 1
	oSlideManage.FrectMasterCode = mastercode
	oSlideManage.FrectDetailCode = detailcode
    oSlideManage.FRectSelDate    = prevDate
    oSlideManage.FRectMenu       = menu
    oSlideManage.FRectDevice     = device
	oSlideManage.getSlideList()
	
%>
<style>
	.slideRegister .preview .swiper {position:relative; min-height:100px; background:#f9fafb;}
</style>
<script type="text/javascript">
$(function(){
	slideTemplate = new Swiper('.swiper-container',{
		loop:true,
		autoplay:3000,
		autoplayDisableOnInteraction:false,
		autoHeight:true,
		speed:800,
		pagination:'.pagination',
		paginationClickable:true,
		nextButton:'.btnNext',
		prevButton:'.btnPrev'
	});
});
</script>
<div class="evtSection swiper">
	<div class="swiper-container">
		<div class="swiper-wrapper">
		<% 
			for i=0 to oSlideManage.FResultCount - 1
        %>
		    <div class="swiper-slide">
                <% If oSlideManage.FItemList(i).Flinkurl <> "" Then %><a href="http://<%=chkiif(oSlideManage.FItemList(i).Fdevice="P","www","m")%>.10x10.co.kr<%=oSlideManage.FItemList(i).Flinkurl%>" target="_blank"><% End If %>
                    <% '// 동영상 %>
                    <% if oSlideManage.FItemList(i).Fisvideo = 1 then %>
                        <%=db2html(oSlideManage.FItemList(i).Fvideohtml)%>
                    <% else %>
                        <% '// 이미지 %>
                        <% if oSlideManage.FItemList(i).Fimageurl <> "" then %>
                            <img src="<%=oSlideManage.FItemList(i).Fimageurl%>" alt="" />
                        <% else %>
                            <img src="/images/admin_login_logo2.png" alt="" />
                        <% end if %>
                    <% end if %>
                <% If oSlideManage.FItemList(i).Flinkurl <> "" Then %></a><% End If %>
            </div>
		<% 
			next
		%>
		</div>
	</div>
	<div class="pagination"></div>
	<button type="button" class="slideNav btnPrev">preview</button>
	<button type="button" class="slideNav btnNext">next</button>
</div> 
<!-- #include virtual="/lib/db/dbclose.asp" -->