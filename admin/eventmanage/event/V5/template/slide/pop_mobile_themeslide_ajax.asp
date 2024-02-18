<%@ language=vbscript %>
<% option explicit %>
<% Response.charset = "euc-kr"
'###############################################
' PageName : pop_pcweb_themeslide_ajax.asp
' Discription : 모바일 slide ajax
' History : 2019-02-12 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , topimg , btmimg , topaddimg 'floating img
Dim sqlStr, menuidx
Dim arrowonoff : arrowonoff = false
eCode = requestCheckvar(request("eC"),16)
menuidx = requestCheckvar(request("menuidx"),10)
%>
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
	<div class="txt" id="spantopaddimg">
		<%IF topaddimg <> "" THEN %>
		<img src="<%=topaddimg%>" alt="" />
		<a href="javascript:jsDelImg('topaddimg','spantopaddimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
		<%END IF%>
	</div>
	<div class="swiper-container">
		<div class="swiper-wrapper">
		<% 
			If eCode <> "" Then
				sqlStr = "SELECT imgurl, videoFullLink" + vbcrlf
				sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_multi_contents] where menuidx='"& menuidx &"'" + vbcrlf
				sqlStr = sqlStr & " and isusing = 'Y' and device ='M'" + vbcrlf
				sqlStr = sqlStr & " order by viewidx asc, idx asc"
				rsget.Open sqlStr,dbget,1
				if Not(rsget.EOF or rsget.BOF) Then
					arrowonoff = true
					Do Until rsget.eof
		%>
			<% If rsget("imgurl") <> "" Then %>
			<div class="swiper-slide"><img src="<%=rsget("imgurl")%>" alt="" /></div>
			<% else %>
			<div class="swiper-slide"><%=rsget("videoFullLink")%></div>
			<% End If %>
		<% 
					rsget.movenext
					Loop
				End If
				rsget.close
			End If
		%>
		</div>
	</div>
	<% If arrowonoff Then %>
	<div class="pagination"></div>
	<button type="button" class="slideNav btnPrev">preview</button>
	<button type="button" class="slideNav btnNext">next</button>
	<% End If %>
</div> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
