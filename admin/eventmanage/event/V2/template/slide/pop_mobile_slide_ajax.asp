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
<%
Dim eCode , topimg , btmimg , topaddimg 'floating img
Dim videolink, videosize
Dim strSql , sqlStr , isarrow
Dim arrowonoff : arrowonoff = false

	eCode = requestCheckvar(request("eC"),16)

	If eCode <> "" Then 
		strSql = "SELECT topimg , btmimg , topaddimg, videosize, videolink , isarrow " + vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event_slide_template where evt_code = '"& eCode &"' and device = 'M' " 
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			topimg	= rsget("topimg")
			btmimg	= rsget("btmimg")
			topaddimg	= rsget("topaddimg")
			videolink	= rsget("videolink")
			videosize	= rsget("videosize")
			isarrow		= rsget("isarrow")
		End If
		rsget.close()
	End if
	
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
<div class="evtSection evtTop">
	<div id="spantopimg">
	<%IF topimg <> "" THEN %>
	<img src="<%=topimg%>" alt="" />
	<a href="javascript:jsDelImg('topimg','spantopimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
	<%END IF%>
	</div>
</div>
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

			sqlStr = "SELECT slideimg , linkurl , sorting " + vbcrlf
			sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"' and device = 'M' " 
			sqlStr = sqlStr & " and isusing = 'Y' " 
			sqlStr = sqlStr & " order by sorting asc , idx asc " 
			rsget.Open sqlStr,dbget,1
			if Not(rsget.EOF or rsget.BOF) Then
				arrowonoff = true
				Do Until rsget.eof
		%>
			<div class="swiper-slide"><% If rsget("linkurl") <> "" Then %><a href="http://m.10x10.co.kr<%=rsget("linkurl")%>" target="_blank"><% End If %><img src="<%=rsget("slideimg")%>" alt="" /><% If rsget("linkurl") <> "" Then %></a><% End If %></div>
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
		<% if isarrow = 1 then %>
		<button type="button" class="slideNav btnPrev">preview</button>
		<button type="button" class="slideNav btnNext">next</button>
		<% end if %>
	<% End If %>
</div> 
<div class="evtSection evtBtm">
	<div id="spanbtmimg">
	<%IF btmimg <> "" THEN %>
	<img src="<%=btmimg%>" alt="" />
	<a href="javascript:jsDelImg('btmimg','spanbtmimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
	<%END IF%>
	</div>
</div>
<% if videosize<>"" And videolink<>"" then %>
<div class='vod-wrap <% if videosize="W" then %>shape-rtgl<% end if %>'>
	<div class="vod"><iframe src="<%=videolink%>?rel=0&showinfo=0&playsinline=1&title=0&byline=0&portrait=0" frameborder="0" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe></div><!-- for dev msg : 동영상 경로 뒤 ? 이하 옵션값 무조건 붙여주세요 -->
</div>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
