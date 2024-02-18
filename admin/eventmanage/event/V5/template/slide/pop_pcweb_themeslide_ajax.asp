<%@ language=vbscript %>
<% option explicit %>
<% Response.charset = "euc-kr"
'###############################################
' PageName : pop_pcweb_themeslide_ajax.asp
' Discription : PCWEB slide ajax
' History : 2019-02-11 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode 
Dim strSql , sqlStr , gubuncls, menuidx
Dim topimg , topaddimg , btmYN , btmimg , btmcode , btmaddimg , pcadd1 , gubun

	eCode = requestCheckvar(request("eC"),16)
	gubun = requestCheckvar(request("gU"),16)
	menuidx = requestCheckvar(request("menuidx"),10)

	If gubun = "" Then gubun = 1 '와이드 슬라이드
	
	If gubun = "1" Then
		gubuncls = "wideSlide" '//와이드 슬라이드
	ElseIf gubun = "2" Then
		gubuncls = "wideSwipe" '//와이드+풀단 슬라이드
	ElseIf gubun = "3" Then
		gubuncls = "fullSlide" '//풀단 슬라이드
	End If

	'//구분 값에 따른 입력 필드 노출
	if gubun = "1" or gubun = "2" then
		Response.write "<script>$('.punit1').show();$('.punit2').hide();</script>"
	elseif gubun = "3" then
		Response.write "<script>$('.punit1').hide();$('.punit2').show();</script>"
		Response.write "<script>$('#spanbtmaddimg_bg').css('background-image','')</script>"
		Response.write "<script>$('#spantopaddimg_bg').css('background-image','')</script>"
	end if

%>
<script type="text/javascript">
$(function(){
	$("#gubun").val(<%=gubun%>); //ajax 호출후 gubun form에 저장
	// wide swipe
	var evtSwiper = new Swiper('.wideSwipe .swiper-container',{
		loop:true,
		slidesPerView:'auto',
		centeredSlides:true,
		speed:1200, 
		autoplay:3500,
		simulateTouch:false,
		pagination:'.wideSwipe .pagination',
		paginationClickable:true,
		nextButton:'.wideSwipe .btnNext',
		prevButton:'.wideSwipe .btnPrev'
	})
	$('.wideSwipe .btnPrev').on('click', function(e){
		e.preventDefault();
		evtSwiper.swipePrev();
	})
	$('.wideSwipe .btnNext').on('click', function(e){
		e.preventDefault();
		evtSwiper.swipeNext();
	});

	// wide slide
	$('.wideSlide .swiper-wrapper').slidesjs({
		width:1920,
		height:800,
		pagination:{effect:'fade'},
		navigation:{effect:'fade'},
		play:{interval:3000, effect:'fade', auto:true},
		effect:{fade: {speed:1200, crossfade:true}
		},
		callback: {
			complete: function(number) {
				var pluginInstance = $('.wideSlide .swiper-wrapper').data('plugin_slidesjs');
				setTimeout(function() {
					pluginInstance.play(true);
				}, pluginInstance.options.play.interval);
			}
		}
	});

	// full slide
	$('.fullSlide .swiper-wrapper').slidesjs({
		width:520,
		height:322,
		pagination:{effect:'fade'},
		navigation:{effect:'fade'},
		play:{interval:3000, effect:'fade', auto:true},
		effect:{fade: {speed:1200, crossfade:true}
		},
		callback: {
			complete: function(number) {
				var pluginInstance = $('.fullSlide .swiper-wrapper').data('plugin_slidesjs');
				setTimeout(function() {
					pluginInstance.play(true);
				}, pluginInstance.options.play.interval);
			}
		}
	});
});
</script>
<div class="<%=gubuncls%>">

	<div class="evtSection swiper">
		<div class="swiper-container" <% If pcadd1 <>"" Then %>style="background-image:url(<%=pcadd1%>);"<% End If %> id="spanpcadd1_bg">
			<div class="swiper-wrapper">
			<% 
				If eCode <> "" Then 
					sqlStr = "SELECT imgurl, videoFullLink" + vbcrlf
					sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_multi_contents] where menuidx='"& menuidx &"'" + vbcrlf
					sqlStr = sqlStr & " and isusing = 'Y' and device ='W'" + vbcrlf
					sqlStr = sqlStr & " order by viewidx asc, idx asc"
					rsget.Open sqlStr,dbget,1
					if Not(rsget.EOF or rsget.BOF) Then
						Do Until rsget.eof
			%>
				<div class="swiper-slide">
					<% If rsget("imgurl") <> "" Then %>
						<img src="<%=rsget("imgurl")%>" alt="" />
					<% else %>
						<%=rsget("videoFullLink")%>
					<% End If %>
				</div>
			<% 
						rsget.movenext
						Loop
					End If
					rsget.close
				End If
			%>
			</div>
			<div class="pagination"></div>
			<button class="slideNav btnPrev">이전</button>
			<button class="slideNav btnNext">다음</button>
			<div class="mask left"></div>
			<div class="mask right"></div>
		</div>
	</div>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->