<%@ language=vbscript %>
<% option explicit %>
<% Response.charset = "euc-kr"
'###############################################
' PageName : pop_pcweb_slide_ajax.asp
' Discription : PCWEB slide ajax
' History : 2016-02-16 ����ȭ
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
	menuidx = requestCheckvar(request("menuidx"),16)
	if menuidx="" or isnull(menuidx) then menuidx=0
	If gubun = "" Then gubun = 1 '���̵� �����̵�
	If gubun = "0" Then gubun = "3" '���̵� �����̵�
	
	If gubun = "1" Then
		gubuncls = "wideSlide" '//���̵� �����̵�
	ElseIf gubun = "2" Then
		gubuncls = "wideSwipe" '//���̵�+Ǯ�� �����̵�
	ElseIf gubun = "3" Then
		gubuncls = "fullSlide" '//Ǯ�� �����̵�
	End If

	'//���� ���� ���� �Է� �ʵ� ����
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
	$("#gubun").val(<%=gubun%>); //ajax ȣ���� gubun form�� ����
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

				sqlStr = "SELECT slideimg , linkurl , sorting " & vbcrlf
				sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_top_slide_addimage] where evt_code = '"& eCode &"' " & vbcrlf
				sqlStr = sqlStr & " and isusing = 'Y' and device ='W' and menuidx=" & menuidx & vbcrlf
				sqlStr = sqlStr & " order by sorting asc , idx asc "
				rsget.Open sqlStr,dbget,1
				if Not(rsget.EOF or rsget.BOF) Then
					Do Until rsget.eof
			%>
				<div class="swiper-slide">
					<% If rsget("linkurl") <> "" Then %><a href="<%=rsget("linkurl")%>"><% End If %><img src="<%=rsget("slideimg")%>" alt="" /><% If rsget("linkurl") <> "" Then %></a><% End If %>
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
			<button class="slideNav btnPrev">����</button>
			<button class="slideNav btnNext">����</button>
			<div class="mask left"></div>
			<div class="mask right"></div>
		</div>
	</div>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->