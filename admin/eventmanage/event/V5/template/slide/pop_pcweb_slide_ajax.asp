<%@ language=vbscript %>
<% option explicit %>
<% Response.charset = "euc-kr"
'###############################################
' PageName : pop_pcweb_slide_ajax.asp
' Discription : PCWEB slide ajax
' History : 2016-02-16 이종화
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

	If eCode <> "" Then 
		strSql = "SELECT topimg , topaddimg , btmYN , btmimg , btmcode , btmaddimg , pcadd1 " & vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event_slide_template where evt_code = '"& eCode &"'" & vbcrlf 
		strSql = strSql & " and device = 'W' and menuidx=" & menuidx & vbcrlf 
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			topimg		= rsget("topimg")
			topaddimg	= rsget("topaddimg")
			btmYN		= rsget("btmYN")
			btmimg		= rsget("btmimg")
			btmcode		= rsget("btmcode")
			btmaddimg	= rsget("btmaddimg")
			pcadd1		= rsget("pcadd1")
		End If
		rsget.close()
	End If

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
<div class="selectType">
	<span><input type="radio" id="sType01" <%=chkiif(gubun = 1 ,"checked","")%> name="wgubun" onclick="dfslide('1');"/> <label for="sType01">와이드 슬라이드</label></span>
	<span class="lMar05"><input type="radio" id="sType02" <%=chkiif(gubun = 2 ,"checked","")%> name="wgubun" onclick="dfslide('2');"/> <label for="sType02">와이드+풀단 슬라이드</label></span>
	<span class="lMar05"><input type="radio" id="sType03" <%=chkiif(gubun = 3 ,"checked","")%> name="wgubun" onclick="dfslide('3');"/> <label for="sType03">풀단 슬라이드</label></span>
</div>

<div class="<%=gubuncls%>">
	<div class="evtSection evtTop" <% If topaddimg <>"" Then %>style="background-image:url(<%=topaddimg%>);"<% End If %> id="spantopaddimg_bg">
		<div id="spantopimg">
			<%IF topimg <> "" THEN %>
			<img src="<%=topimg%>" alt="" />
			<a href="javascript:jsDelImg('topimg','spantopimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
			<%END IF%>
		</div>
	</div>

	<div class="evtSection swiper">
		<div class="swiper-container" <% If pcadd1 <>"" Then %>style="background-image:url(<%=pcadd1%>);"<% End If %> id="spanpcadd1_bg">
			<div class="swiper-wrapper">
			<% 
				If eCode <> "" Then 

				sqlStr = "SELECT slideimg , linkurl , bgimg , sorting " & vbcrlf
				sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"' " & vbcrlf
				sqlStr = sqlStr & " and isusing = 'Y' and device ='W' and menuidx=" & menuidx & vbcrlf
				sqlStr = sqlStr & " order by sorting asc , idx asc "
				rsget.Open sqlStr,dbget,1
				if Not(rsget.EOF or rsget.BOF) Then
					Do Until rsget.eof
			%>
				<div class="swiper-slide" <% If rsget("bgimg") <>"" Then %>style="background-image:url(<%=rsget("bgimg")%>);"<% End If %>>
					<% If rsget("linkurl") <> "" Then %><a href="http://www.10x10.co.kr<%=rsget("linkurl")%>" target="_blank"><% End If %><img src="<%=rsget("slideimg")%>" alt="" /><% If rsget("linkurl") <> "" Then %></a><% End If %>
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

	<div class="evtSection evtBtm" <% If btmaddimg <>"" Then %>style="background-image:url(<%=btmaddimg%>); background-position-y:0;"<% End If %> id="spanbtmaddimg_bg">
		<div id="spanbtmimg">
		<% If btmYN = "Y" Then %>
			<%IF btmimg <> "" THEN %>
			<img src="<%=btmimg%>" alt="" />
			<a href="javascript:jsDelImg('btmimg','spanbtmimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
			<%END IF%>
		<% Else %>
			<%=db2html(btmcode)%>
		<% End If %>
		</div>
	</div>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->