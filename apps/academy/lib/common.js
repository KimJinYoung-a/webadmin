$(function(){
	// lnb영역 스크롤링
	var lnbSwiper = new Swiper('.lnb .swiper-container', {
		//scrollContainer:true,
		scrollbar:'.lnb .swiper-scrollbar',
		direction:'vertical',
		slidesPerView: 'auto',
		mousewheelControl: true,
		freeMode: true
	});

	// gnb
	$(".headA .head h1").click(function(){
		if($(".gnb").is(":hidden")){
			$(".layerMask").fadeIn();
			$(".container").addClass('openGnb');
		} else {
			$(".layerMask").fadeOut();
			$(".container").removeClass('openGnb');
		}
	});
	$(".layerMask").click(function(){
		$(this).fadeOut(function(){
			lnbSwiper.slideTo(0);
			$(".myFingersList li ul").hide();
		});
		$(".container").removeClass('openLnb');
		$(".container").removeClass('openGnb');
		$(".layerPopup").fadeOut();
	});

	// lnb
	$(".head .btnMypage").click(function(){
		$(".layerMask").fadeIn(function(){
			lnbSwiper.update();
			jsGetLnbPrderCount();
		});
		$(".container").addClass('openLnb');
		$(".container").removeClass('openGnb');
	});

	// footer
	$("footer h1").click(function(){
		if($(".footInfo").is(":hidden")){
			$(this).addClass('show');
		} else {
			$(this).removeClass('show');
		}
		$(".footInfo").toggle();
		$('html, body').animate({scrollTop:$(document).height()}, 'slow');
	});

	// go top
	$('#btnGotop').hide();
	$(window).scroll(function(){
		var vSpos = $(window).scrollTop();
		var docuH = $(document).height() - $(window).height();
		if (vSpos > 180){
			if($('#btnGotop').css("display")=="none"){
				$('#btnGotop').fadeIn();
			}
		} else {
			$('#btnGotop').hide();
		}
	});
	$('#btnGotop').click(function(){
		$('html, body').animate({scrollTop:0}, 'fast');
	});

	// tab
	$(".fingerTab .tabCont").hide();
	$(".fingerTab .tabContainer").find(".tabCont:first").show();
	$(".fingerTab .tabNav li").click(function() {
		$(this).siblings("li").removeClass("current");
		$(this).addClass("current");
		$(this).closest(".fingerTab .tabNav").nextAll(".fingerTab .tabContainer:first").find(".tabCont").hide();
		var activeTab = $(this).attr("name");
		$(".tabCont[id|='"+ activeTab +"']").show();
	});

	$(".moreInfo > dt").click(function(){
		$(this).next("dd").toggle();
		$(this).closest(".moreInfo").toggleClass("current");
	});

});

function closeLayer(){
	$(".layerPopup").fadeOut();
	$(".layerMask").fadeOut();
	$("#hBoxes").remove();
}