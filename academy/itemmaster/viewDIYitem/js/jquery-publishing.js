$(document).ready(function() {

	// news ticker
	newsSwiper= new Swiper('.newsList .swiper-container',{
		autoplay:3000,
		loop: true,
		pagination:false,
		mode: 'vertical',
		noSwiping:true
	});

	$('.newsList .btnPrev').on('click', function(e){
		e.preventDefault()
		newsSwiper.swipePrev()
	});
	$('.newsList .btnNext').on('click', function(e){
		e.preventDefault()
		newsSwiper.swipeNext()
	});

	// unb
	$(".unbV16 li").mouseover(function(){
		$(".unbV16 li").removeClass("subUnbOff");
		$(this).addClass("subUnbOn");
	});
	$(".unbV16 li").mouseleave(function(){
		$(this).removeClass("subUnbOn");
	});

	// gnb
	$(".gnbV16 > ul > li").mouseover(function(){
		$(".gnbV16 > ul > li").removeClass("menuOn");
		$(this).addClass("menuOn");
	});
	$(".gnbV16 > ul > li").mouseleave(function(){
		$(".gnbV16 > ul > li").removeClass("menuOn");
	});


	$(".lnb li").each(function(){
		var link = $(this).children("a");
		var image = $(link).children("img");
		var imgsrc = $(image).attr("src");

		// add mouseover
		$(link).mouseover(function(){
			var on = imgsrc.replace(/_off.gif$/gi,"_on.gif");
			$(image).attr("src",on);
		});
		// add mouseover
		$(link).focus(function(){
			var on = imgsrc.replace(/_off.gif$/gi,"_on.gif");
			$(image).attr("src",on);
		});

		// add mouse out
		$(link).mouseout(function(){
			$(image).attr("src",imgsrc);
		});
		// add mouse out
		$(link).blur(function(){
			$(image).attr("src",imgsrc);
		});
	});


	$('.qnaAdd .qArea').click(function() {
		if($(this).next().is(':hidden') == true) {
			$('.qnaAdd .aArea').hide();
			$(this).next().show();
		} else {
			$(this).next().hide();
		}
	});
	$('.qnaAdd .aArea').hide();


	$(".tabWrap li").each(function(){
		var link = $(this).children("a");
		var image = $(link).children("img");
		var imgsrc = $(image).attr("src");

		if($(this).hasClass("tabOn")) {
			var on = imgsrc.replace(/_off.gif$/gi,"_on.gif");
			$(image).attr("src",on);
		}
	});

	/*
	$(".tabWrap2 li").each(function(){
		var link = $(this).children("a");
		var image = $(link).children("img");
		var imgsrc = $(image).attr("src");

		if($(this).hasClass("tabOn")) {
			var on = imgsrc.replace(/_off.gif$/gi,"_on.gif");
			$(image).attr("src",on);
		}
	});
	*/

	$('.photoList li').mouseover(function(){
		$('.photoList li').removeClass("photoOn");
		$(this).addClass("photoOn");
		$('.photoView').hide();
		$("p[class='photoView'][id='"+'B'+$(this).attr("id")+"']").show();
	});

	$('.scrapDo').mouseover(function(){
		$('.scrapView').show();
	});
	$('.scrapDo').mouseleave(function(){
		$('.scrapView').hide();
	});
	$('.scrapView').hide();


	$('.openLyr .titArea').click(function() {
		if($(this).next().is(':hidden') == true) {
			$('.openLyr .viewArea').hide();
			$(this).next().show();
		} else {
			$(this).next().hide();
		}
	});
	$('.openLyr .viewArea').hide();

	$('.btnQnaWrite').click(function(){
		$('.questionWrite').css('display', 'block');
		return false;
	});

	$('.btnQnaCnacel').click(function(){
		$('.questionWrite').css('display', 'none');
		return false;
	});

	$('.btnQnaWrite').click(function(){
		$('.questionWrite').css('display', 'block');
		return false;
	});


	$('.lyrCloseBtn').click(function(){
		$('.lyrWrap').css('display', 'none');
		return false;
	});

	$('.lyrCloseBtn2').click(function(){
		$('.lyrWrap').css('display', 'block');
		return false;
	});

	$(".cash").click(function(){
		$("#cashDiv").toggle();
	});

	$(".eleIns").click(function(){
		$("#eleInsDiv").toggle();
	});

	$('.howView > li').click(function(){
		$('.howView > li').removeClass('selected');
		$(this).addClass('selected');
		return false;
	});

//	$('.lnbDiyshop li').find('.subLnbWrap').hide();
//	$('.lnbDiyshop .first').find('.subLnbWrap').show();
//	$('.lnbDiyshop li > span').click(function(){
//		$('.lnbDiyshop li > .subLnbWrap:visible').parents("li").addClass('hide').removeClass('show');
//		$('.lnbDiyshop li > .subLnbWrap:visible').slideUp(100);
//		$(this).parents("li").removeClass('hide').addClass('show');
//		$(this).parents("li").find('.subLnbWrap').slideDown(100);
//	});

	$('.hotKey dd li:last-child').css('border-bottom', 'none');
	$('.mainBnr02 .brickWrap ul li:first-child').css('border-right', '1px solid #e2e2e2');

	var $items1 = $('.bestBrand .callTabs li');
	$items1.mouseover(function() {
		$items1.removeClass('selected');
		$(this).addClass('selected');
		var index1 = $items1.index($(this));
		$('.rollingView .viewBnr').hide().eq(index1).fadeIn();
	}).eq(0).mouseover();

	var $items2 = $('.todayPick .callTab li');
	$items2.click(function() {
		$items2.removeClass('selected');
		$(this).addClass('selected');
		var index2 = $items2.index($(this));
		$('.rollingView .view').hide().eq(index2).fadeIn();
	}).eq(0).click();

	var $items3 = $('.tabRolling .callTab li');
	$items3.click(function() {
		$items3.removeClass('selected');
		$(this).addClass('selected');
		var index3 = $items3.index($(this));
		$('.rollingView .viewImg').hide().eq(index3).fadeIn();
	}).eq(0).click();

//	lecture Pick Detail Layer View
	$('.pickList li').mouseover(function(){
		$(this).find('.lyrDetail').show();
		$(this).css('z-index', '1000');
		$(this).parent().parent().css('z-index', '1000');
		$(this).children('img').css('opacity', '0.5');
		$(this).children('img').css('filter', 'alpha(opacity=50)');
	});
	$('.pickList li').mouseout(function(){
		$(this).find('.lyrDetail').hide();
		$(this).css('z-index', '90');
		$(this).parent().parent().css('z-index', '89');
		$(this).children('img').css('opacity', '1');
		$(this).children('img').css('filter', 'alpha(opacity=100)');
	});

//	Video LNB
//	$('.lnbVideo li').find('.lnbContView').hide();
//	$('.lnbVideo li.videoAll').find('.lnbContView').show();
//	$('.lnbVideo li > span').click(function(){
//		$('.lnbVideo li > .lnbContView:visible').parent("li").addClass('lnbOff').removeClass('lnbOn');
//		$('.lnbVideo li > .lnbContView:visible').hide();
//		$(this).parent("li").removeClass('lnbOff').addClass('lnbOn');
//		$(this).parent("li").find('.lnbContView').show();
//	});

//	Lecture Calendar Selectbox */
	$(".allViewSelect dt").click(function(){
		if($(".allViewSelect dd").is(":hidden")){
			$(this).parent('.allViewSelect').children('dd').slideDown(100);
			$(this).addClass("selected");
		}else{
			$(this).parent('.allViewSelect').children('dd').slideUp(100);
			$(this).removeClass("selected");
		};
	});
	$(".allViewSelect dd li a").click(function(){
		$(".allViewSelect dd li a").removeClass("on");
		$(this).addClass("on");
		$(this).parent().parent().parent().parent('.allViewSelect').children('dt').children('span').text($(this).text());
		$(this).parent().parent().parent('dd').slideUp(100);
		$(this).parent().parent().parent().parent('.allViewSelect').children('dt').removeClass("selected");
	});
	$(".allViewSelect dd").mouseleave(function(){
		$(this).hide();
		$(this).parent('.allViewSelect').children('dt').removeClass("selected");
	});

//	Lecture Calendar Year, Month Layer View
	$(".scheduleCal dt div div").click(function(){
		$(this).addClass("monthSelect");
		$(".scheduleCal .lyrWrap").show();
	});

	$(".thisMonth .lyrCloseBtn").click(function(){
		$(".scheduleCal dt div div").removeClass("monthSelect");
		$('.lyrWrap').css('display', 'none');
		return false;
	});

//	$(".scheduleCal dt div div").mouseleave(function(){
//		$(this).removeClass("monthSelect");
//		$(".scheduleCal .lyrWrap").hide();
//	});

//	Lecture Calendar Detail Layer View
	$('.classList li').mouseover(function(){
		$(this).find('.lyrDetail').show();
		$(this).css('z-index', '1000');
		$(this).parent().parent().css('z-index', '1000');
	});
	$('.classList li').mouseout(function(){
		$(this).find('.lyrDetail').hide();
		$(this).css('z-index', '90');
		$(this).parent().parent().css('z-index', '89');
	});

//	Good Teacher activity more view (8/11 Ãß°¡)
	$('.teacherView .btnMoreView').click(function(){
		$('.teacherView .actView').toggleClass('extend');
		$(this).toggleClass('fold');
		return false;
	});

});