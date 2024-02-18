$(function() {
	//GNB Control
	$('.gnb li').mouseover(function() {
		$('.gnb li').removeClass('subOn');
		$(this).addClass('subOn');
		var subW = $(this).children('.subNavi').outerWidth();
		$('.subNavi').css('margin-left',-subW/2+'px');
	});
	$('.gnb li').mouseout(function() {
		$('.gnb li').removeClass('subOn');
	});
//	$('.gnb li:last-child .subNavi').css('right', '-1px');

	//LNB toggle Control
	$('.toggle').click(function() {
		$(this).parent().toggleClass('lnbView');
		if($(this).children('span').html()=="�ݱ�") {
			$(this).children('span').html('����');
		} else {
			$(this).children('span').html('�ݱ�');
		}
		if($('.container').hasClass('lnbView')) {
			$('.wrap, .contSection').css('min-width', '960px');
		} else {
			$('.wrap, .contSection').css('min-width', '1140px');
		}
	});

	//LNB Control
	$('.lnb dl dt').click(function() {
		$('.lnb dl').removeClass('current');
		$(this).parent().addClass('current');
		$('.lnb dl dd').hide();
		$(this).parent().children('dd').show();
	});

	//List table sorting Control
	$('.sorting').click(function() {
		$(this).children('span').toggleClass('sortWay');
	});

	//help layer Control
	$('.helpBox').mouseover(function(){
		$(this).children('dl').children('dd').show();
	});
	$('.helpBox').mouseleave(function(){
		$(this).children('dl').children('dd').hide();
	});
});