<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width" />
<title></title>
<link rel="stylesheet" type="text/css" href="/admin/eventmanage/event/v5/lib/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/admin/eventmanage/event/v5/lib/css/adminCommon.css" />
<link rel="stylesheet" href="https://cdn.materialdesignicons.com/3.6.95/css/materialdesignicons.min.css">
<style type="text/css">
html {overflow:auto;}
body {background-color:#f4f4f4;}
</style>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script>
$( function() {
	// datepicker
	// https://jqueryui.com/datepicker/#default
	$('#datepicker1').datepicker({
		inline: true,
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd',
	});
	$('#datepicker2').datepicker({
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd'
	});
	$('#datepicker3').datepicker({
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd'
	});
	$('#datepicker5').datepicker({
		inline: true,
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd',
	});
	$('#datepicker6').datepicker({
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd'
	});
	$('#datepicker7').datepicker({
		inline: true,
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd',
	});
	$('#datepicker8').datepicker({
		showOtherMonths: true,
		showMonthAfterYear: true,
		monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		dateFormat: 'yy-mm-dd'
	});
	$('.mdi-calendar-month').click(function() {
		$(this).siblings(".hasDatepicker").focus();
	});
} );
</script>
</head>
<body>
<!-- 팝업 사이즈 : 1024*000 -->