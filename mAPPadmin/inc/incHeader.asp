<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>텐바이텐 어드민</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<script type="text/javascript" charset="utf-8" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" charset="utf-8" src="/mAppadmin/js/cordova_<%= flgDevice %>.js"></script>
<script type="text/javascript" charset="utf-8" src="/mAppadmin/js/PushNotification.js"></script>
<script type="text/javascript" charset="utf-8" src="/mAppadmin/js/BrowserNavigationBar.js"></script>
<script type="text/javascript" charset="utf-8" src="/mAppadmin/js/index.js"></script>
<script type="text/javascript">
$(function() {
	app.initialize();

	$("#btn-reload").bind("click", function(event, ui) {
		document.location.reload();
	});

	$("#btn-logout").bind("click", function(event, ui) {
		document.location.href="/mAPPadmin/login/doMobileAppLogout.asp";
	});
});
</script>
