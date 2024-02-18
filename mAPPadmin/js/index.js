var pushNotification;

var app = {

	initialize: function() {

		this.bindEvents();

	},

	bindEvents: function() {

		document.addEventListener('deviceready', this.onDeviceReady, false);

	},

	onDeviceReady: function() {

		log("onDeviceReady");
		app.receivedEvent('deviceready');

	},

	receivedEvent: function(id) {

		//alert("onNotificationAPN 000");
		log("receivedEvent");

		pushNotification = window.plugins.pushNotification;
		if (device.platform == 'android' || device.platform == 'Android')
		{
			pushNotification.register(successHandler, errorHandler, {"senderID":"1000979769700", "ecb":"onNotificationGCM"});
		}
		else
		{
			//alert("onNotificationAPN333");
			pushNotification.register(tokenHandler, errorHandler, {"badge":"true", "sound":"true", "alert":"true", "ecb":"onNotificationAPN"});
			//alert("onNotificationAPN444");
		}

		window.plugins.browserNavigationBar.init();

		$('#btn-refresh').click(function() {
			document.location.reload();
		});

		$('#btn-unregister').click(function() {
			pushNotification.unregister(successHandler, errorHandler);
		});
	}
};

function successHandler(result) {
	alert("sss" + result);
	log('result = ' + result)
}

function errorHandler(error) {
	alert("eee" + error);
	log('error = ' + error)
}

function tokenHandler(result) {
	alert("ttt" + result);
	pushNotification.regid = result;
	log('device token = ' + result)
}

// iOS
function onNotificationAPN(event) {

	alert("onNotificationAPN");
	if ( event.url )
	{
		if ( window.confirm(event.alert) )
		{
			window.location = event.url;
		}
	}
	else
	{
		alert(event.alert);
	}

	if ( event.sound ) {
		var snd = new Media(event.sound);
		snd.play();
	}

	alert("onNotificationAPN22");
	if ( event.badge ) {
		pushNotification.setApplicationIconBadgeNumber(successHandler, errorHandler, event.badge);
	}
}

// Android
function onNotificationGCM(e) {

	alert("aaa")
	switch ( e.event ) {

	case 'registered':
		if ( e.regid.length > 0 ) {
			pushNotification.regid = e.regid;
			log("regID = " + e.regid);
		}
		break;

	case 'message':

		if ( e.foreground )
		{
			//var my_media = new Media("/android_asset/www/" + e.soundname);
			//my_media.play();
		}
		else
		{	// otherwise we were launched because the user touched a notification in the notification tray.
			if ( e.coldstart )
				log('COLDSTART NOTIFICATION');
			else
				log('BACKGROUND NOTIFICATION');
		}

		if ( e.payload.url.length > 0 )
		{
			if ( window.confirm(e.payload.message) )
			{
				window.location = e.payload.url;
			}
		}
		else
		{
			alert(e.payload.message);
			// alert(e.payload.message + " 한글");
			// alert("111 한글");
		}

		break;

	case 'error':
		log('ERROR MSG:' + e.msg);
		break;

	default:
		log('EVENT -> Unknown, an event was received and we do not know what it is');
		break;

	}
}

window.log = function(message) {
	$('#log').append("<li>" + message + "</li>");
}
