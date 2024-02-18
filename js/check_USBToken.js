function checkUSBKey ()
{
	var sn='';
	// ActiveX 설치유무 검사
	try {sn = MaGerAuth.GetSN();} catch(e) {alert('프로그램을 설치하여 주세요');}
	// USB Token 확인 없으면 로그아웃
	if(sn=='') {
		top.location = '/login/usbNotFound.asp';
	}

	setTimeout("checkUSBKey()",10 * 60 * 1000);   //10분마다 재실행(60초 -> 10분, 2014-07-17, skyer9)
}
