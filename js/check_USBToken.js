function checkUSBKey ()
{
	var sn='';
	// ActiveX ��ġ���� �˻�
	try {sn = MaGerAuth.GetSN();} catch(e) {alert('���α׷��� ��ġ�Ͽ� �ּ���');}
	// USB Token Ȯ�� ������ �α׾ƿ�
	if(sn=='') {
		top.location = '/login/usbNotFound.asp';
	}

	setTimeout("checkUSBKey()",10 * 60 * 1000);   //10�и��� �����(60�� -> 10��, 2014-07-17, skyer9)
}
