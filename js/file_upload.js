/*==============================================================*/
/*           �ǽ����ε� ��Ƽ��X ��Ʈ�� �Լ� start               */
/*==============================================================*/


// ���� ���� ��ȭ���ڸ� �����ؼ� ���ε��� ������ �����Ѵ�.
function addFiles()  
{
    frm_upload.TABSFileup.AddFile();
}

// ������ ������ ������ �信�� �����Ѵ�.
function removeFiles()  
{
    frm_upload.TABSFileup.RemoveFile();
}

// ���ε��� ���� ����� �����Ѵ�.
function listFiles()  
{
	var UploadFiles = frm_upload.TABSFileup.UploadFiles;
	var i;
	for (i = 1; i <= UploadFiles.Count; i++) {
		alert("���� �̸�: " + UploadFiles(i).SourceFile + "\n\n���� ũ��: " + UploadFiles(i).FileSize + " ����Ʈ");
	}
}

// submitForm()�� ����� �߻��ϴ� �̺�Ʈ
function OnCompletedPostMultipartFormData(ErrType, ErrCode, ErrText, retURL)
{
	if (ErrType == 0) {
		// ������ �߻����� ���� ��� ���� �������� ��ȯ�Ѵ�.
		location.href = retURL
		// alert("������:" + frm_upload.TABSFileup.Response);
		
    } else {
		// ���� ������ �����ϰ� ����Ѵ�.		
		alert("��������:" + ErrType + "   �����ڵ�:" + ErrCode + "   ��������:" + ErrText);
		alert("������:" + frm_upload.TABSFileup.Response);
    }		
		
}

// ������ �� ����� ���ʷ� �����Ѵ�.
function changeViewStyle()
{
	var ViewStyle = frm_upload.TABSFileup.ViewStyle;
	if (++ViewStyle > 3) ViewStyle = 0;
	frm_upload.TABSFileup.ViewStyle = ViewStyle;
}


// ���۳�Ʈ �κ��� //
function TabsEmbed(md,fid,wd,ht,tg,ft,vs,bg)
{
	switch(md)
	{
		case "view" :
			document.write('<OBJECT ID="' + fid + '" width=' + wd + ' height=' + ht + ' CLASSID="CLSID:BCFA4759-1193-4EC3-92A0-F03F6461DA78"  CODEBASE="/lib/util/Tabs/TABSFileup.cab#version=2,1,0,15">');
			document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			document.write('<PARAM NAME="UploadStyle" VALUE="1">');
			document.write('<PARAM NAME="Mode" VALUE="3">');
			document.write('<PARAM NAME="Key" VALUE="D5C0D007C92B0E978C352FD326461664398A221751847865">');
			document.write('<PARAM NAME="BkColor" VALUE="' + bg + '"> ');
			document.write('</OBJECT>');
			break;

		case "write" :
			document.write('<OBJECT ID="' + fid + '" width="' + wd + '" height="' + ht + '" border=0 CLASSID="CLSID:BCFA4759-1193-4EC3-92A0-F03F6461DA78" CODEBASE="/lib/util/Tabs/TABSFileup.cab#version=2,1,0,15">');
			document.write('<PARAM NAME="UploadURL" VALUE="' + tg + '">');
			document.write('<PARAM NAME="CodePage" VALUE="949">');
			document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			document.write('<PARAM NAME="UploadStyle" VALUE="1">');
			document.write('<PARAM NAME="BkColor" VALUE="' + bg + '">');
			document.write('<PARAM NAME="Key" VALUE="D5C0D007C92B0E978C352FD326461664398A221751847865">');
			document.write('<PARAM NAME="Mode" VALUE="1">');
			document.write('<PARAM NAME="FileFilter" VALUE="' + ft + '">');
			document.write('</OBJECT>');
			break;

		case "modi" :
			document.write('<OBJECT ID="' + fid + '" width="' + wd + '" height="' + ht + '" border=0 CLASSID="CLSID:BCFA4759-1193-4EC3-92A0-F03F6461DA78" CODEBASE="/util/Tabs/TABSFileup.cab#version=2,1,0,15">');
			document.write('<PARAM NAME="UploadURL" VALUE="' + tg + '">');
			document.write('<PARAM NAME="CodePage" VALUE="949">');
			document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			document.write('<PARAM NAME="UploadStyle" VALUE="1">');
			document.write('<PARAM NAME="BkColor" VALUE="' + bg + '">');
			document.write('<PARAM NAME="Key" VALUE="D5C0D007C92B0E978C352FD326461664398A221751847865">');
			document.write('<PARAM NAME="Mode" VALUE="2">');
			document.write('<PARAM NAME="FileFilter" VALUE="' + ft + '">');
			document.write('</OBJECT>');

			//�ǽ�4
			//document.write('<OBJECT ID="' + fid + '" width="' + wd + '" height="' + ht + '" border=0 CLASSID="CLSID:2342E134-C396-43EC-BCB8-13D513BC5FE0" CODEBASE="/util/Tabs/tabsfileup4setup.cab">');
			//document.write('<PARAM NAME="Mode" VALUE="edit">');
			//document.write('<PARAM NAME="licensekey" VALUE="RleQfLjGMBxH4IHssrGkVJAINQYqF548drZ+Y4Hi+vDrB4mmdrYXHw==">');
			//document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			//document.write('<PARAM NAME="CodePage" VALUE="949">');
			//document.write('<PARAM NAME="UploadURL" VALUE="' + tg + '">');
			//document.write('<PARAM NAME="FileFilter" VALUE="' + ft + '">');
			//document.write('<PARAM NAME="BkColor" VALUE="' + bg + '">');
			//document.write('</OBJECT>');
			break;
	}
}

function OnChangingUploadFile(TotalCount, TotalFileSize)
{
	var UploadFiles = frm_upload.TABSFileup.UploadFiles;
	var i, ft;

    if (TotalCount > 10) {
        alert("�ʹ� ���� ������ �����ϼ̽��ϴ�.\n\n�� ������ 10������ ���ÿ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
        frm_upload.TABSFileup.StopUpload = true;
    }
	else if ((TotalFileSize/TotalCount) > (2 * 1024 * 1024))
	{
        alert("�ʹ� �뷮�� ū������ �����ϼ̽��ϴ�.\n\n�� ��� 2MB ������ �뷮���� ���������Ͽ� �ֽʽÿ�.");
        frm_upload.TABSFileup.StopUpload = true;
	}
}
