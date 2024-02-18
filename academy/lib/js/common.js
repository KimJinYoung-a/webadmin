/*************************************************************
	PageName 	: /academy/lib/js/common.js
	Description : �ΰŽ� ���� ��ũ��Ʈ
	History 	: 2006.11.16 ������ ����
	----------------------------------------------------------
	Index		:
		1. fnChkFile : ���� ��ȿ�� Ȯ�� (ex. fnChkFile(���ϸ�, �ִ�뷮(MB)))	
*************************************************************/
//----------------------------------------
// 1. fnChkFile ���� ��ȿ�� Ȯ�� 
//----------------------------------------	
   function fnChkFile(sFile, sMaxSize){   
    //���� ���ε� ����Ȯ��
    if (!sFile){
     return true;
    }
    
   	//���� �뷮 Ȯ��
   	var maxsize = sMaxSize * 1024 * 1024;
   	
   	var img = new Image();
		img.dynsrc = sFile;
		var fSize = img.fileSize ;
		
		if (fSize > maxsize){
			alert("����ũ��� "+sMaxSize+"MB���ϸ� �����մϴ�.");
			return false;
		}
		
   	//���� Ȯ���� Ȯ��
   	var pPoint = sFile.lastIndexOf('.');
		var fPoint = sFile.substring(pPoint+1,sFile.length);
		var fExet = fPoint.toLowerCase();

		if (!(fExet == "jpg" || fExet == "gif")) {
			alert("JPG�Ǵ� GIF������ ���ϸ� �����մϴ�.");
			return false;
		}
		
		return true;
   }