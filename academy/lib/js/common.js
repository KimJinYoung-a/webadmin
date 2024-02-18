/*************************************************************
	PageName 	: /academy/lib/js/common.js
	Description : 핑거스 공용 스크립트
	History 	: 2006.11.16 정윤정 생성
	----------------------------------------------------------
	Index		:
		1. fnChkFile : 파일 유효성 확인 (ex. fnChkFile(파일명, 최대용량(MB)))	
*************************************************************/
//----------------------------------------
// 1. fnChkFile 파일 유효성 확인 
//----------------------------------------	
   function fnChkFile(sFile, sMaxSize){   
    //파일 업로드 유무확인
    if (!sFile){
     return true;
    }
    
   	//파일 용량 확인
   	var maxsize = sMaxSize * 1024 * 1024;
   	
   	var img = new Image();
		img.dynsrc = sFile;
		var fSize = img.fileSize ;
		
		if (fSize > maxsize){
			alert("파일크기는 "+sMaxSize+"MB이하만 가능합니다.");
			return false;
		}
		
   	//파일 확장자 확인
   	var pPoint = sFile.lastIndexOf('.');
		var fPoint = sFile.substring(pPoint+1,sFile.length);
		var fExet = fPoint.toLowerCase();

		if (!(fExet == "jpg" || fExet == "gif")) {
			alert("JPG또는 GIF형식의 파일만 가능합니다.");
			return false;
		}
		
		return true;
   }