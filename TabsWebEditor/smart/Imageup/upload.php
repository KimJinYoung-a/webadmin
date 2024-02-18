<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>파일 업로드</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <script type="text/javascript" src="../../fckeditor/prototype.js"></script>
    <script type="text/javascript" src="../../fckeditor/imageup.js"></script> 
</head>

<body>
	<script type="text/javascript">
	<?php
	$uploaddir = realpath("../Temp/");

	$srcfilename = $_FILES['uploadFile']['name'];
	// 파일 확장자 구하기
	$fileext = substr(strrchr($srcfilename, '.'), 0);
	$filesize = $_FILES['uploadFile']['size'];

	// Temp에 저장될 임시 파일명 생성
	$tmpfilename = uniqid('tmp', true) . $fileext;
	$savefile = $uploaddir . $tmpfilename;

	if (move_uploaded_file($_FILES['uploadFile']['tmp_name'], $savefile)) {
		printf('onCompleteUpload("%s", "%s", "%d");', $tmpfilename, $srcfilename, $filesize);
	}
	?>	
	</script>
</body>
</html>




