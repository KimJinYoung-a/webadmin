<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹ� ��� ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim uploadform, objfile, sDefaultPath, sFolderPath ,orderCsGbn ,monthFolder
Dim iML, sFile, sFilePath, SellSite, iMaxLen, sUploadPath, orgFileName, maybeSheetName
Dim overseasPrice, overseasDeliveryPrice, overseasRealPrice, reserve01, beasongNum11st, outmalloptionno
Dim outMallGoodsNo,shoplinkermallname,shoplinkerPrdCode,shoplinkerOrderID,shoplinkerMallID
Dim tmpOverseasRealprice, sItemname
Dim tmpStr1, tmpStr2, loops
dim tmpVal, tmpItem
Dim isValid

monthFolder = Replace(Left(CStr(now()),7),"-","")

IF (application("Svr_Info")	= "Dev") then
    'Set uploadform = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
	'2019-10-11 15:05 ������ TABSUpload4.Upload�� ����
	Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
ELSE
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

Set objfile	   = Server.CreateObject("Scripting.FileSystemObject")
''sDefaultPath   = Server.MapPath("\admin\etc\orderInput\upFiles\")
sDefaultPath   = Server.MapPath("/admin/etc/orderInput/upFiles/")
uploadform.Start sDefaultPath '���ε���

iMaxLen 		= uploadform.Form("iML")	'�̹�������ũ��
SellSite 	= uploadform.Form("sellsite")

IF (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) THEN	'����üũ

    '���� ����
    sFolderPath = sDefaultPath&"/"&sellsite&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    sFolderPath = sDefaultPath&"/"&sellsite&"/"&monthFolder&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    '��������
	sFile = fnMakeFileName(uploadform("sFile"))
	sFilePath = sFolderPath&sFile
	sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

	orgFileName = uploadform("sFile").FileName
	maybeSheetName = Replace(orgFileName,"."&uploadform("sFile").FileType,"")
END IF

Set objfile		= Nothing
Set uploadform = Nothing

Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i,j, k, m
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24

    ''����Ŀ
    ''2013-11-27 10:45 ������ ����.. 11,-1,22,27,31 => 11,-1,22,27,27�� ��������18446�� ���û���
    if (SellSite="shoplinker") then
'        xlPosArr = Array(3,2,2,-1,-1,   33,34,35,-1,37,   38,39,40,41,41,   11,-1,22,27,31,   26,15,20,10,-1,   -1,42,-1,13,-1,	  6,9,47,-1,12)
'        xlPosArr = Array(3,2,2,-1,-1,   33,34,35,-1,37,   38,39,40,41,41,   11,-1,22,27,27,   26,15,20,10,-1,   -1,42,-1,13,-1,	  6,9,47,7,12)		'2015-07-03������ ��ȣ
        xlPosArr = Array(3,2,2,-1,-1,   36,37,38,-1,40,   41,42,43,44,44,   11,-1,25,30,30,   29,15,20,10,-1,   -1,45,-1,13,-1,	  6,9,50,7,12)
        ArrayLen = UBound(xlPosArr)
	    skipString = "No."
	    afile = sFilePath
	    aSheetName = ""

	'/hmall
    elseif (SellSite="hmall1010") then
		xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "Sheet1" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/��ؼ� - ������
    elseif (SellSite="dnshop") then
'	    xlPosArr = Array(2,41,41,-1,-1,  18,20,21,-1,19,   20,21,22,23,24,   37,-1,9,27,27,   -1,5,6,4,-1,   3,15,17,-1,-1, -1)
	    xlPosArr = Array(2,38,38,-1,-1,  15,17,18,-1,16,   17,18,19,20,21,   34,-1,9,24,24,   -1,5,6,4,-1,   3,13,14,-1,-1, -1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/�ؿ��߱�����Ʈ
	elseif (SellSite="cn10x10") then
	    'xlPosArr = Array(0,1,3,-1,-1,   37,38,38,39,37,   38,38,40,41,41,   9,7,15,16,16,   20,-1,-1,9,7,   -1,42,-1,28,-1,	  -1,-1,-1,36,-1)
	    xlPosArr = Array(0,1,3,-1,-1,   37,38,38,39,37,   38,38,40,41,41,   9,7,15,16,16,   20,-1,-1,9,7,   -1,-1,-1,28,-1,	  -1,-1,-1,36,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Worksheet"  '' sheet name Maybe filename in

	'/�ؿ��߱�����Ʈ
	elseif (SellSite="cnglob10x10") then
		'xlPosArr = Array(0,13,49,-1,-1,   14,19,19,24,22,   29,30,31,33,33,   92,-1,7,11,12,   -1,68,6,1,91,   -1,40,21,61,12,	  89,-1,-1,35,-1)
		xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1,	44,62,12,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "combine_data"  '' sheet name Maybe filename in

	elseif (SellSite="cnhigo") then
		'xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1)
		xlPosArr = Array(0,20,21,-1,-1,   11,12,12,-1,13,   14,15,19,17,17,   6,8,5,1,2,   -1,7,9,-1,-1,   -1,-1,-1,3,4,	  -1,-1,-1,18,-1,	1,3,2,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	elseif (SellSite="celectory") then
	    'xlPosArr = Array(5,32,-1,-1,-1,   21,24,25,-1,21,    24,25,22,23,23,   -1,-1,14,16,16,   -1,12,13,-1,-1,	 6,26,-1,18,-1)
	    xlPosArr = Array(3,2,-1,-1,-1,   11,12,12,-1,13,    14,14,15,16,16,   5,-1,10,9,9,   -1,7,8,4,-1,	 -1,17,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in


	elseif (SellSite="cnugoshop") then
		'xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1)
		xlPosArr = Array(1,21,22,-1,-1,   12,13,13,-1,14,   15,16,20,18,18,   7,9,6,2,3,   -1,8,10,-1,-1,   -1,-1,-1,4,5,	  -1,-1,-1,19,-1,	2,4,3,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/11���� �����̽þ�
	elseif (SellSite="11stmy") then
		'xlPosArr = Array(2,3,5,-1,-1,   26,13,13,-1,12,   13,13,15,14,14,   8,-1,10,20,46,   -1,9,11,7,-1,   4,17,-1,23,18,	  -1,-1,-1,16,-1,	20,23,20,6,-1)
		xlPosArr = Array(0,2,3,-1,-1,   15,16,16,-1,15,   16,16,17,19,19,   9,-1,6,25,32,   -1,4,5,8,-1,   1,20,-1,24,22,	  -1,-1,-1,18,-1,	25,24,25,13,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24

	elseif (SellSite="zilingo") then
		'xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1)
		xlPosArr = Array(1,3,3,-1,-1,   17,21,21,-1,17,   21,21,20,19,19,   -1,-1,8,11,11,   -1,7,-1,0,-1,   -1,-1,-1,-1,-1,	  -1,-1,-1,18,-1,	9,-1,9,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "Sheet0" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in


	'//�ؿ� etsy
	elseif (SellSite="etsy") then
		'xlPosArr = Array(2,3,5,-1,-1,   26,13,13,-1,12,   13,13,15,14,14,   8,-1,10,20,46,   -1,9,11,7,-1,   4,17,-1,23,18,	  -1,-1,-1,16,-1,	20,23,20,6,-1)
		xlPosArr = Array(26,16,16,-1,-1,   2,-1,-1,-1,18,   -1,-1,23,19,21,   -1,-1,3,4,7,   -1,1,-1,15,-1,   14,-1,-1,9,12,	  -1,-1,-1,25,-1,	4,9,4,22,24)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

    '/�Ե����� - ��ۺ� ���� 30,000 �̻� ������ // ������ ��ȭ �ڵ��� ����.. // 37 �ֹ� ��ǰ ����(�ֹ�,��ȯ�ֹ�)
    elseif (SellSite="lotteCom") then
        xlPosArr = Array(2,0,-1,-1,-1,   15,-1,-1,-1,10,    13,14,11,12,12,   -1,-1,32,30,30,   29,5,27,4,-1    ,3,17,26,-1,42    ,-1,-1,37) ''29Col
                   ''Array(0,-1,-1,-1,-1,26,28,29,-1,30,31,32,33,34,34,42,43,5,9,-1,11,3,4,42,43,-1,35,38,20)
	    ArrayLen = UBound(xlPosArr)
	    skipString="���������"
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

    elseif (SellSite="lotteon") then
        xlPosArr = Array(1,0,-1,-1,-1,   10,9,9,-1,10,    11,11,13,12,12,   -1,-1,45,44,44,   -1,33,34,35,37    ,17,52,40,-1,-1    ,-1,-1,-1) ''29Col
	    ArrayLen = UBound(xlPosArr)
	    skipString="���������"
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in


    '/�Ե����̸�
    elseif (SellSite="lotteimall") then
       'xlPosArr = Array(5,0,-1,-1,-1,  38,39,40,-1,24,   30,31,25,26,26,   14,-1,18,17,17,   -1,10,12,8,-1,	 -1,36,-1,-1,-1)
       'xlPosArr = Array(6,0,-1,-1,-1,  40,41,42,-1,26,   32,33,27,28,28,   16,-1,20,19,19,   -1,12,14,10,-1,	 -1,38,-1,-1,-1)
	   xlPosArr = Array(6,0,-1,-1,-1,  41,42,43,-1,25,   31,32,26,27,27,   15,-1,19,18,18,   -1,11,13,9,-1,	 -1,39,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/11����_���̶��
	elseif (SellSite="11stITS") then
	    xlPosArr = Array(2,46,4,-1,-1,  33,29,28,-1,12,   29,28,30,31,31,   -1,-1,10,38,38,   -1,6,7,-1,-1,	 -1,-1,-1,24,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/GS SHOP
	elseif (SellSite="gseshop") then
	    'xlPosArr = Array(6,8,-1,-1,-1,  15,16,17,-1,10,  11,12,13,14,14,  42,-1,41,44,45,  -1,39,37,35,-1,  7,18,66,47,-1)  ''���ǸŰ� �߰� 2014/03/17 (19) : �հ�ݾ� �������� ���������.
		'xlPosArr = Array(6,8,-1,-1,-1,  15,16,17,-1,10,  11,12,13,14,14,  42,-1,41,44,45,  -1,39,37,35,-1,  7,18,68,47,-1)  ''���ǸŰ� �߰� 2014/03/17 (19) : �հ�ݾ� �������� ���������.
		'xlPosArr = Array(8,10,-1,-1,-1,  17,18,19,-1,12,  14,13,15,16,16,  44,-1,43,46,47,  -1,41,39,37,-1,  9,20,70,49,-1)  ''���ǸŰ� �߰� 2014/03/17 (19) : �հ�ݾ� �������� ���������.
		xlPosArr = Array(8,10,-1,-1,-1,  17,18,19,-1,12,  14,13,15,16,16,  44,-1,43,46,47,  -1,41,39,37,-1,  9,20,70,49,-1,  -1,-1,6)	''��ȯ�ֹ� ��������
	    ArrayLen = UBound(xlPosArr)
	    skipString = "����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/Homeplus
	elseif (SellSite="homeplus") then
'	    xlPosArr = Array(6,8,-1,-1,-1,  15,16,17,-1,10,  11,12,13,14,14,  42,-1,41,44,45,  -1,39,37,35,-1,  7,18,66,47,-1)  ''���ǸŰ� �߰� 2014/03/17 (19) : �հ�ݾ� �������� ���������.
	    xlPosArr = Array(3,2,-1,-1,-1,  8,12,13,-1,9,  12,13,10,11,11,  16,-1,22,21,23,  -1,18,19,14,-1,  27,24,-1,-1,-1)  ''���ǸŰ� �߰� 2014/03/17 (19) : �հ�ݾ� �������� ���������.
	    ArrayLen = UBound(xlPosArr)
	    skipString = "����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "��۸���Ʈ"  '' sheet name Maybe filename in

	'/���������
	elseif (SellSite="ezwel") then
'	    xlPosArr = Array(3,2,-1,-1,-1,  8,12,13,-1,9,  12,13,10,11,11,  16,-1,22,21,23,  -1,18,19,14,-1,  -1,24,-1,-1,-1)
'	    xlPosArr = Array(1,6,-1,-1,-1,  3,5,4,-1,21,  23,22,24,25,25,  -1,-1,17,Array(10,11,12),10,  -1,9,15,8,-1,  2,26,-1,19,20)  '' ����غ��� üũ �߰� 2015/03/02
	    xlPosArr = Array(1,6,-1,-1,-1,  3,5,4,-1,23,  25,24,26,27,27,  -1,-1,17,Array(10,11,12),10,  -1,9,15,8,-1,  2,28,-1,19,20)  '' ����غ��� üũ �߰� 2015/03/02
	    ArrayLen = UBound(xlPosArr)
	    skipString = "����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "��۸���Ʈ"  '' sheet name Maybe filename in

	'/gsisuper
	elseif (SellSite="gsisuper") then
	    'xlPosArr = Array(2,1,-1,-1,-1,  14,15,16,-1,20,  21,22,23,24,25,  6,-1,8,10,9,  -1,7,-1,5,-1,  -1,26,-1,11,-1)
	    'xlPosArr = Array(3,1,-1,-1,-1,  4,5,6,-1,9,  10,11,12,13,13,  -1,-1,20,22,22,  -1,18,19,17,-1,  16,15,-1,-1,-1)
		xlPosArr = Array(3,1,-1,-1,-1,  16,17,18,-1,16,  17,18,19,20,20,  -1,-1,8,12,12,  -1,6,7,5,-1,  4,10,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "�ֹ�����"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24
	'/GS25
	elseif (SellSite="GS25") then
	    xlPosArr = Array(3,1,-1,-1,-1,  9,10,11,-1,14,  16,17,18,19,20,  -1,-1,8,23,23,  -1,6,7,5,-1,  4,21,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ�����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "e-ī�ٷα׻�ǰ����_MD"  '' sheet name Maybe filename in

	elseif (SellSite="cjmall") then
	    'xlPosArr = Array(9,4,-1,-1,-1,   10,13,-1,-1,14,  15,16,17,50,50,  41,-1,22,24,-1,   -1,27,28,25,-1,   -1,34,-1,-1,-1)
		'xlPosArr = Array(8,3,-1,-1,-1,   9,10,-1,-1,11,  12,13,14,15,15,  34,-1,19,21,20,   -1,24,25,23,-1,   -1,27,-1,-1,-1)
		'xlPosArr = Array(10,4,-1,-1,-1,   11,12,-1,-1,13,  14,15,16,17,17,  40,-1,23,25,24,   -1,29,30,27,-1,   10,20,-1,-1,-1)
		xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

'	'/�������
'	elseif (SellSite="privia") then
'	    xlPosArr = Array(3,3,-1,-1,-1,  5,28,27,-1,25,  28,27,37,26,26,  -1,-1,15,17,17,  -1,12,14,-1,-1,  -1,29,-1,19,-1)
'	    ArrayLen = UBound(xlPosArr)
'	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
'	    afile = sFilePath
'	    aSheetName = "first"  '' sheet name Maybe filename in

	'/�������		''2013-11-15 16:40 ������ �����غ�
	elseif (SellSite="privia") then
	    xlPosArr = Array(3,2,-1,-1,-1,  5,29,28,-1,25,  29,28,26,27,27,  -1,-1,17,19,19,  -1,12,14,11,-1,  -1,30,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "first"  '' sheet name Maybe filename in

	'/momastore
	elseif (SellSite="momastore") then
	    'xlPosArr = Array(3,2,-1,-1,-1,  5,29,28,-1,25,  29,28,26,27,27,  -1,-1,17,19,19,  -1,12,14,11,-1,  -1,30,-1,21,-1)
	    xlPosArr = Array(1,0,-1,-1,-1,  5,6,6,-1,5,  6,6,7,8,9,  -1,-1,18,15,16,  -1,11,-1,10,-1,  -1,-1,-1,17,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "�ֹ���ȸ"  '' sheet name Maybe filename in

	'/�����̴���
	elseif (SellSite="NJOYNY") or (SellSite="itsNJOYNY") then
	    xlPosArr = Array(0,3,-1,-1,-1,  5,12,13,14,15,   17,18,19,20,20,   -1,-1,6,11,11,   26,8,10,9,-1,	 -1,33,-1,31,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "�ֹ�����"  '' sheet name Maybe filename in

	'/Ƽ�ϸ���
	elseif (SellSite="ticketmonster") then
		'xlPosArr = Array(1,13,-1,-1,-1,  6,17,17,-1,15,   17,17,19,18,18,   -1,-1,11,10,10,   -1,8,9,-1,-1,	 -1,20,-1,-1,-1)
		'xlPosArr = Array(1,17,-1,-1,-1,  10,22,22,-1,19,   22,22,24,23,23,   29,-1,15,14,14,   -1,12,13,-1,-1,	 -1,25,-1,-1,-1)
		xlPosArr = Array(1,18,-1,-1,-1,  10,12,12,-1,20,   23,23,25,24,24,   30,-1,16,15,15,   -1,13,14,-1,-1,	 -1,26,-1,-1,-1)
		ArrayLen = UBound(xlPosArr)
		skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
		afile = sFilePath
		aSheetName = "�ֹ�����"  '' sheet name Maybe filename in

	'/����Ŭ��
	elseif (SellSite="halfclub") then
		xlPosArr = Array(0,26,-1,-1,-1,  11,14,15,-1,13,   14,15,16,17,17,   2,-1,9,8,8,   -1,3,4,2,-1,	 1,18,-1,-1,-1)
		ArrayLen = UBound(xlPosArr)
		skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
		afile = sFilePath
		aSheetName = "�ֹ�����"  '' sheet name Maybe filename in


	'/��ũ��ٿ���
	elseif (SellSite="thinkaboutyou") then
	    'xlPosArr = Array(1,13,-1,-1,-1,  6,17,17,-1,15,   17,17,19,18,18,   -1,-1,11,10,10,   -1,8,9,-1,-1,	 -1,20,-1,-1,-1)
	    xlPosArr = Array(1,0,-1,-1,-1,  4,5,5,6,17,   18,19,20,21,21,   -1,-1,11,12,15,   -1,9,10,-1,-1,	 -1,22,-1,27,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "�ֹ�����"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24

	'/��ٿ���
	elseif (SellSite="aboutpet") then
	    'xlPosArr = Array(0,42,-1,-1,-1,  3,5,5,4,45,   43,43,44,45,46,   -1,58,23,18,22,   -1,13,-1,10,-1,	 1,68,-1,13,-1)
		xlPosArr = Array(0,45,-1,-1,-1,  3,5,5,4,48,   49,49,50,51,52,   -1,66,22,17,30,   -1,13,-1,10,-1,	 1,76,-1,37,73)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ�����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "order"  '' sheet name Maybe filename in

	'/��Ĺ
	elseif (SellSite="cookatmall") then
	    'xlPosArr = Array(1,13,-1,-1,-1,  6,17,17,-1,15,   17,17,19,18,18,   -1,-1,11,10,10,   -1,8,9,-1,-1,	 -1,20,-1,-1,-1)
	    xlPosArr = Array(0,0,-1,-1,-1,  2,3,4,-1,2,   3,4,5,6,6,   -1,-1,11,10,10,   -1,9,-1,-1,-1,	 1,15,-1,13,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ�����" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "�ֹ�����"  '' sheet name Maybe filename in


	'/��ť
	elseif (SellSite="momQ") then
	    'xlPosArr = Array(1,0,-1,-1,-1,  4,5,5,6,17,   18,19,20,21,21,   -1,-1,11,12,15,   -1,9,10,-1,-1,	 -1,22,-1,27,-1)
	    xlPosArr = Array(3,32,-1,-1,-1,  16,18,19,-1,17,   18,19,20,21,22,   -1,-1,13,27,27,   -1,6,7,5,-1,	 -1,-1,-1,39,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "�ֹ�����"  '' sheet name Maybe filename in

	'/������	'/�̴ϼȼ���
	elseif (SellSite="gabangpop") or (SellSite="itsGabangpop") then
	    xlPosArr = Array(5,31,-1,-1,-1,  20,23,24,-1,20, 23,24,21,22,22,   -1,-1,14,16,16,   -1,12,13,-1,-1,	 6,25,36,18,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "gabangpop_sample"  '' sheet name Maybe filename in

	'/���Ż�
	elseif (SellSite="musinsaITS") or (SellSite="itsMusinsa") then
	    xlPosArr = Array(5,32,-1,-1,-1,   21,24,25,-1,21,    24,25,22,23,23,   -1,-1,14,16,16,   -1,12,13,-1,-1,	 6,26,-1,18,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/GVG
	elseif (SellSite="GVG") then
	    xlPosArr = Array(0,1,-1,-1,-1,   8,12,13,-1,9,   12,13,10,11,11,   -1,-1,7,6,6,   -1,2,4,-1,-1, -1,14,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "GVG_�ֹ��� �������"  '' sheet name Maybe filename in

	'/�÷��̾�
	elseif (SellSite="player") or (SellSite="itsPlayer1") then
	    xlPosArr = Array(1,1,-1,-1,-1,   8,11,12,-1,7,   11,12,9,10,10,   -1,-1,21,24,24,   -1,17,17,-1,-1,   2,13,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24

	'/������ũ
	elseif (SellSite="interpark") then
	    ''xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   42,43,5,9,-1,   11,3,4,42,43, 1,35,38,20)
	    ''xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   41,42,5,9,-1,   11,3,4,41,42,   1,35,38,20) ''�ֹ�������Ű �߰� 20120831
	    ' xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   41,42,5,9,8,   11,3,4,2,42,   1,35,38,20)    ''���ǸŰ� �߰� 2014/01/14 (19) : �հ�ݾ� �������� ���������.
		'xlPosArr = Array(0,-1,-1,-1,-1,   25,27,28,-1,29,   30,31,32,33,33,   42,43,6,9,8,   11,4,5,2,3,   1,36,39,19)    ''���ǸŰ� �߰� 2014/01/14 (19) : �հ�ݾ� �������� ���������.
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,6,9,8,   11,4,5,2,3,   1,37,40,20)    ''���ǸŰ� �߰� 2014/01/14 (19) : �հ�ݾ� �������� ���������.
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,5,9,8,   11,3,4,2,-1,   1,37,40,20)
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,6,9,8,   11,4,5,2,-1,   1,37,40,20)
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,5,9,8,   11,3,4,2,-1,   1,37,40,20)
		'xlPosArr = Array(0,-1,-1,-1,-1,   34,28,29,-1,30,   31,32,33,34,34,   43,44,5,9,8,   11,3,4,2,-1,   1,37,40,20)
        xlPosArr = Array(0,28,-1,-1,-1,   34,35,36,-1,19,   20,21,22,23,23,   25,32,7,15,-1,   27,5,6,4,-1,   1,24,-1,14)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "Sheet0" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/�ݵ�ط��̽�
	elseif (SellSite="bandinlunis") then
	    xlPosArr = Array(0,1,-1,-1,-1,   2,6,7,-1,5,   6,7,11,8,8,   -1,-1,4,-1,-1,   -1,3,10,-1,-1,   -1,9,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "List"  '' sheet name Maybe filename in

	'/��Ʈ��		'/������
	elseif (SellSite="mintstore") or (SellSite="itsMintstore") then
	    xlPosArr = Array(0,1,-1,-1,-1,   2,6,7,-1,2,   6,7,11,8,9,   -1,-1,5,-1,-1,   -1,3,4,-1,-1,   -1,3,4,10,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "List"  '' sheet name Maybe filename in

	'/���� ��Ʈ����
	elseif (SellSite="hottracks") or (SellSite="itsHottracks") then
	    xlPosArr = Array(6,3,-1,-1,-1,   5,8,8,-1,7,   8,8,12,13,14,   -1,-1,18,-1,-1,   26,16,16,15,-1,   -1,9,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "hottracks_sample"  '' sheet name Maybe filename in

	'/����
	elseif (SellSite="byulshopITS") or (SellSite="itsByulshop") then
	    xlPosArr = Array(2,18,-1,-1,-1,   3,5,6,-1,4,   5,6,7,8,8,   -1,-1,14,15,15,   -1,10,10,-1,-1,   1,9,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "byulshop_order_sample"  '' sheet name Maybe filename in
	'/�̹̹ڽ�
	elseif (SellSite="itsMemebox") then
		'xlPosArr = Array(14,0,-1,-1,-1,   10,11,11,-1,10,   11,11,12,13,13,   -1,-1,7,9,9,   -1,5,6,2,-1,   -1,19,-1,-1,-1)
		'xlPosArr = Array(11,0,-1,-1,-1,   7,8,8,-1,7,   8,8,9,10,10,   -1,-1,5,6,6,   -1,3,4,2,-1,   -1,15,-1,-1,-1)
		xlPosArr = Array(4,0,-1,-1,-1,   16,17,17,-1,16,   17,17,19,20,20,   -1,-1,12,13,13,   -1,10,11,7,-1,   -1,24,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/SUHA
	elseif (SellSite="suhaITS") then
	    xlPosArr = Array(0,1,-1,-1,-1,   14,17,18,-1,8,   11,12,9,10,10,   -1,-1,7,6,6,   -1,2,4,3,-1,   -1,13,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "121.88.197.9_orders(1)"  '' sheet name Maybe filename in

	'/������
	elseif (SellSite="gmarket") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(24,3,31,-1,-1,   6,26,25,-1,13,  15,14,16,17,17,   -1,-1,9,5,5,   33,8,10,1,-1,   2,18,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/����
	elseif (SellSite="auction1010") OR (SellSite="gmarket1010") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(23,3,33,-1,-1,   6,22,21,-1,12,  14,13,15,16,16,   34,-1,9,4,24,   25,8,10,1,-1,   2,17,-1,20,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/���̶�� �͵���
	elseif (SellSite="itsWadiz") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/11����
	elseif (SellSite="11st1010") then
'		xlPosArr = Array(2,46,4,-1,-1,  33,29,28,-1,12,   29,28,30,31,31,   -1,-1,10,38,38,   -1,6,7,-1,-1,	 -1,-1,-1,24,-1)
    	xlPosArr = Array(2,46,4,-1,-1,  33,29,28,-1,12,  29,28,30,31,31,   37,-1,10,38,40,   -1,6,7,36,-1,   3,32,-1,27,5, -1,41,39,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24

	'/�ż���TV����..���������� �۾�
	elseif (SellSite="shintvshopping") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/skstoa..���������� �۾�
	elseif (SellSite="skstoa") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/LFmall..���������� �۾�
	elseif (SellSite="LFmall") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/�¿����..���������� �۾�
	elseif (SellSite="goodwearmall10") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/wconcept1010..���������� �۾�
	elseif (SellSite="wconcept1010") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/�������
	elseif (SellSite="nvstorefarm") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "���ֹ߼۰���"  '' sheet name Maybe filename in

	'/������� ���汸
	elseif (SellSite="nvstoremoonbangu") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "���ֹ߼۰���"  '' sheet name Maybe filename in

	'/������� Ĺ�ص�
	elseif (SellSite="Mylittlewhoopee") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "���ֹ߼۰���"  '' sheet name Maybe filename in

	'/������� �����ϱ�
	elseif (SellSite="nvstoregift") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "���ֹ߼۰���"  '' sheet name Maybe filename in

	'/�������Ŭ����
	elseif (SellSite="nvstorefarmclass") then
    	xlPosArr = Array(1,55,13,-1,-1,   7,42,42,-1,9,  39,39,-1,-1,-1,   36,18,19,21,24,   -1,15,17,14,-1,   0,44,-1,-1,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = "���ֹ߼۰���"  '' sheet name Maybe filename in

	''cjmallITS		'/���� 2013/03/07�߰�../���� 2014-02-21 ���� ���� �ٲ� ���� �ϴ� ����../���� 2014-10-02 ���� ���� �ٲ� ���� �ϴ� ����../
	elseif (SellSite="cjmallITS") or (SellSite="itsCjmall") then
		'xlPosArr = Array(9,4,-1,-1,-1,   10,13,-1,-1,14,  15,16,17,50,50,  -1,-1,22,24,-1,   -1,27,28,26,-1,   -1,34,-1,-1,-1)
		xlPosArr = Array(8,3,-1,-1,-1,   9,10,-1,-1,11,  12,13,14,15,15,  -1,-1,19,21,-1,   -1,24,25,23,-1,   -1,27,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in
    '/������
	elseif (SellSite="hiphoper") or (SellSite="itsHiphoper") then
	    xlPosArr = Array(5,-1,-1,-1,-1,   0,3,4,-1,0,   3,4,1,2,2,   -1,-1,10,11,11,   -1,8,7,6,-1,   -1,12,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "������" ''
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

    '/���̶��_29cm
	elseif (SellSite="its29cm") then
	    xlPosArr = Array(0,1,-1,-1,-1,   2,3,4,5,6,   7,8,9,10,10,   -1,-1,17,16,16,   -1,14,15,13,-1,   -1,11,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/��������		'/�̴ϼȼ���
    elseif (SellSite="wizwid") or (SellSite="itsWizwid") then
	    xlPosArr = Array(4,2,-1,-1,-1,   14,20,21,-1,15,   18,19,16,17,17,   -1,-1,11,13,13,   -1,7,9,6,-1,   3,23,29,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ù���ڵ�" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/����������	'/�̴ϼȼ���
	elseif (SellSite="wconcept") or (SellSite="itsWconcept") then
	    If SellSite = "itsWconcept" Then
	    	xlPosArr = Array(2,1,-1,-1,-1,   13,19,20,-1,14,   17,18,15,16,16,   -1,-1,10,12,12,   -1,6,8,-1,-1,   3,21,27,-1,-1)
	    Else
	    	xlPosArr = Array(3,2,-1,-1,-1,   16,22,23,-1,17,   20,21,18,19,19,   -1,-1,13,15,15,   -1,7,9,-1,-1,   4,24,-1,-1,-1)
	    End If
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ù���ڵ�" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/�÷�TV		'/������
	elseif (SellSite="ollehtv") then
	    xlPosArr = Array(2,17,18,-1,-1,  9,11,12,-1,14,   11,12,15,16,16,   -1,-1,8,7,7,   -1,5,6,4,-1,   -1,-1,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "�ֹ���ȣ" ''Array("�߼۰���","�ݻ����ð�","�۹߼۰���","�׸��","�ֹ���ȣ")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/�ϳ�����		'/������
    elseif (SellSite="hanatour") then
        xlPosArr = Array(0,15,15,-1,-1		,2,4,5,-1,3		,4,5,16,16,16		,1,-1,8,9,9		,10,6,7,1,-1		,-1,17,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/�м��÷���
    elseif (SellSite="fashionplus") or (SellSite="itsFashionplus") then
    	xlPosArr = Array(1,6,6,-1,-1,     28,27,26,-1, 2,     27,26,25,23,23 ,     11,-1,13,14,14,     -1,12,4 ,11,-1,     -1,29,-1,24,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/������
    elseif (SellSite="giftting") then
    	xlPosArr = Array(1,13,13,-1,-1,     16,17,19,-1, 18,     17,19,20,21,21,     7,-1,22,23,23,     -1,5,6,-1,-1,     10,14,-1,24,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="������"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in


''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

'' 18 �ǸŰ� = ���ΰ�
'' 19 ���ǸŰ�  (���� �ݿ� �ݾ�) �ִ°�� �� �־��ֽǰ�. 2013/10/24

	'/���̶��_�����Ǿ�
    elseif (SellSite="itsbenepia") then
    	xlPosArr = Array(0,4,4,-1,-1,     5,22,23,-1, 14,     22,23,25,26,27,     -1,-1,10,11,11,     -1,6,8,7,-1,     1,28,-1,13,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/���̶��_īī���彺���
    elseif (SellSite="itskakaotalkstore") then
    	'xlPosArr = Array(2,20,-1,-1,-1,     11,13,13,-1, 11,     13,13,17,16,16,     31,-1,6,21,23,     -1,4,5,3,-1,     -1,18,-1,37,-1)
		'xlPosArr = Array(2,19,-1,-1,-1,     10,12,12,-1, 10,     12,12,16,15,15,     30,-1,6,20,22,     -1,4,5,3,-1,     -1,17,-1,35,-1)
		xlPosArr = Array(2,19,-1,-1,-1,     10,12,12,-1, 10,     12,12,16,15,15,     30,-1,6,20,22,     -1,4,5,3,-1,     -1,17,-1,35,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/����Ŀ������īī��
    elseif (SellSite="itsKaKaoMakers") then
    	'xlPosArr = Array(2,0,0,-1,-1,     12,13,14,-1, 12,     13,14,16,15,15,     7,-1,10,8,8,     -1,3,4,6,-1,     -1,17,-1,9,-1)
    	xlPosArr = Array(34,0,0,-1,-1,     11,13,13,-1, 11,     13,13,17,16,16,     -1,-1,6,22,22,     -1,4,5,3,-1,     2,18,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/īī������Ʈ
    ' elseif (SellSite="kakaogift") then
    ' 	xlPosArr = Array(2,0,0,-1,-1,     11,13,15,-1, 11,     13,15,17,16,16,     30,31,6,21,21,     -1,4,5,3,-1,     -1,18,-1,36,-1)
	'     ArrayLen = UBound(xlPosArr)
	'     skipString="������"
	'     afile = sFilePath
	'     aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/īī������Ʈ
    elseif (SellSite="kakaogift") then
    	xlPosArr = Array(3,1,1,-1,-1,     11,13,15,-1, 11,     13,15,17,16,16,     33,34,7,21,21,     25,5,6,4,-1,     -1,18,-1,38,22)
	    ArrayLen = UBound(xlPosArr)
	    skipString="������"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
	'/���̶��_īī�������ϱ�
    elseif (SellSite="itskakao") then
    	'xlPosArr = Array(2,0,0,-1,-1,     12,13,14,-1, 12,     13,14,16,15,15,     7,-1,10,8,8,     -1,3,4,6,-1,     -1,17,-1,9,-1)
    	'xlPosArr = Array(2,0,0,-1,-1,     11,13,13,-1, 11,     13,13,17,16,16,     30,-1,6,21,23,     -1,4,5,3,-1,     -1,18,-1,-1,-1)
		xlPosArr = Array(2,0,0,-1,-1,     10,12,12,-1, 10,     12,12,16,15,15,     29,-1,6,20,22,     -1,4,5,3,-1,     -1,17,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/���̶�Ҽ�
    elseif (SellSite="ithinksoshop") then
'    	xlPosArr = Array(0,0,0,-1,-1,     11,12,12,-1, 11,     12,12,13,14,15,     -1,-1,9,10,4,     -1,7,8,6,-1,     1,2,-1,-1,-1)
    	xlPosArr = Array(1,3,4,-1,-1,     28,31,30,-1, 32,     34,33,27,25,26,     -1,-1,24,7,9,     -1,6,13,5,-1,     2,35,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/����ũ �����̽�
	elseif (SellSite="wemakeprice") then
	    'xlPosArr = Array(0,1,1,-1,-1,     6,15,15,-1,14,     15,15,16,17,17,     23,-1,12,13,13,     -1,9,11,8,10,     -1,22,19,-1,-1)
	    xlPosArr = Array(0,1,1,-1,-1,     6,14,14,-1,13,     14,14,15,16,16,     22,-1,11,12,12,     -1,8,10,7,9,     -1,21,18,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ�ID"
	    afile = sFilePath
	    aSheetName = "Worksheet"  '' sheet name Maybe filename in

''''				0				1				2				3				4
''''				�ֹ���ȣ, 		�ֹ���, 		�Ա���, 		���Ҽ���, 		�ֹ���ID,
''''				5				6				7				8				9
''''				�ֹ���, 		�ֹ�����ȭ,		�ֹ����޴���ȭ,	�ֹ����̸���, 	������,
''''				10				11				12				13				14
''''				��������ȭ,		�������ڵ���,	������Zip,		������addr1,	������addr2,
''''				15				16				17				18				19
''''				��ǰ�ڵ�, 		�ɼ��ڵ�, 		����, 			�ǸŰ�, 		���ǸŰ� XX�Һ��ڰ�,
''''				20				21				22				23				24
''''				�����, 		��ǰ��, 		�ɼǸ�, 		��ü��ǰ�ڵ�,	��ü�ɼ��ڵ�,
''''				25				26				27				28				29
''''				�ֹ�������Ű, 	�ֹ����ǻ���, 	��ǰ�䱸����1, 	��ۺ�, 		ETC1��,
''''				30				31				32				33				34
''''				ETC2(����Ŀ���θ�),		ETC3(����Ŀ��ǰ�ڵ�),			ETC4(����Ŀ�ֹ���ȣ),			�����ڵ�,			ETC5(����Ŀ���ҿ���)
''''				35				36				37				38				39
''''				�ؿ��ǸŰ�,		�ؿܹ�ۺ�		�ؿܽ��ǸŰ�	����Ʈ���߰�����(�ؿ�)	�̻��

	'/�ż����(SSG)
	elseif (SellSite="ssg") then
        'xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
        'xlPosArr = Array(8,35,-1,-1,-1,  36,38,39,-1,37,  38,39,40,41,41,   21,-1,29,30,30,   -1,20,23,22,-1,   11,47,24,-1,6, -1,-1,-1,-1,-1,	-1,-1,-1,7,-1)
		xlPosArr = Array(9,4,-1,-1,-1,  36,38,39,-1,37,  38,39,40,41,41,   21,-1,30,31,31,   -1,20,24,23,25,   12,47,-1,-1,7, -1,-1,-1,-1,-1,	-1,-1,-1,8,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet1"
	    afile = sFilePath
	    aSheetName = maybeSheetName

	'/����
	elseif (SellSite="coupang") then
	    'xlPosArr = Array(2,4,-1,-1,-1,	11,12,12,10,13,		14,14,15,16,16,	-1,-1,7,5,5,	-1,9,-1,-1,-1,	-1,17,-1,6,-1)
        'xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
		xlPosArr = Array(2,9,-1,-1,-1,	24,26,26,25,27,		28,28,29,30,30,	16,16,22,23,23,	-1,10,11,13,-1,	-1,31,14,20,1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="��۰���"
	    afile = sFilePath
	    aSheetName = maybeSheetName

	'/Ž�� - ������
	elseif (SellSite="5") then
	    xlPosArr = Array(0,2,42,39,6,5,9,10,11,13,16,17,14,15,15,22,-1,28,29,31,-1,24,25,-1,-1,-1,12,19,20,21)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = "TOMS �ֹ�����"  '' sheet name Maybe filename in

	'/�ٹ����� - �ε����߰�����
	elseif (SellSite="4") then
	    xlPosArr = Array(1,2,-1,-1,-1,3,4,5,6,7,8,9,10,11,12,15,-1,19,18,-1,-1,16,17,21,-1,0,13,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�Ϸù�ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

    else
	    response.write "<script>alert('���θ� �ڵ尡 �������� �ʾҽ��ϴ�. -"&SellSite&"');</script>"
	    response.end
	end if

''ReDim xlRow(ArrayLen)
Dim xlRowALL

''rw "ArrayLen="&ArrayLen

dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('������ �ùٸ��� �ʰų� ������ �����ϴ�. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="�ܺ� ���̺� ������ �߸��Ǿ����ϴ�.") then
        response.write "<script>alert('�������� Save As Excel 97 -2003 ���չ��� ���·� ������ ����ϼ���.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

''������ ó��.
Dim iLine, iResult
Dim paramInfo, retParamInfo, RetErr, retErrStr, sqlStr
Dim POS1,POS2,POS3, okCNT, bufItemName, bufItemSplit, bufrowObj, bufOneItemName, bufItemNo, bufOptionName
Dim errCNT, totErrMsg, tmpItemname
Dim rtitemid, rtitemoption, rtSellPrice
Dim t_addDlvPrice ,t_deliverytype ,t_sellcash ,t_defaultFreeBeasongLimit ,t_defaultDeliverPay, tempShippay
Dim tmppreAddr, tmpnextAddr
okCNT = 0 : errCNT = 0

Dim pcnt : pcnt = UBound(xlRowALL)
'    IF (sellsite="wemakeprice") then
'        for i=0 to pcnt
'            if IsObject(xlRowALL(i)) then
'                set iLine = xlRowALL(i)
'                bufItemName = iLine.FItemArray(21)
'                if (InStr(bufItemName,"] [")>0) then
'                    bufItemSplit = split(bufItemName,"] [")
'                    for k=LBound(bufItemSplit) to UBound(bufItemSplit)
'                        bufOneItemName = bufItemSplit(k)
'                        bufItemNo  = Trim(replace(Right(Replace(Replace(bufOneItemName,"]",""),"��",""),2)," ",""))
'
'                        IF K=0 then
'                            iLine.FItemArray(21) = bufOneItemName + "]"
'                            iLine.FItemArray(17) = bufItemNo
'    'rw iLine.FItemArray(21)
'    'rw iLine.FItemArray(17)
'                        ELSE
'                            set bufrowObj = new TXLRowObj
'                            bufrowObj.setArrayLength(UBound(iLine.FItemArray))
'
'                            For m=LBound(iLine.FItemArray) to UBound(iLine.FItemArray)
'                                bufrowObj.FItemArray(m) = iLine.FItemArray(m)
'                            Next
'
'                            bufrowObj.FItemArray(21)="[" + bufOneItemName
'                            bufrowObj.FItemArray(17)=bufItemNo
'    'rw "B:"&bufrowObj.FItemArray(21)
'    'rw "B:"&bufrowObj.FItemArray(17)
'                            ReDim Preserve xlRowALL(UBound(xlRowALL)+1)
'                            set xlRowALL(UBound(xlRowALL)) =  bufrowObj
'                        ENd IF
'                    next
'                end if
'             end if
'        Next
'    End IF

IF (sellsite<>"interpark") then
    dbget.BeginTrans
end if

    for i=0 to UBound(xlRowALL)
    ''rw UBound(xlRowALL)
        if (i>3000) then Exit For  ''''
        if IsObject(xlRowALL(i)) then
            set iLine = xlRowALL(i)
            ''��¥ ���� ���� - wemakePrice
            iLine.FItemArray(1) = Replace(iLine.FItemArray(1),".","-")
            iLine.FItemArray(1) = Replace(iLine.FItemArray(1),"/","-")
            iLine.FItemArray(2) = Replace(iLine.FItemArray(2),"/","-")

            if Len(iLine.FItemArray(1)) > 19 then
            	iLine.FItemArray(1) = Left(iLine.FItemArray(1), 19)
            end if

            if Len(iLine.FItemArray(2)) > 19 then
            	iLine.FItemArray(2) = Left(iLine.FItemArray(2), 19)
            end if

            if (iLine.FItemArray(9)="-") then iLine.FItemArray(9)=iLine.FItemArray(5)
            if (iLine.FItemArray(10)="-") then iLine.FItemArray(10)=iLine.FItemArray(6)
            if (iLine.FItemArray(11)="-") then iLine.FItemArray(11)=iLine.FItemArray(7)

''------------------------------------------
            IF (sellsite="shoplinker") then
                ''�ֹ���/�Ա��� ''20130813125800
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),5,2)&"-"&Mid(iLine.FItemArray(1),7,2)&" "&Mid(iLine.FItemArray(1),9,2)&":"&Mid(iLine.FItemArray(1),11,2)&":"&Mid(iLine.FItemArray(1),13,2)
                iLine.FItemArray(2) = iLine.FItemArray(1)

                ''iLine.FItemArray(16) �ɼ��ڵ�
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(24)<>"") then
'                    iLine.FItemArray(16)=getOptionCodByOption(iLine.FItemArray(15),iLine.FItemArray(24))
'                end if
'
'                if (iLine.FItemArray(16)="") and ((iLine.FItemArray(24)="NONE") or (iLine.FItemArray(24)="")) then
'                    iLine.FItemArray(16)="0000"
'                end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),replace(replace(iLine.FItemArray(22)," / FREE","")," FREE",""))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(16)="") then
                    iLine.FItemArray(16)="0000"
                end if

                ''��۸޼��� // gs / ��ü����,����ؼ���

                if (iLine.FItemArray(26)=",,") then iLine.FItemArray(26)="" ''HOTTRACKS

                if (iLine.FItemArray(30)="HOTTRACKS") and (iLine.FItemArray(28)>2500) then ''��ۺ��ʵ忡 ��ǰ���� �� ����
                    if (iLine.FItemArray(34)="������") then
                        iLine.FItemArray(28)="2500"
                    else
                        iLine.FItemArray(28)="0"
                    end if
                end if

            end if

            IF (sellsite="cn10x10") then
            	'//���޸� ��ǰ���� ������ �ٹ����� ��ǰ���� �ֳ� Ȯ���ؼ� �޾ƿ�
                if iLine.FItemArray(23)<>"" then
                    iLine.FItemArray(15) = getItemIDByUpcheItemCode(sellsite,iLine.FItemArray(23))
                end if

				if (iLine.FItemArray(24)<>"") then
					if right(iLine.FItemArray(24),4)<>"0000" then
						iLine.FItemArray(24) = right(iLine.FItemArray(24),4)
						iLine.FItemArray(16)=getOptionCodByOption(iLine.FItemArray(15),iLine.FItemArray(24))
					else
						iLine.FItemArray(16)="0000"
					end if
				else
					iLine.FItemArray(16)="0000"
				end if

				iLine.FItemArray(5) = replace(replace(replace(iLine.FItemArray(5),"  ","!@@@!")," ",""),"!@@@!"," ")
				iLine.FItemArray(9) = replace(replace(replace(iLine.FItemArray(9),"  ","!@@@!")," ",""),"!@@@!"," ")
            END IF

			If (sellsite="cnglob10x10") Then
'On Error Resume Next
				tmpOverseasRealprice = iLine.FItemArray(19)
				If Lcase(iLine.FItemArray(30)) <> "gift" Then
					iLine.FItemArray(6) = replace(iLine.FItemArray(6), "+", "")						'��ȭ��ȣ�� + �������� ġȯ
					iLine.FItemArray(6) = replace(iLine.FItemArray(6), " ", "-")					'��ȭ��ȣ�� ��ũ�� -�� ġȯ

					iLine.FItemArray(7) = replace(iLine.FItemArray(7), "+", "")						'��ȭ��ȣ�� + �������� ġȯ
					iLine.FItemArray(7) = replace(iLine.FItemArray(7), " ", "-")					'��ȭ��ȣ�� ��ũ�� -�� ġȯ

					iLine.FItemArray(10) = replace(iLine.FItemArray(10), "+", "")					'��ȭ��ȣ�� + �������� ġȯ
					iLine.FItemArray(10) = replace(iLine.FItemArray(10), " ", "-")					'��ȭ��ȣ�� ��ũ�� -�� ġȯ

					iLine.FItemArray(11) = replace(iLine.FItemArray(11), "+", "")					'��ȭ��ȣ�� + �������� ġȯ
					iLine.FItemArray(11) = replace(iLine.FItemArray(11), " ", "-")					'��ȭ��ȣ�� ��ũ�� -�� ġȯ

					iLine.FItemArray(18) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(18)) / iLine.FItemArray(17)) * 0.1) * 10		'������ �ֹ�ǰ�� �����ݾ��� ��ǰ��ü������ �� * ������ ����...���� (ȯ�� * �ֹ�ǰ�� �����ݾ�)/���� �� ������ ����
					iLine.FItemArray(19) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(19)) / iLine.FItemArray(17)) * 0.1) * 10		'���ǸŰ��ε�..���� ����� �ݾ��� ���� ������ ����..���ִ븮�Կ��� ���߰��� ���� �Է� ��û��..
					iLine.FItemArray(37) = CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))								'�ؿܽ��ǸŰ�
					iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10	'��ۺ� = ȯ�� * ��ۺ� �� ������ ����

					overseasPrice			= iLine.FItemArray(35)
					overseasDeliveryPrice	= iLine.FItemArray(36)
					overseasRealPrice		= CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))
					'######################## ����������� �� If�� #################################
'					If (iLine.FItemArray(24) <> "") Then
'						If right(iLine.FItemArray(24),4) <> "0000" Then
'							iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(24))
'						Else
'							iLine.FItemArray(16)="0000"
'						End if
'					Else
'						iLine.FItemArray(16)="0000"
'					End If
					'######################## ����������� �� If�� #################################

					'######################## ����� �������� �� If��  #################################
					If iLine.FItemArray(22) <> "" Then
						iLine.FItemArray(22) = Trim(iLine.FItemArray(22))
						iLine.FItemArray(22) = Replace(iLine.FItemArray(22), " ", "")			'���� ġȯ
						'iLine.FItemArray(22) = Split(iLine.FItemArray(22), ":")(1)
					End If

					If (iLine.FItemArray(24) <> "") Then
						iLine.FItemArray(16) = getOptionCodeByMakeShopOptCode(iLine.FItemArray(15),iLine.FItemArray(24))
					Else
						iLine.FItemArray(16)="0000"
					End If

'					if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
'						if (iLine.FItemArray(22)<>"") then
'							iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'						else
'							iLine.FItemArray(16)="0000"
'						end if
'					end if
					'######################## ����� �������� �� If��  #################################
				Else

				End If
'On Error Goto 0
			End If

			If (sellsite="cnhigo") Then
				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(19)
				iLine.FItemArray(18) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(18)) / iLine.FItemArray(17)) * 0.1) * 10		'������ �ֹ�ǰ�� �����ݾ��� ��ǰ��ü������ �� * ������ ����...���� (ȯ�� * �ֹ�ǰ�� �����ݾ�)/���� �� ������ ����
				iLine.FItemArray(19) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(19)) / iLine.FItemArray(17)) * 0.1) * 10		'���ǸŰ��ε�..���� ����� �ݾ��� ���� ������ ����..���ִ븮�Կ��� ���߰��� ���� �Է� ��û��..
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10	'��ۺ� = ȯ�� * ��ۺ� �� ������ ����

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))

				If iLine.FItemArray(16) = "" Then
					iLine.FItemArray(16)="0000"
				End If
			End If

			If (sellsite="cnugoshop") Then
				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(19)
				iLine.FItemArray(18) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(18)) / iLine.FItemArray(17)) * 0.1) * 10		'������ �ֹ�ǰ�� �����ݾ��� ��ǰ��ü������ �� * ������ ����...���� (ȯ�� * �ֹ�ǰ�� �����ݾ�)/���� �� ������ ����
				iLine.FItemArray(19) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(19)) / iLine.FItemArray(17)) * 0.1) * 10		'���ǸŰ��ε�..���� ����� �ݾ��� ���� ������ ����..���ִ븮�Կ��� ���߰��� ���� �Է� ��û��..
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10	'��ۺ� = ȯ�� * ��ۺ� �� ������ ����

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))

				If iLine.FItemArray(16) = "" Then
					iLine.FItemArray(16)="0000"
				End If
			End If

			If (sellsite="11stmy") Then
				Dim oDateArr, spOptName
				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(18) - iLine.FItemArray(19)		'�ǸŰ�(iLine.FItemArray(18)) - ������(iLine.FItemArray(19))�� tmpOverseasRealprice�� �Է�
				reserve01 = iLine.FItemArray(38)

				oDateArr = ""
				oDateArr		= Split(iLine.FItemArray(1), "-")
				If Ubound(oDateArr) > 0 Then
					iLine.FItemArray(1) = oDateArr(2)&"-"&oDateArr(1)&"-"&oDateArr(0)
				End If
				iLine.FItemArray(2) = iLine.FItemArray(1)
				iLine.FItemArray(18) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(18)) * 0.1) * 10		'�ǸŰ� = (ȯ�� * ���ܰ�) ������ ����
				iLine.FItemArray(19) = Int(CLng(iLine.FItemArray(29) * tmpOverseasRealprice) * 0.1) * 10		'���ǸŰ� = (ȯ�� * ���ܰ�) ������ ����
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10		'��ۺ� = ȯ�� * ��ۺ� �� ������ ����

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= tmpOverseasRealprice

				If (iLine.FItemArray(16) = "") and (iLine.FItemArray(15) <> "") then       ''�ɼ��ڵ尡 ��.
					If (iLine.FItemArray(22) <> "") then
						spOptName = Trim(Split(iLine.FItemArray(22), "/")(0))
						iLine.FItemArray(16) = get11stOptionCodeByOptionName(iLine.FItemArray(15), Trim(Split(iLine.FItemArray(22), "/")(0)) )
						iLine.FItemArray(22) = spOptName
					Else
						iLine.FItemArray(16) = "0000"
					End If
				End if
			End If

			If (sellsite="zilingo") Then
				iLine.FItemArray(1) = CDate(iLine.FItemArray(1))
				iLine.FItemArray(2) = CDate(iLine.FItemArray(2))

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= 0
				overseasRealPrice		= iLine.FItemArray(35)

				iLine.FItemArray(28) = 0			'��ۺ�
				iLine.FItemArray(36) = 0			'�ؿܹ�ۺ�

				tmpVal = getItemidOptionCodeByZilignoGoodno(iLine.FItemArray(23))
				iLine.FItemArray(15) = Split(tmpVal, "||")(0)
                iLine.FItemArray(16) = Split(tmpVal, "||")(1)
			End If

			If (sellsite="etsy") Then
				Dim splitYear, splitMonth, splitDay
				splitYear	= Split(iLine.FItemArray(1), "-")(2)
				splitMonth	= Split(iLine.FItemArray(1), "-")(0)
				splitDay	= Split(iLine.FItemArray(1), "-")(1)
				iLine.FItemArray(1) = splitYear & "-" & splitMonth & "-" & splitDay
				iLine.FItemArray(2) = splitYear & "-" & splitMonth & "-" & splitDay

				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(18) - iLine.FItemArray(19)		'�ǸŰ�(iLine.FItemArray(18)) - ������(iLine.FItemArray(19))�� tmpOverseasRealprice�� �Է�

				iLine.FItemArray(18) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(18)) * 0.1) * 10		'�ǸŰ� = (ȯ�� * ���ܰ�) ������ ����
				iLine.FItemArray(19) = Int(CLng(iLine.FItemArray(29) * tmpOverseasRealprice) * 0.1) * 10		'���ǸŰ� = (ȯ�� * ���ܰ�) ������ ����
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10		'��ۺ� = ȯ�� * ��ۺ� �� ������ ����

				iLine.FItemArray(14) = iLine.FItemArray(14) & " " & iLine.FItemArray(38) & " " & iLine.FItemArray(39)	'�ּҰ� ������ ������ ����

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= tmpOverseasRealprice

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
			End If

            IF (sellsite="bandinlunis") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(18) = rtSellPrice
                iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''�ֹ���
                iLine.FItemArray(5) = Trim(splitvalue(iLine.FItemArray(5),"(",0))   ''�ֹ���

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="hottracks") or (SellSite="itsHottracks") then

            	'//��ǰ�� �ɼǸ��� ���� ���ֽ�.. ��ġ ����ؼ� ©��
            	if instr(iLine.FItemArray(21),"(") > 0 then
            		iLine.FItemArray(21) = left(iLine.FItemArray(21), instr(iLine.FItemArray(21),"(")-2 )
            		iLine.FItemArray(22) = rtrim(replace(mid(iLine.FItemArray(22), instr(iLine.FItemArray(22),"(")+1 , 96 ),")",""))

            	'//�ɼǾ���
            	else
            		iLine.FItemArray(21) = iLine.FItemArray(21)
            		iLine.FItemArray(22) = ""
            	end if

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(18) = rtSellPrice
                iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''�ֹ���
                iLine.FItemArray(5) = Trim(splitvalue(iLine.FItemArray(5),"(",0))   ''�ֹ���

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

	        	'//�����ȣ ���̿� "-" �� �������
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="byulshopITS") or (SellSite="itsByulshop") then

            	'//��ǰ�� �ɼǸ��� ���� ���ֽ�.. ��ġ ����ؼ� ©��
            	if instr(iLine.FItemArray(21),"(") > 0 then
            		iLine.FItemArray(21) = left(iLine.FItemArray(21), instr(iLine.FItemArray(21),"(")-1 )
            		iLine.FItemArray(22) = rtrim(replace(mid(iLine.FItemArray(22), instr(iLine.FItemArray(22),"(")+1 , 96 ),")",""))

            	'//�ɼǾ���
            	else
            		iLine.FItemArray(21) = iLine.FItemArray(21)
            		iLine.FItemArray(22) = ""
            	end if

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))    ''�ֹ���

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="itsMemebox") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                iLine.FItemArray(18) = Trim(iLine.FItemArray(18))
                iLine.FItemArray(17) = Trim(iLine.FItemArray(17))
                iLine.FItemArray(19) = Trim(iLine.FItemArray(19))

                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))      '''������ ������ �ܰ��� ����.
                    end if
				end if

                iLine.FItemArray(19) = iLine.FItemArray(18)
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="ithinksoshop") then
				Dim preAddr, nextAddr
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))
                iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2),10))

				'############ 2017-03-03 ������ ���� ###############
                preAddr		= Trim(iLine.FItemArray(13))		'�����ּ� ��ü
                nextAddr	= Trim(iLine.FItemArray(14))		'���θ��ּ� ��ü
				If (preAddr = "") AND (nextAddr <> "") Then		'���� ������ �����ּҴ� ���� ���θ� �ּҸ� �ִٸ�
					iLine.FItemArray(13) = nextAddr
					iLine.FItemArray(14) = nextAddr
				ElseIf (preAddr <> "") AND (nextAddr = "") Then	'���� ������ ���θ� �ּҰ� ���� ���� �ּҸ� �ִٸ�
					iLine.FItemArray(13) = preAddr
					iLine.FItemArray(14) = preAddr
				End If
				'############ 2017-03-03 ������ ���� �� ###############

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="suhaITS") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/�̴ϼȼ����ϰ�� ����ڰ� �ǸŰ����� �ȳִ� ��쿡 �Һ��ڰ��� ��ü
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
                    if (iLine.FItemArray(22)<>"��ǰ-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

			'cjmallITS �����߰�
            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			IF (sellsite="cjmallITS")  or (SellSite="itsCjmall") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
				iLine.FItemArray(6) = Replace(iLine.FItemArray(6),")","-")
				iLine.FItemArray(10)= Replace(iLine.FItemArray(10),")","-")
				iLine.FItemArray(11)= Replace(iLine.FItemArray(11),")","-")

'				Dim cjAddr1, cjAddr2
'				cjAddr1 = Split(iLine.FItemArray(13),"|")(0)
'				cjAddr2 = Split(iLine.FItemArray(13),"|")(1)
'				iLine.FItemArray(13) = cjAddr1
'				iLine.FItemArray(14) = cjAddr2
                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''�ֹ���

                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)
                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gmarket") then
				'########################################### 2015-07-03 �Ͻ��� ��� ###########################################
				sItemname = Split(iLine.FItemArray(22),".")(1)
				sItemname = Split(Trim(sItemname),":")(0)
                iLine.FItemArray(1) = left(iLine.FItemArray(1),10)
                iLine.FItemArray(2) = left(iLine.FItemArray(2),10)
				iLine.FItemArray(18) = 4900			'�ǸŰ�
				iLine.FItemArray(19) = 4900			'���ǸŰ�
				iLine.FItemArray(21) = sItemname	'��ǰ��
         		iLine.FItemArray(22) = mid(iLine.FItemArray(22),1,instr(iLine.FItemArray(22),"/")-1)
        		iLine.FItemArray(22) = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),":")+1,100)
				iLine.FItemArray(22) = left(iLine.FItemArray(22),2)		'�ɼǸ�

				Select Case sItemname
					Case "YELLOW"			iLine.FItemArray(15) = 849958
					Case "PINK"				iLine.FItemArray(15) = 849956
					Case "PASTEL PINK"		iLine.FItemArray(15) = 849954
					Case "PASTEL BLUE"		iLine.FItemArray(15) = 849949
					Case "ORANGE"			iLine.FItemArray(15) = 849948
					Case "GREEN"			iLine.FItemArray(15) = 849947
				End Select
				iLine.FItemArray(16) = getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
				'########################################### 2015-07-03 �Ͻ��� �� ###########################################

'            	'//�ɼǾ��°��
'            	if instr(iLine.FItemArray(22),":") = "" or instr(iLine.FItemArray(22),":") = "0" then
'            		iLine.FItemArray(22) = ""
'
'            	'//�ɼ��ִ°�� : �� / ���̿� �ɼ��� �߷���
'            	else
'            		iLine.FItemArray(22) = mid(iLine.FItemArray(22),1,instr(iLine.FItemArray(22),"/")-1)
'            		iLine.FItemArray(22) = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),":")+1,100)
'            	end if
'
'            	'//�ɼǾ��°��	����ó��
'            	if instr(iLine.FItemArray(17),":") = "" or instr(iLine.FItemArray(17),":") = "0" then
'            		iLine.FItemArray(17) = left(iLine.FItemArray(17), instr(iLine.FItemArray(17),"��")-1)
'
'            	'//�ɼ��ִ°�� ����ó�� / �� �� ���̿� ������ �߷���
'            	else
'            		iLine.FItemArray(17) = mid(iLine.FItemArray(17), instr(iLine.FItemArray(17),"/")+1,100)
'            		iLine.FItemArray(17) = left(iLine.FItemArray(17), len(iLine.FItemArray(17))-1)
'            		'iLine.FItemArray(17) = mid(iLine.FItemArray(17), instr(iLine.FItemArray(17),"/")+1 , (len(iLine.FItemArray(17))-instr(iLine.FItemArray(17),"/"))-1  )
'            	end if
'
'                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
'                iLine.FItemArray(1) = left(iLine.FItemArray(1),10)
'                iLine.FItemArray(2) = left(iLine.FItemArray(2),10)
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
'                iLine.FItemArray(15) = rtitemid
'                iLine.FItemArray(16) = rtitemoption
'                '//�������� ������
'                'iLine.FItemArray(18) = rtSellPrice
'                'iLine.FItemArray(19) = rtSellPrice
'
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
'                    if (iLine.FItemArray(22)<>"��ǰ-") then
'                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'
'                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
'
'				if iLine.FItemArray(28) <> "" then
'					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
'				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="mintstore") or (SellSite="itsMintstore") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(18) = rtSellPrice
                iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="wizwid") or (SellSite="itsWizwid") then
            	iLine.FItemArray(0) = Trim(iLine.FItemArray(0))
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(23) = Trim(iLine.FItemArray(23))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

                '//�������� ������
                iLine.FItemArray(18) = rtSellPrice			'2017-05-11 ������ �ּ� ����
                iLine.FItemArray(19) = rtSellPrice			'2017-05-11 ������ �ּ� ����
                '/�̴ϼȼ����ϰ�� ����ڰ� �ǸŰ����� �ȳִ� ��쿡 �Һ��ڰ��� ��ü
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),5,2)&"-"&Mid(iLine.FItemArray(1),7,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
                    if (iLine.FItemArray(22)<>"��ǰ-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="hanatour") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),6,2)&"-"&Mid(iLine.FItemArray(1),9,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
                    if (iLine.FItemArray(22)<>"��ǰ-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="fashionplus") or (SellSite="itsFashionplus") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = "20" & Left(iLine.FItemArray(1),2)&"-"&Mid(iLine.FItemArray(1),4,2)&"-"&Mid(iLine.FItemArray(1),7,2)
				iLine.FItemArray(2) = "20" & Left(iLine.FItemArray(2),2)&"-"&Mid(iLine.FItemArray(2),4,2)&"-"&Mid(iLine.FItemArray(2),7,2)
                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
                    if (iLine.FItemArray(22)<>"��ǰ-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

                iLine.FItemArray(25) = iLine.FItemArray(0)

				'//�ֹ���ȣ�� - �� ������ -�������� ©�󳻼� �ֹ���ȣ�� �Է�
                if instr(iLine.FItemArray(0),"-") > 0 then
                	iLine.FItemArray(0) = left(iLine.FItemArray(0),instr(iLine.FItemArray(0),"-")-1)
                end if

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="giftting") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                'CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)



                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
				iLine.FItemArray(1) = Left(iLine.FItemArray(0),4)&"-"&Mid(iLine.FItemArray(0),5,2)&"-"&Mid(iLine.FItemArray(0),7,2)
				iLine.FItemArray(2) = Left(iLine.FItemArray(0),4)&"-"&Mid(iLine.FItemArray(0),5,2)&"-"&Mid(iLine.FItemArray(0),7,2)
				'iLine.FItemArray(2) = "20" & Left(iLine.FItemArray(2),2)&"-"&Mid(iLine.FItemArray(2),4,2)&"-"&Mid(iLine.FItemArray(2),7,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
                   if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(16)="") then iLine.FItemArray(16)="0000"

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gsisuper") then
				iLine.FItemArray(1) = LEFT(iLine.FItemArray(1), 10)
				iLine.FItemArray(2) = LEFT(iLine.FItemArray(2), 10)
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itsbenepia") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itsKaKaoMakers") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itskakaotalkstore") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1), 10))
            	iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2), 10))
 				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))

                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
				'iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))			'�Ǹűݾ��� ���� ���� �ݾ��� ����
				iLine.FItemArray(18) = rtSellPrice
				iLine.FItemArray(12)=Trim(iLine.FItemArray(12))

            	If (isnumeric(iLine.FItemArray(19))) Then
           			iLine.FItemArray(19) = CLNG(iLine.FItemArray(18)) - CLNG(iLine.FItemArray(19))
            	Else
            		iLine.FItemArray(19) = iLine.FItemArray(18)
            	End If

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itskakao") then
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1), 10))
            	iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2), 10))
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
            	iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))			'�Ǹűݾ��� ���� ���� �ݾ��� ����

            	If (isnumeric(iLine.FItemArray(19))) Then
            		If iLine.FItemArray(19) > 0 Then
            			iLine.FItemArray(19) = CLNG((iLine.FItemArray(18) - (iLine.FItemArray(19)) / iLine.FItemArray(17)))
            		End If
            	Else
            		iLine.FItemArray(19) = iLine.FItemArray(18)
            	End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                    	iLine.FItemArray(22) = Trim(Split(iLine.FItemArray(22), ":")(1))
           				iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if


                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			t_addDlvPrice=""
			t_deliverytype=""
			t_sellcash=""
			t_defaultFreeBeasongLimit=""
			t_defaultDeliverPay=""
            if (SellSite="kakaogift") then
				tempShippay = 0
				'2023-07-19 ������ ����..
				iLine.FItemArray(5) = Trim(iLine.FItemArray(5))
				iLine.FItemArray(9) = Trim(iLine.FItemArray(9))
				If Len(iLine.FItemArray(5)) = 0 Then
					iLine.FItemArray(5) = "-"
					iLine.FItemArray(9) = "-"
				End If
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1), 10))
            	iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2), 10))
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				iLine.FItemArray(27) = ""

				rtSellPrice = Clng(iLine.FItemArray(20))
                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��. ��ǰ�ڵ尡 ���εǾ� �ִ°��
					iLine.FItemArray(16)="0000"
                end if

				If instr(iLine.FItemArray(22),":") > 0 Then
					iLine.FItemArray(22) = Trim(Split(iLine.FItemArray(22), ":")(1))
				End If

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
				
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 isNull(k.addDlvPrice, 0) as addDlvPrice, i.deliverytype, i.sellcash, tu.defaultFreeBeasongLimit, tu.defaultDeliverPay " 
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i  " 
				sqlStr = sqlStr & " LEFT JOIN [db_etcmall].[dbo].tbl_kakaoGift_regItem as k on i.itemid = k.itemid " 
				sqlStr = sqlStr & " JOIN [db_user].[dbo].tbl_user_c  as tu on i.makerid= tu.userid " 
				sqlStr = sqlStr & " WHERE i.itemid = '"& Trim(iLine.FItemArray(15)) &"' " 
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					t_addDlvPrice = CLng(rsget("addDlvPrice"))
					t_deliverytype = rsget("deliverytype")
					t_sellcash = CLng(rsget("sellcash"))
					t_defaultFreeBeasongLimit = CLng(rsget("defaultFreeBeasongLimit"))
					t_defaultDeliverPay = CLng(rsget("defaultDeliverPay"))
				Else
					response.write "<script>alert('[��������]���޸�����>>īī������Ʈ�� ����� ��ǰ�� �ƴմϴ�.');</script>"
					response.end
				End If
				rsget.Close

				If t_addDlvPrice > 0 Then
					tempShippay = t_addDlvPrice * iLine.FItemArray(17)
				Else
					If t_deliverytype = "1" and t_sellcash < 30000 Then
						tempShippay = 2500 * iLine.FItemArray(17)
					ElseIf t_deliverytype = "9" and t_sellcash < t_defaultFreeBeasongLimit Then
						tempShippay = t_defaultDeliverPay * iLine.FItemArray(17)
					End If
				End If

				iLine.FItemArray(18) = (rtSellPrice - CLng(tempShippay)) / iLine.FItemArray(17)
				iLine.FItemArray(19) = (rtSellPrice - CLng(tempShippay)) / iLine.FItemArray(17)
				iLine.FItemArray(28) = CLng(tempShippay)
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="GVG") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
'				If chr(asc(iLine.FItemArray(6))) = "?" Then
'					iLine.FItemArray(6) = replace(iLine.FItemArray(6),LEFT(iLine.FItemArray(6),1),"")
'				End If
'				If chr(asc(iLine.FItemArray(7))) = "?" Then
'					iLine.FItemArray(7) = replace(iLine.FItemArray(7),LEFT(iLine.FItemArray(7),1),"")
'				End If
'				If chr(asc(iLine.FItemArray(10))) = "?" Then
'					iLine.FItemArray(10) = replace(iLine.FItemArray(10),LEFT(iLine.FItemArray(10),1),"")
'				End If
'				If chr(asc(iLine.FItemArray(11))) = "?" Then
'					iLine.FItemArray(11) = replace(iLine.FItemArray(11),LEFT(iLine.FItemArray(11),1),"")
'				End If

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="11stITS") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="11st1010") then
            	rtitemoption = iLine.FItemArray(22)
            	beasongNum11st	= iLine.FItemArray(29)
            	Dim vMemo, rtrealSellPrice
	            iLine.FItemArray(1) = LEFT(Replace(iLine.FItemArray(1),"/","-"), 10)
	            iLine.FItemArray(2) = LEFT(Replace(iLine.FItemArray(2),"/","-"), 10)

				rtSellPrice = Clng(iLine.FItemArray(18)) + Clng(Clng(iLine.FItemArray(32)) / iLine.FItemArray(17))		'2017-06-30 ������..�ɼ��߰��ݾ��� �������� ����
				iLine.FItemArray(18) = rtSellPrice

				rtrealSellPrice = rtSellPrice - Clng((Clng(iLine.FItemArray(19))+Clng(iLine.FItemArray(31))) / iLine.FItemArray(17))
				iLine.FItemArray(19) = rtrealSellPrice

				iLine.FItemArray(16) = getOptionCodByOptionName11st(iLine.FItemArray(15), rtitemoption, vMemo)
				If vMemo <> "" then
					iLine.FItemArray(27) = vMemo
				End If
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="lotteimall") then
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
				Dim isOptAddLtimall
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
				sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
				sqlStr = sqlStr & " WHERE convert(varchar(20),itemid) + convert(varchar(20),itemoption) = '"&iLine.FItemArray(15)&"' "
				sqlStr = sqlStr & " and mallid = 'lotteimall' "
				rsget.Open sqlStr,dbget,1
				If (Not rsget.EOF) Then
					isOptAddLtimall = "Y"
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				Else
					isOptAddLtimall = "N"
				End If
				rsget.Close

				'If iLine.FItemArray(0) = "2016-10-22-G97886" Then
				'	iLine.FItemArray(15) = "1254995"
				'	iLine.FItemArray(16) = "Z220"
				'End If

				If isOptAddLtimall = "N" Then
	                if (iLine.FItemArray(16)="") then
	                    if (iLine.FItemArray(22)<>"") then
	                        iLine.FItemArray(16)=getOptionCodByOptionNameimall(iLine.FItemArray(15),iLine.FItemArray(22))
	                     else
	                        iLine.FItemArray(16)="0000"
	                     end if
	                end if
	            End If
                iLine.FItemArray(0) = replace(iLine.FItemArray(0),"-","")
            end if

'            rtitemid=""
'            rtitemoption=""
'            rtSellPrice=""
'            if (SellSite="gseshop") then
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
'                iLine.FItemArray(15) = rtitemid
'                iLine.FItemArray(16) = rtitemoption
'                '//�������� ������
'                'iLine.FItemArray(18) = rtSellPrice
'                'iLine.FItemArray(19) = rtSellPrice
'
'				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
'
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
'                    if (iLine.FItemArray(22)<>"") then
'           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gseshop") then
            	Dim isOptAddGS
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
				sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
				sqlStr = sqlStr & " WHERE convert(varchar(20),itemid) + convert(varchar(20),itemoption) = '"&iLine.FItemArray(15)&"' "
				sqlStr = sqlStr & " and mallid = 'gsshop' "
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					isOptAddGS = "Y"
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				Else
					isOptAddGS = "N"
				End If
				rsget.Close

				If isOptAddGS = "N" Then
	                if (iLine.FItemArray(16)="") then
	                    if (iLine.FItemArray(22)<>"") then
	                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
	                     else
	                        iLine.FItemArray(16)="0000"
	                     end if
	                end if
	            End If

				If iLine.FItemArray(5) <> "" Then
					iLine.FItemArray(5) = LEFT(iLine.FItemArray(5), 20)
				End If

				If iLine.FItemArray(9) <> "" Then
					iLine.FItemArray(9) = LEFT(iLine.FItemArray(9), 20)
				End If

                '''���� ���� ������ 1�� �̻��϶�  ''2014-03-17 �߰�
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                end if

                ''��ۺ� 2014-03-17 �߰�
                if (iLine.FItemArray(28)="����") then
                    iLine.FItemArray(28)=3000
                else
                    iLine.FItemArray(28)=0
                end if

                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

				If Right(iLine.FItemArray(1), 1) = ":" Then
					iLine.FItemArray(1) = iLine.FItemArray(1) & "00"
				End If

'				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
            	iLine.FItemArray(1) = dateconvert(iLine.FItemArray(1))
            	iLine.FItemArray(2) = dateconvert(iLine.FItemArray(1))
            end if

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstorefarm") Then
				Dim isDisCountYn
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'�ǸŰ� = �ǸŰ� + �ɼ��߰��ݾ�

					sqlStr = ""
					sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
					sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
					sqlStr = sqlStr & " WHERE mallgubun = '"& SellSite &"' "
					sqlStr = sqlStr & " and itemid = '"& Trim(iLine.FItemArray(15)) &"' "
					sqlStr = sqlStr & " and '"& LEFT(iLine.FItemArray(2), 10) &"' between startDate and endDate "
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						If rsget("cnt") > 0 Then
							isDisCountYn = "Y"
						Else
							isDisCountYn = "N"
						End If
					rsget.Close

					If isDisCountYn = "Y" Then
						iLine.FItemArray(18) = rtSellPrice - CLng(iLine.FItemArray(30) / iLine.FItemArray(17))
					Else
						iLine.FItemArray(18) = rtSellPrice
					End If
				End If
				iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstoregift") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'�ǸŰ� = �ǸŰ� + �ɼ��߰��ݾ�
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstoremoonbangu") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'�ǸŰ� = �ǸŰ� + �ɼ��߰��ݾ�
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="Mylittlewhoopee") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'�ǸŰ� = �ǸŰ� + �ɼ��߰��ݾ�
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstorefarmclass") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				iLine.FItemArray(16) = getOptionCodByOptionNameClass(iLine.FItemArray(15), iLine.FItemArray(22))
				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'�ǸŰ� = �ǸŰ� + �ɼ��߰��ݾ�
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="auction1010") OR (SellSite="gmarket1010") then
				'���� ������ ������ trimó���� ���� �� �Ǿ��־ ���� trimó��
				iLine.FItemArray(0) = trim(iLine.FItemArray(0))
				iLine.FItemArray(1) = trim(iLine.FItemArray(1))
				iLine.FItemArray(2) = trim(iLine.FItemArray(2))
				iLine.FItemArray(3) = trim(iLine.FItemArray(3))
				iLine.FItemArray(4) = trim(iLine.FItemArray(4))
				iLine.FItemArray(5) = trim(iLine.FItemArray(5))
				iLine.FItemArray(6) = trim(iLine.FItemArray(6))
				iLine.FItemArray(7) = trim(iLine.FItemArray(7))
				iLine.FItemArray(8) = trim(iLine.FItemArray(8))
				iLine.FItemArray(9) = trim(iLine.FItemArray(9))
				iLine.FItemArray(10) = trim(iLine.FItemArray(10))
				iLine.FItemArray(11) = trim(iLine.FItemArray(11))
				iLine.FItemArray(12) = trim(iLine.FItemArray(12))
				iLine.FItemArray(13) = trim(iLine.FItemArray(13))
				iLine.FItemArray(14) = trim(iLine.FItemArray(14))
				iLine.FItemArray(15) = trim(iLine.FItemArray(15))
				iLine.FItemArray(16) = trim(iLine.FItemArray(16))
				iLine.FItemArray(17) = CLNG(trim(iLine.FItemArray(17)))
				iLine.FItemArray(18) = trim(iLine.FItemArray(18))
				iLine.FItemArray(19) = CLNG(trim(iLine.FItemArray(19)))
				iLine.FItemArray(20) = CLNG(trim(iLine.FItemArray(20)))
				iLine.FItemArray(21) = trim(iLine.FItemArray(21))
				iLine.FItemArray(22) = trim(iLine.FItemArray(22))
				iLine.FItemArray(23) = trim(iLine.FItemArray(23))
				iLine.FItemArray(24) = trim(iLine.FItemArray(24))
				iLine.FItemArray(25) = trim(iLine.FItemArray(25))
				iLine.FItemArray(26) = trim(iLine.FItemArray(26))
				iLine.FItemArray(27) = trim(iLine.FItemArray(27))
				iLine.FItemArray(28) = trim(iLine.FItemArray(28))
				iLine.FItemArray(29) = trim(iLine.FItemArray(29))

                iLine.FItemArray(1) = left(iLine.FItemArray(1),10)
                iLine.FItemArray(2) = left(iLine.FItemArray(2),10)

                iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))			'�Ǹűݾ��� ���� ���� �ݾ��� ����
                'iLine.FItemArray(19) ����(realsellprice)�� �ǸŰ� -(������ / ����)�̾�� �ϳ�
				'(�ǸŰ� - ������)/�������� ���Ǿ���..						'2015-10-06 15�� �� ������ ����
				'(�����̿��+���꿹���ݾ�)/���� = �ǰ����ݾ�=realsellprice	'2015-10-07 17:45�� ������ ����
                'iLine.FItemArray(19) = CLNG(iLine.FItemArray(18) - (iLine.FItemArray(19) / iLine.FItemArray(17)))		'ver) �ǸŰ� - (������ / ����)
                iLine.FItemArray(19) = CLNG((iLine.FItemArray(19) + iLine.FItemArray(20)) / iLine.FItemArray(17))		'ver) (�����̿�� + ���꿹���ݾ�) / ����
				If LEFT(iLine.FItemArray(21), 4) = "�ٹ�����" Then
					iLine.FItemArray(21) = Trim(replace(iLine.FItemArray(21), LEFT(iLine.FItemArray(21), 4), ""))
				End If

				If (iLine.FItemArray(23) = "B329873664") OR (iLine.FItemArray(23) = "B291782397")  Then				'Ư�� ��ǰ�̸�
					Dim spItemname
					iLine.FItemArray(16) = "0000"
	         		spItemname = mid(iLine.FItemArray(22),1,instr(iLine.FItemArray(22),"/")-1)
	        		spItemname = mid(spItemname,instr(spItemname,":")+1,100)
					iLine.FItemArray(22) = ""
					If iLine.FItemArray(23) = "B329873664" Then
						Select Case spItemname
							Case "�ֽ�"
								iLine.FItemArray(15) = 1443173
								iLine.FItemArray(21) = "���ǿ��� �ְ������� Ŭ���� - �ֽ�"
							Case "ĵ����ũ"
								iLine.FItemArray(15) = 1444046
								iLine.FItemArray(21) = "���ǿ��� �ְ������� Ŭ���� - ĵ����ũ"
							Case "��ũ�׷���"
								iLine.FItemArray(15) = 1444047
								iLine.FItemArray(21) = "���ǿ��� �ְ������� Ŭ���� - ��ũ�׷���"
						End Select
					ElseIf iLine.FItemArray(23) = "B291782397" Then
						Select Case spItemname
							Case "2016 Ź�� �޷�"
								iLine.FItemArray(15) = 1401873
								iLine.FItemArray(21) = "[�ٹ�����X�����϶�1988] 2016 Ź�� �޷�"
							Case "���� ��ƼĿ"
								iLine.FItemArray(15) = 1401875
								iLine.FItemArray(21) = "[�ٹ�����X�����϶�1988] ���� ��ƼĿ"
							Case "û�� ��Ʈ"
								iLine.FItemArray(15) = 1401877
								iLine.FItemArray(21) = "[�ٹ�����X�����϶�1988] û��ô� ��Ʈ"
						End Select
					End If

	                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
				Else													'���� �ֹ��̸�
	                iLine.FItemArray(16)=getOptionCodByOptionNameAuction(iLine.FItemArray(15),iLine.FItemArray(22), iLine.FItemArray(0))
	                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
					If instr(iLine.FItemArray(22),"�ؽ�Ʈ�� �Է��ϼ���") > 0 Then
'						Dim madeText
'						madeText = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),"�ؽ�Ʈ�� �Է��ϼ���")+1,1000)

'						If instr(madeText, "�ؽ�Ʈ�� �Է��ϼ���") > 0 Then
'							iLine.FItemArray(27) = Trim(Split(madeText, "��")(1))
'						Else
'							iLine.FItemArray(27) = ""
'						End If
						If instr(iLine.FItemArray(22),"�ؽ�Ʈ�� �Է��ϼ��䣺") > 0 Then
							iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(22), "�ؽ�Ʈ�� �Է��ϼ��䣺")(1))
						ElseIf instr(iLine.FItemArray(22),"�ؽ�Ʈ�� �Է��ϼ���:") > 0 Then
							iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(22), "�ؽ�Ʈ�� �Է��ϼ���:")(1))
						End If
					Else
						iLine.FItemArray(27) = ""
					End If
				End If

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if

'                iLine.FItemArray(16)=getOptionCodByOptionNameAuction(iLine.FItemArray(15),iLine.FItemArray(22), iLine.FItemArray(0))
'
'                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
'
'				If instr(iLine.FItemArray(22),"�ؽ�Ʈ�� �Է��ϼ���") > 0 Then
'					Dim madeText
'					madeText = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),"�ؽ�Ʈ�� �Է��ϼ���")-1,1000)
'					If instr(madeText, "�ؽ�Ʈ�� �Է��ϼ���") > 0 Then
'						iLine.FItemArray(27) = Trim(Split(madeText, "��")(1))
'					Else
'						iLine.FItemArray(27) = ""
'					End If
'				Else
'					iLine.FItemArray(27) = ""
'				End If
'
'				if iLine.FItemArray(28) <> "" then
'					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
'				end if
            end if

			rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="ezwel") then
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))    ''�ֹ���

				If LEFT(iLine.FItemArray(22), 3) = "����:" Then
					iLine.FItemArray(22) = replace(iLine.FItemArray(22), LEFT(iLine.FItemArray(22), 3), "")
				End If

				If Right(iLine.FItemArray(22), 1) = "^" Then
					iLine.FItemArray(22) = replace(iLine.FItemArray(22), Right(iLine.FItemArray(22), 1), "")
				End If

'				'2017-07-19 ������ �߰�..realsellprice����..������ �����ʵ� ����..2017-07-24 ������ ������ ���� �Ѿ��;;
'				rtrealSellPrice = Clng(iLine.FItemArray(18)) - Clng(Clng(iLine.FItemArray(19)) / iLine.FItemArray(17))
'				iLine.FItemArray(19) = rtrealSellPrice

				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid FROM db_etcmall.dbo.tbl_ezwel_regitem WHERE ezwelgoodno = '"&iLine.FItemArray(23)&"' "
				rsget.Open sqlStr,dbget,1
				If (Not rsget.EOF) Then
					iLine.FItemArray(15) = rsget("itemid")
				End If
				rsget.Close

                if (iLine.FItemArray(16)="") then
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            If (SellSite="GS25") Then
                tmppreAddr		= Trim(iLine.FItemArray(13))			'�����ּ� ��ü
                tmpnextAddr		= Trim(iLine.FItemArray(14))			'���θ��ּ� ��ü
				If (tmppreAddr = "") AND (tmpnextAddr <> "") Then		'���� ������ �����ּҴ� ���� ���θ� �ּҸ� �ִٸ�
					iLine.FItemArray(13) = tmpnextAddr
					iLine.FItemArray(14) = tmpnextAddr
				ElseIf (tmppreAddr <> "") AND (tmpnextAddr = "") Then	'���� ������ ���θ� �ּҰ� ���� ���� �ּҸ� �ִٸ�
					iLine.FItemArray(13) = tmppreAddr
					iLine.FItemArray(14) = tmppreAddr
				ElseIf (tmppreAddr <> "") AND (tmpnextAddr <> "") Then
					iLine.FItemArray(13) = tmpnextAddr
					iLine.FItemArray(14) = tmpnextAddr
				End If

				'Select Case iLine.FItemArray(23)
				' 	Case "2800100203602"
				' 		iLine.FItemArray(15) = "3313868"
				' 		iLine.FItemArray(18) = "29000"
				' 	Case "2800100204449"
				' 		iLine.FItemArray(15) = "3471382"
				' 		iLine.FItemArray(18) = "45000"
				' 	Case "2800100204456"
				' 		iLine.FItemArray(15) = "4524679"
				' 		iLine.FItemArray(18) = "35000"
				' 	Case "2800100204463"
				' 		iLine.FItemArray(15) = "4890940"
				' 		iLine.FItemArray(18) = "55000"
				' 	Case "2800100204487"
				' 		iLine.FItemArray(15) = "4509495"
				' 		iLine.FItemArray(18) = "35000"
				' 	Case "2800100204494"
				' 		iLine.FItemArray(15) = "4509498"
				' 		iLine.FItemArray(18) = "35000"
				' 	Case "2800100204500"
				' 		iLine.FItemArray(15) = "3504305"
				' 		iLine.FItemArray(18) = "39000"
				' 	Case "2800100204517"
				' 		iLine.FItemArray(15) = "4728736"
				' 		iLine.FItemArray(18) = "29000"
				' End Select
'				iLine.FItemArray(16) = "0000"	'�ɼ��ڵ�

				Select Case iLine.FItemArray(23)
					Case "2800100218279"
						iLine.FItemArray(15) = "4568721"
						iLine.FItemArray(16) = "0015"
						iLine.FItemArray(18) = "32000"
					Case "2800100218286"
						iLine.FItemArray(15) = "4568721"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(18) = "32000"
					Case "2800100218293"
						iLine.FItemArray(15) = "4495213"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "22000"
					Case "2800100218309"
						iLine.FItemArray(15) = "3581268"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "14000"
					Case "2800100218316"
						iLine.FItemArray(15) = "4504295"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "11000"
					Case "2800100218323"
						iLine.FItemArray(15) = "3471386"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "17000"

					Case "2800100218330"
						iLine.FItemArray(15) = "5683080"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "22000"
					Case "2800100218347"
						iLine.FItemArray(15) = "5683121"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "45000"
					Case "2800100218354"
						iLine.FItemArray(15) = "5683124"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "45000"
					Case "2800100218361"
						iLine.FItemArray(15) = "5683125"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "35000"
					Case "2800100218378"
						iLine.FItemArray(15) = "5683126"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "43000"

					'2023-10-23 ������ �ϴ� �ڵ� �߰�
					Case "2840000121644"
						iLine.FItemArray(15) = "4524590"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "22000"
					Case "2840000121651"
						iLine.FItemArray(15) = "4524590"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(18) = "22000"
					Case "2840000121675"
						iLine.FItemArray(15) = "4524590"
						iLine.FItemArray(16) = "0015"
						iLine.FItemArray(18) = "22000"
					Case "2840000121712"
						iLine.FItemArray(15) = "5661015"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "25000"
					Case "2840000121620"
						iLine.FItemArray(15) = "4568721"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "54000"
					Case "2840000121521"
						iLine.FItemArray(15) = "5014914"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(18) = "15000"
					Case "2840000121514"
						iLine.FItemArray(15) = "5014914"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "15000"
					Case "2840000121507"
						iLine.FItemArray(15) = "4546794"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "29000"
					Case "2840000121569"
						iLine.FItemArray(15) = "5616003"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "8000"
					Case "2840000121538"
						iLine.FItemArray(15) = "5616004"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "15000"
					Case "2840000121743"
						iLine.FItemArray(15) = "5109313"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "15000"
					Case "2840000121699"
						iLine.FItemArray(15) = "5109313"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "15000"
					Case "2840000121583"
						iLine.FItemArray(15) = "4958612"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "49000"
					Case "2840000121668"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "12000"
					Case "2840000121637"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(18) = "12000"
					Case "2840000121613"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "12000"
					Case "2840000121590"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(18) = "12000"
					Case "2840000121705"
						iLine.FItemArray(15) = "3581268"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "24000"
				End Select
				iLine.FItemArray(19) = iLine.FItemArray(18)		'���ǸŰ�
				iLine.FItemArray(28) = "0"		'��ۺ�
				iLine.FItemArray(26) = ""		'deliverymemo
            End If

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="homeplus") then
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                if (iLine.FItemArray(16)="") then
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                If iLine.FItemArray(19) < 30000 Then
                	iLine.FItemArray(28)="2500"
                Else
                	iLine.FItemArray(28)="0"
                End If

                '''���� ���� ������ 1�� �̻��϶�  ''2014-03-17 �߰�
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                end if
            end if

'            rtitemid=""
'            rtitemoption=""
'            rtSellPrice=""
'            if (SellSite="cjmall") then
'                if (iLine.FItemArray(16)="") then
'                    if (iLine.FItemArray(22)<>"") then
'                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'				iLine.FItemArray(6) = Replace(iLine.FItemArray(6),")","-")
'				iLine.FItemArray(10)= Replace(iLine.FItemArray(10),")","-")
'				iLine.FItemArray(11)= Replace(iLine.FItemArray(11),")","-")
'
'				Dim cjmAddr1, cjmAddr2
'				cjmAddr1 = Split(iLine.FItemArray(13),"|")(0)
'				cjmAddr2 = Split(iLine.FItemArray(13),"|")(1)
'				iLine.FItemArray(13) = cjmAddr1
'				iLine.FItemArray(14) = cjmAddr2
'                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''�ֹ���
' 				if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)
'            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			IF (sellsite="cjmall") then
				'���������� ���� ��..
                ' If instr(iLine.FItemArray(0),"-") > 0 Then
                ' 	iLine.FItemArray(0) = Split(iLine.FItemArray(0), "-")(0)
                ' End If

				' Dim oTenId
				' oTenId = iLine.FItemArray(15)
				' iLine.FItemArray(15) = Split(oTenId, "_")(0)
				' iLine.FItemArray(16) = Split(oTenId, "_")(1)

				' iLine.FItemArray(6) = Replace(iLine.FItemArray(6),")","-")
				' iLine.FItemArray(10)= Replace(iLine.FItemArray(10),")","-")
				' iLine.FItemArray(11)= Replace(iLine.FItemArray(11),")","-")
                ' iLine.FItemArray(1) = LEFT(Trim(replace(iLine.FItemArray(1),"/","-")), 10)    ''�ֹ���
				' iLine.FItemArray(19) = iLine.FItemArray(19) / iLine.FItemArray(17)

                ' if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)
                ' if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                '     if (iLine.FItemArray(22)<>"") then
                '             iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                '      else
                '         iLine.FItemArray(16)="0000"
                '      end if
                ' end if

				'���������� ���� ��..2023-10-19 ������ ����
				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="NJOYNY") or (SellSite="itsNJOYNY") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="ticketmonster") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="halfclub") then
            	iLine.FItemArray(15) = Split(iLine.FItemArray(23), "_")(0)
                iLine.FItemArray(16) = getOptionCodByOptionNameHalfClub(Split(iLine.FItemArray(23), "_")(0), iLine.FItemArray(21))
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="thinkaboutyou") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

	        	'//�����ȣ ���̿� "-" �� �������
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if


            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			tmpItemname = ""
            if (SellSite="aboutpet") then
				If iLine.FItemArray(16) <> "0000" Then
					If InStr(iLine.FItemArray(21), "_") > 0 Then
						tmpItemname = Trim(Split(iLine.FItemArray(21), "_")(0))
					Else
						tmpItemname = Trim(iLine.FItemArray(21))
					End If
				Else
					tmpItemname = Trim(iLine.FItemArray(21))
				End If
				beasongNum11st	= Replace(iLine.FItemArray(29), ",", "")

				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid "
				sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_aboutpet_regitem "
				sqlStr = sqlStr & " WHERE RTRIM(LTRIM(itemname)) = '"& tmpItemname &"' "
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If (Not rsget.EOF) Then
					iLine.FItemArray(15) = rsget("itemid")
				End If
				rsget.Close

				iLine.FItemArray(19) = Clng(iLine.FItemArray(19)/iLine.FItemArray(17))

	        	'//�����ȣ ���̿� "-" �� �������
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If
            end if


            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="cookatmall") then
				iLine.FItemArray(1) = Left(iLine.FItemArray(0),4) &"-"& mid(iLine.FItemArray(0),5,2) &"-"& mid(iLine.FItemArray(0),7,2)
				Select Case iLine.FItemArray(21)
					Case "�ǳ��� ������ ������_��������"
						iLine.FItemArray(15) = "3649588"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "BR130144"
					Case "�ǳ��� ������ ������_������"
						iLine.FItemArray(15) = "3649588"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "BR130143"
					Case "�ǳ��� �����ǿ� ģ���� �ӱ���_���̳ʽ�"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0016"
						iLine.FItemArray(23) = "BR130142"
					Case "�ǳ��� �����ǿ� ģ���� �ӱ���_����"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0015"
						iLine.FItemArray(23) = "BR130141"
					Case "�ǳ��� �����ǿ� ģ���� �ӱ���_���"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(23) = "BR130140"
					Case "�ǳ��� �����ǿ� ģ���� �ӱ���_��彺Ź"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "BR130139"
					Case "�ǳ��� �����ǿ� ģ���� �ӱ���_��������"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(23) = "BR130138"
					Case "�ǳ��� �����ǿ� ģ���� �ӱ���_������"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "BR130137"
					Case "�ǳ��� ������ ��Ʈ�� �佺�ͱ�"
						iLine.FItemArray(15) = "3471382"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130136"
					Case "�ǳ��� ������ ������ġ&���ø���Ŀ"
						iLine.FItemArray(15) = "2784156"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130135"
					Case "����� ���ǽ��Ʈ Ǫ �ӱ���"
						iLine.FItemArray(15) = "3701715"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130155"
					Case "����� ���� �� Ǫ �ӱ���_Ǫ"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "BR130156"
					Case "����� ���� �� Ǫ �ӱ���_�Ǳ۷�"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "BR130157"
					Case "����� ���� �� Ǫ �ӱ���_Ƽ��"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(23) = "BR130158"
					Case "����� ���� �� Ǫ �ӱ���_�̿丣"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(23) = "BR130159"
					Case "����� ������ Ǫ �ø��� �� (3����Ʈ)"
						iLine.FItemArray(15) = "3646836"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130153"
					Case "����� ������ Ǫ �ø��� �� ��Ʈ (��Ǭ+�ø��� ��)"
						iLine.FItemArray(15) = "3581268"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130154"
					Case "�ǳ��� �������� �ݵ��� ��Ʈ 5ea"
						iLine.FItemArray(15) = "2849183"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130183"
					Case "�긮�� ���ŰƼ ������"
						iLine.FItemArray(15) = "3530000"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130195"
					Case "������ ��Ʈ�� �ż�Ʈ(2����)"
						iLine.FItemArray(15) = "3471386"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "CR130052"
					Case "������ ��Ʈ�� �귱ġ �÷���Ʈ_�����Ǻ극��"
						iLine.FItemArray(15) = "3471393"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "CR130053"
					Case "������ ��Ʈ�� �귱ġ �÷���Ʈ_��彺Ź�극��"
						iLine.FItemArray(15) = "3471393"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "CR130054"
				End Select

	        	'//�����ȣ ���̿� "-" �� �������
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If
            end if


            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="momQ") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

	        	'//�����ȣ ���̿� "-" �� �������
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
'            if (SellSite="privia") then
'				'�ɼǸ�̱�..�ɼ�/�������� ���ͼ� /�� ���ø���
'				iLine.FItemArray(22) = split(iLine.FItemArray(22),"/")(0)
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
'                iLine.FItemArray(15) = rtitemid
'                iLine.FItemArray(16) = rtitemoption
'
'                '//�������� ������
'                'iLine.FItemArray(18) = rtSellPrice
'                'iLine.FItemArray(19) = rtSellPrice
'                '�ֹ��Ͻð� �̻��ϰ� �Ѿ�ͼ� ġȯ
'				iLine.FItemArray(1) = Left(iLine.FItemArray(0),4) &"-"& mid(iLine.FItemArray(0),5,2) &"-"& mid(iLine.FItemArray(0),7,2)
'				'�ֹ��ڸ� �̱�
'				iLine.FItemArray(5) = split(iLine.FItemArray(5),"(")(0)
'
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
'                    if (iLine.FItemArray(22)<>"") then
'           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'            end if
'2013-11-15 16:52 ������ �����غ�
            if (SellSite="privia") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

				'�ֹ��ڸ� �̱�
				If instr(iLine.FItemArray(5),"(") > 0 Then
					iLine.FItemArray(5) = split(iLine.FItemArray(5),"(")(0)
				End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="momastore") then
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

				'�ֹ��ڸ� �̱�
				If instr(iLine.FItemArray(5),"(") > 0 Then
					iLine.FItemArray(5) = split(iLine.FItemArray(5),"(")(0)
				End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

       		rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="its29cm") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/�̴ϼȼ����ϰ�� ����ڰ� �ǸŰ����� �ȳִ� ��쿡 �Һ��ڰ��� ��ü
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gabangpop") or (SellSite="itsGabangpop") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/�̴ϼȼ����ϰ�� ����ڰ� �ǸŰ����� �ȳִ� ��쿡 �Һ��ڰ��� ��ü
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="musinsaITS") or (SellSite="itsMusinsa") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(22) = replace(iLine.FItemArray(22),"NONE","")
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="celectory") then
                iLine.FItemArray(1) = replace(iLine.FItemArray(1),"'","")
                iLine.FItemArray(6) = replace(iLine.FItemArray(6),"'","")
                iLine.FItemArray(7) = replace(iLine.FItemArray(7),"'","")
                iLine.FItemArray(10) = replace(iLine.FItemArray(10),"'","")
                iLine.FItemArray(11) = replace(iLine.FItemArray(11),"'","")
                iLine.FItemArray(12) = replace(iLine.FItemArray(12),"'","")
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(22) = Trim(iLine.FItemArray(22))
                iLine.FItemArray(23) = Trim(iLine.FItemArray(23))

                CALL getEtcSiteNameOrCode2ItemCode(sellsite, iLine.FItemArray(23), iLine.FItemArray(21), iLine.FItemArray(22), rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="player") or (SellSite="itsPlayer1") then
            	'//��ǰ�� �ɼǸ��� ���� ���ֽ�.. ��ġ ����ؼ� ©��
            	if instr(iLine.FItemArray(21),"(") > 0 then
            		iLine.FItemArray(21) = left(iLine.FItemArray(21), instr(iLine.FItemArray(21),"(")-1 )
            		iLine.FItemArray(22) = rtrim(replace(mid(iLine.FItemArray(22), instr(iLine.FItemArray(22),"(")+1 , 96 ),")",""))

            	'//�ɼǾ���
            	else
            		iLine.FItemArray(21) = iLine.FItemArray(21)
            		iLine.FItemArray(22) = ""
            	end if


                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

				'//�Ǹ����� �ȳѿͿͼ�, �ֹ���ȣ ���ڸ� 10�ڸ��� �Ǹ��Ϸ� ó��
                iLine.FItemArray(1) = Left(iLine.FItemArray(0),4) &"-"& mid(iLine.FItemArray(0),5,2) &"-"& mid(iLine.FItemArray(0),7,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            if (SellSite="wconcept") or (SellSite="itsWconcept") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/�̴ϼȼ����ϰ�� ����ڰ� �ǸŰ����� �ȳִ� ��쿡 �Һ��ڰ��� ��ü
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

			If (SellSite="hmall1010") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If


			If (SellSite="lotteon") Then
				iLine.FItemArray(25) = Split(iLine.FItemArray(25), "_")(1)
				If instr(iLine.FItemArray(27),"�ؽ�Ʈ�� �Է��ϼ��� :") > 0 Then
					iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(27), "�ؽ�Ʈ�� �Է��ϼ��� :")(1))
				Else
					iLine.FItemArray(27) = ""
				End If

				sqlStr = ""
				sqlStr = sqlStr & " select top 1 itemid, itemoption "
				sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption "
				sqlStr = sqlStr & " where outmallOptCode = '"& iLine.FItemArray(24) &"' "
				sqlStr = sqlStr & " and mallid = 'lotteon' "
				rw sqlStr
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If (Not rsget.EOF) Then
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				End If
				rsget.Close

				if iLine.FItemArray(18) * iLine.FItemArray(17)  > 50000 then
					iLine.FItemArray(28) = 0
				else
					iLine.FItemArray(28) = 3000
				end if
			' 	isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
			' 	If isValid = "N" Then
			' 		response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
			' 		response.end
			' 	End If
			 	iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="shintvshopping") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="skstoa") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="LFmall") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="goodwearmall10") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="wconcept1010") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="itsWadiz") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''�ɼ��ڵ尡 ��.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 ��ǰ �ڵ�� 10x10 �ɼ��ڵ尡 ��Ī�� �� ���� �´� �� �ٽ� Ȯ���ϼ���');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

            '' CASE
            if (SellSite="hiphoper") or (SellSite="itsHiphoper") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(22) = replace(replace(Trim(replace(iLine.FItemArray(22),iLine.FItemArray(21),"")),"(",""),")","")
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                '//������ �������� �Ѿ�� �������� ������
                if iLine.FItemArray(17) > 1 then
                	iLine.FItemArray(18) = iLine.FItemArray(18)/iLine.FItemArray(17)
                	iLine.FItemArray(19) = iLine.FItemArray(19)/iLine.FItemArray(17)
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="ssg") then
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				iLine.FItemArray(1) = left(Replace(iLine.FItemArray(1),".","-"),10)

				If Len(iLine.FItemArray(1)) = "8" Then
					iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),5,2)&"-"&Mid(iLine.FItemArray(1),7,2)
				End If
				iLine.FItemArray(22) = replace(iLine.FItemArray(22),"NONE","")

				If instr(iLine.FItemArray(22),"�ֹ����۹���:") > 0 Then
					iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(22), "�ֹ����۹���:")(1))
				Else
					iLine.FItemArray(27) = ""
				End If

				beasongNum11st	= iLine.FItemArray(29)
				reserve01 = iLine.FItemArray(38)
				outmalloptionno = iLine.FItemArray(24)
				If outmalloptionno = "" Then
					outmalloptionno = "00000"
				End If

				if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
					if (iLine.FItemArray(22)<>"") then
						iLine.FItemArray(16)=getOptionCodByOptionNameSSG(iLine.FItemArray(15),iLine.FItemArray(22))
					else
						iLine.FItemArray(16)="0000"
					end if
				end if
            end if

            '' CASE
'            If (SellSite="coupang") then
'            	Dim tmpOptName
'            	tmpOptName = Trim(Replace(iLine.FItemArray(21),"����1) Ǫġ�ٺ� ������ 5 5S ���̽� ",""))
'                iLine.FItemArray(21) = "Puchibabie Ǫġ�ٺ�_iPhone5/5S Case"
'                Select Case tmpOptName
'                	Case "����_����Ʈ�׷���"
'                		iLine.FItemArray(22) = "01_Soft Grey"
'                		iLine.FItemArray(16) = "0011"
'                	Case "����_����Ʈ��ũ"
'                		iLine.FItemArray(22) = "02_Soft Pink"
'                		iLine.FItemArray(16) = "0012"
'                	Case "����_����Ʈ��Ʈ"
'                		iLine.FItemArray(22) = "03_Soft Mint"
'                		iLine.FItemArray(16) = "0013"
'                	Case "����_����ũ"
'                		iLine.FItemArray(22) = "04_Hot Pink"
'                		iLine.FItemArray(16) = "0014"
'                	Case "����_�˺��"
'                		iLine.FItemArray(22) = "05_Pop Blue"
'                		iLine.FItemArray(16) = "0015"
'                	Case "����_�˽�Ʈ������"
'                		iLine.FItemArray(22) = "06_Pop Stripe"
'                		iLine.FItemArray(16) = "0016"
'                	Case "����_����Ʈ����"
'                		iLine.FItemArray(22) = "07_Sweet Fruit"
'                		iLine.FItemArray(16) = "0017"
'                	Case "����_�Ľ��ں��"
'                		iLine.FItemArray(22) = "08_Pastel Blue"
'                		iLine.FItemArray(16) = "0018"
'                	Case "����_�׸������"
'                		iLine.FItemArray(22) = "09_Green Rose Dot"
'                		iLine.FItemArray(16) = "0019"
'                	Case "����_ȭ��Ʈ����ġ"
'                		iLine.FItemArray(22) = "10_White Sketch"
'                		iLine.FItemArray(16) = "0020"
'                End Select
'                iLine.FItemArray(15) = "783540"
'                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
'				iLine.FItemArray(18) = 13800
'				iLine.FItemArray(19) = 13800
'            end if

			If (SellSite="coupang") then
                ' iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                ' iLine.FItemArray(1) = left(Replace(iLine.FItemArray(1),".","-"),10)
				' iLine.FItemArray(22) = replace(iLine.FItemArray(22),"NONE","")
                ' CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                ' iLine.FItemArray(15) = rtitemid
                ' iLine.FItemArray(16) = rtitemoption
                ' '//�������� ������
                ' 'iLine.FItemArray(18) = rtSellPrice
                ' 'iLine.FItemArray(19) = rtSellPrice

                ' if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                '     if (iLine.FItemArray(22)<>"") then
           		' 		temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                '      else
                '         iLine.FItemArray(16)="0000"
                '      end if
                ' end if
				Dim spItemId, spItemOption
                spItemId        = Split(iLine.FItemArray(15), "_")(0)
                spItemOption    = Split(iLine.FItemArray(15), "_")(1)

            	iLine.FItemArray(15) = spItemId
				iLine.FItemArray(16) = spItemOption
				beasongNum11st	= iLine.FItemArray(29)
				outmalloptionno = iLine.FItemArray(27)
				iLine.FItemArray(27) = ""

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
			end if

            if (SellSite="ollehtv") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

                iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))      '''������ ������ �ܰ��� ����.
                iLine.FItemArray(19) = iLine.FItemArray(18)
                '//�������� ������
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''�ɼ��ڵ尡 ��.
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="wemakeprice")  then
				Select Case Trim(iLine.FItemArray(24))
					Case "230200538"	iLine.FItemArray(16) = "0011"
					Case "230200539"	iLine.FItemArray(16) = "0012"
					Case "230200540"	iLine.FItemArray(16) = "0011"
					Case "230200541"	iLine.FItemArray(16) = "0012"
					Case "230200542"	iLine.FItemArray(16) = "0011"
					Case "230200543"	iLine.FItemArray(16) = "0012"
					Case "230200544"	iLine.FItemArray(16) = "0011"
					Case "230200545"	iLine.FItemArray(16) = "0012"
					Case "230200548"	iLine.FItemArray(16) = "0011"
					Case "230200549"	iLine.FItemArray(16) = "0012"
					Case "230200550"	iLine.FItemArray(16) = "0011"
					Case "230200551"	iLine.FItemArray(16) = "0012"
					Case "230200552"	iLine.FItemArray(16) = "0013"
					Case "230200553"	iLine.FItemArray(16) = "0014"
					Case "230200556"	iLine.FItemArray(16) = "0011"
					Case "230200557"	iLine.FItemArray(16) = "0012"
					Case "230200558"	iLine.FItemArray(16) = "0013"
					Case "230200559"	iLine.FItemArray(16) = "0014"
					Case "230200560"	iLine.FItemArray(16) = "0011"
					Case "230200561"	iLine.FItemArray(16) = "0012"
					Case "230200562"	iLine.FItemArray(16) = "0014"

					Case "222648154"	iLine.FItemArray(16) = "0011"
					Case "222648155"	iLine.FItemArray(16) = "0012"
					Case "222648156"	iLine.FItemArray(16) = "0013"
					Case "222648157"	iLine.FItemArray(16) = "0014"
					Case "222648158"	iLine.FItemArray(16) = "0011"
					Case "222648159"	iLine.FItemArray(16) = "0012"
					Case "222648160"	iLine.FItemArray(16) = "0013"
					Case "222648161"	iLine.FItemArray(16) = "0014"

					Case "235610341"	iLine.FItemArray(16) = "0011"
					Case "235610342"	iLine.FItemArray(16) = "0012"
					Case "235610343"	iLine.FItemArray(16) = "0013"
					Case "235610344"	iLine.FItemArray(16) = "0014"
					Case "235610345"	iLine.FItemArray(16) = "0015"
					Case "235610346"	iLine.FItemArray(16) = "0016"
					Case "235610335"	iLine.FItemArray(16) = "0011"
					Case "235610336"	iLine.FItemArray(16) = "0012"
					Case "235610337"	iLine.FItemArray(16) = "0013"
					Case "235610338"	iLine.FItemArray(16) = "0014"
					Case "235610339"	iLine.FItemArray(16) = "0015"
					Case "235610340"	iLine.FItemArray(16) = "0016"
					Case "235554716"	iLine.FItemArray(16) = "0011"
					Case "235554717"	iLine.FItemArray(16) = "0012"
					Case "235554718"	iLine.FItemArray(16) = "0013"
					Case "235554719"	iLine.FItemArray(16) = "0015"
					Case "235554720"	iLine.FItemArray(16) = "0016"
					Case "235554714"	iLine.FItemArray(16) = "0011"
					Case "235554715"	iLine.FItemArray(16) = "0012"
					Case "235554709"	iLine.FItemArray(16) = "0011"
					Case "235554710"	iLine.FItemArray(16) = "0012"
					Case "235554711"	iLine.FItemArray(16) = "0013"
					Case "235554712"	iLine.FItemArray(16) = "0014"
					Case "235554713"	iLine.FItemArray(16) = "0015"

					Case "237893189"	iLine.FItemArray(16) = "0011"
					Case "237893190"	iLine.FItemArray(16) = "0012"
					Case "237893191"	iLine.FItemArray(16) = "0013"
					Case "237893192"	iLine.FItemArray(16) = "0014"
					Case "237893193"	iLine.FItemArray(16) = "0015"
					Case "237893194"	iLine.FItemArray(16) = "0016"
					Case "237893195"	iLine.FItemArray(16) = "0011"
					Case "237893196"	iLine.FItemArray(16) = "0012"
					Case "237893197"	iLine.FItemArray(16) = "0013"
					Case "237893198"	iLine.FItemArray(16) = "0014"
					Case "237893199"	iLine.FItemArray(16) = "0015"
					Case "237893200"	iLine.FItemArray(16) = "0016"
					Case Else			iLine.FItemArray(16)="0000"
				End Select

'�������� 9700�� �̻��̸� ������ / ���� �ϴ� ���� 3�����̻��̸� ������
'				if (iLine.FItemArray(18)>=9700) Then
'					iLine.FItemArray(28)="0"
'				else
'					iLine.FItemArray(28)="2500"
'				end if

				If (iLine.FItemArray(18)>=30000) Then
					iLine.FItemArray(28)="0"
				Else
					iLine.FItemArray(28)="2500"
				End If

                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))      '''������ ������ �ܰ��� ����.
                    end if
				end if
				iLine.FItemArray(19) = iLine.FItemArray(18)
            end if
'            IF (sellsite="wemakeprice") then
'                if (Left(iLine.FItemArray(21),Len("[��ǰ���� : ���� �ƺ� �ڵ����� ��Ʈ"))="[��ǰ���� : ���� �ƺ� �ڵ����� ��Ʈ") then
'                    iLine.FItemArray(15) = 475219
'                    iLine.FItemArray(21) = "���� �ƺ� �ڵ����� ��Ʈ"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 4400
'                    iLine.FItemArray(19) = 4400
'                elseif (Left(iLine.FItemArray(21),Len("[��ǰ���� : ī���̼� ���� ��Ʈ"))="[��ǰ���� : ī���̼� ���� ��Ʈ") then
'                    iLine.FItemArray(15) = 475372
'                    iLine.FItemArray(21) = "ī���̼� ���� ��Ʈ"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 4350
'                    iLine.FItemArray(19) = 4350
'                elseif (Left(iLine.FItemArray(21),Len("[��ǰ���� : ī���̼� �ڵ�����(����)"))="[��ǰ���� : ī���̼� �ڵ�����(����)") then
'                    iLine.FItemArray(15) = 475218
'                    iLine.FItemArray(21) = "ī���̼� �ڵ�����"
'                    iLine.FItemArray(16) = "0011"
'                    iLine.FItemArray(22) = "����"
'                    iLine.FItemArray(18) = 4000
'                    iLine.FItemArray(19) = 4000
'                elseif (Left(iLine.FItemArray(21),Len("[��ǰ���� : ī���̼� �ڵ�����(��ũ)"))="[��ǰ���� : ī���̼� �ڵ�����(��ũ)") then
'                    iLine.FItemArray(15) = 475218
'                    iLine.FItemArray(21) = "ī���̼� �ڵ�����"
'                    iLine.FItemArray(16) = "0012"
'                    iLine.FItemArray(22) = "��ũ"
'                    iLine.FItemArray(18) = 4000
'                    iLine.FItemArray(19) = 4000
'                elseif (Left(iLine.FItemArray(21),Len("[��ǰ���� : ũ����Ż ī���̼� ���ġ"))="[��ǰ���� : ũ����Ż ī���̼� ���ġ") then
'                    iLine.FItemArray(15) = 475457
'                    iLine.FItemArray(21) = "ũ����Ż ī���̼� ���ġ"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 6750
'                    iLine.FItemArray(19) = 6750
'                elseif (Left(iLine.FItemArray(21),Len("[��ǰ���� : ���� ī���̼� ���ġ"))="[��ǰ���� : ���� ī���̼� ���ġ") then
'                    iLine.FItemArray(15) = 475459
'                    iLine.FItemArray(21) = "���� ī���̼� ���ġ"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 4250
'                    iLine.FItemArray(19) = 4250
'                end if
'            end if

            IF (sellsite="lotteCom") then
            	Dim isOptAddLotteCom
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
				sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m "
				sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteAddOption_regItem as r on m.idx = r.midx "
				sqlStr = sqlStr & " WHERE IsNULL(r.LotteGoodNo, r.LotteTmpGoodNo)= '"&iLine.FItemArray(23)&"' "
				sqlStr = sqlStr & " and m.mallid = 'lotteCom' "
				rsget.Open sqlStr,dbget,1
				If (Not rsget.EOF) Then
					isOptAddLotteCom = "Y"
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				Else
					isOptAddLotteCom = "N"
				End If
				rsget.Close

				If isOptAddLotteCom = "N" Then
	                if (iLine.FItemArray(15)="") and (iLine.FItemArray(23)<>"") then
	                    iLine.FItemArray(15) = getItemIDByUpcheItemCode(sellsite,iLine.FItemArray(23))
	                end if

	                if (iLine.FItemArray(16)="") then
	                    if (iLine.FItemArray(22)<>"") then
	                        iLine.FItemArray(16)=getOptionCodByOptionNameLotte(iLine.FItemArray(15),iLine.FItemArray(22))
	                     else
	                        iLine.FItemArray(16)="0000"
	                     end if
	                end if
				End If

				iLine.FItemArray(0) = replace(iLine.FItemArray(0),"-","")				'2014-12-31 ������ �߰�
                if (iLine.FItemArray(26)="") then
                    iLine.FItemArray(26) = iLine.FItemArray(29)
                else
                    iLine.FItemArray(26) = iLine.FItemArray(26)&VBCRLF&iLine.FItemArray(29)
                end if

                if (iLine.FItemArray(32)="��ȯ�ֹ�") then
                    iLine.FItemArray(32)=3
                elseif (iLine.FItemArray(32)<>"�ֹ�") then
                    iLine.FItemArray(32)=9
                else
                    iLine.FItemArray(32)=0
                end if
            END IF

            IF (sellsite="dnshop") then
                if (iLine.FItemArray(16)="") then
                     if (iLine.FItemArray(22)<>"") and (iLine.FItemArray(22)<>"��ǰ����") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))

                        ''if (iLine.FItemArray(16)="") then iLine.FItemArray(16)="0000"
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

				'2015-03-12 ������ �ϴ� �ּ�ó��, �ٲ� �������� ������ȣ�� ����
				''������ȣ4 ������.
				'if (iLine.FItemArray(29)="4") then
				'   iLine.FItemArray(28) = -1
				'end if

            END IF

            IF (sellsite="interpark") then
                ''selldate
                iLine.FItemArray(1) = Left(iLine.FItemArray(0),4)&"-"&Mid(iLine.FItemArray(0),5,2)&"-"&Mid(iLine.FItemArray(0),7,2)

                ''�ֹ��� ID�� ��������
                iLine.FItemArray(5) = ReplaceText(iLine.FItemArray(5),"(\()[\s\S]*(\))","")

                ''rw iLine.FItemArray(15)&"|"&iLine.FItemArray(16)
                if (iLine.FItemArray(15)=iLine.FItemArray(16)) then
                    iLine.FItemArray(16)="0000"
                end if

                '''�ɼ��� ���� ���̽��� �̻��� ���̽�..***
                if (iLine.FItemArray(16)="") then
                    IF (Trim(iLine.FItemArray(22))<>"") then
                        POS1 = InStr(iLine.FItemArray(22),"/")
                        bufOptionName = ""
                        IF (POS1>0) THEN bufOptionName=Mid(iLine.FItemArray(22),POS1+1,255)
                        bufOptionName = Trim(bufOptionName)
                        iLine.FItemArray(16) = getOptionCodByOptionName(iLine.FItemArray(15),bufOptionName)

                        'if iLine.FItemArray(16)="" then iLine.FItemArray(16)="0000"  ''' �ɼ��� ���� �ȵǸ� �ȵ�����.. 2012/03/02
                        if iLine.FItemArray(16)="" then iLine.FItemArray(16)="0000" ''�ϴ� ���� �ֹ��Է½� �����ϰԲ�.
                    else
                        iLine.FItemArray(16)="0000"
                    end if
                end if

                ''�ֹ����۹��� �߰� : 2012-09-14
                POS1 = InStr(iLine.FItemArray(22),"| �ֹ����۹���")
                IF (POS1<1) then
                    POS1 = InStr(iLine.FItemArray(22),"|�ֹ����۹���")
                end if

                if (POS1>0) then
                    if (iLine.FItemArray(16) = "") then iLine.FItemArray(16) = "0000" ''20121219�߰� ''|�ֹ����۹���/tea party for two/ all that razz/ groovy grape

                    POS2 = InStr(Mid(iLine.FItemArray(22),pos1,512),"/")
                    if (POS2>0) then '' 27 :: �ֹ����۹���
                        iLine.FItemArray(27) = Trim(Mid(iLine.FItemArray(22),pos1+pos2,512))
                    end if
                end if


                '''383048 �̻���.. // �ɼ��ڵ� �ʿ� ���̰� �ɼǸ� ���� / French Lilac" �ɼǱ��� / �ɼǸ� �̷������� �ö���� ��� ����.
                '''256712
                if (iLine.FItemArray(15)="383048") and ((iLine.FItemArray(16)="0000") or (iLine.FItemArray(16)="")) then
                    if (Trim(iLine.FItemArray(22))="�ɼ� / ������ | ����2 / Cobalt Blue") then
                        iLine.FItemArray(16)="0015"
                    elseif (Trim(iLine.FItemArray(22))="�ɼ� / ������ | ����2 / Ivory") then
                        iLine.FItemArray(16)="0014"
                    elseif (Trim(iLine.FItemArray(22))="�ɼ� / ������ | ����2 / Black") then
                        iLine.FItemArray(16)="0013"
                    elseif (Trim(iLine.FItemArray(22))="�ɼ� / ������ | ����2 / Orange Red") then
                        iLine.FItemArray(16)="0012"
                    elseif (Trim(iLine.FItemArray(22))="�ɼ� / ������ | ����2 / Brown") then
                        iLine.FItemArray(16)="0011"
                    end if
                end if

                if (iLine.FItemArray(15)="256712") and ((iLine.FItemArray(16)="0000") or (iLine.FItemArray(16)="")) then
                    if (Trim(iLine.FItemArray(22))="����1 / ȭ��Ʈ+������Ʈ���þ���") then
                        iLine.FItemArray(16)="Z310"
                    elseif (Trim(iLine.FItemArray(22))="����1 / ���߷�+������Ʈ���þ���") then
                        iLine.FItemArray(16)="Z210"
                    elseif (Trim(iLine.FItemArray(22))="����1 / ��+������Ʈ���þ���") then
                        iLine.FItemArray(16)="Z110"
                    end if
                end if
            End IF

            ''rw "@partnerItemID="&iLine.FItemArray(15)
            ''rw "@partnerItemName="&iLine.FItemArray(21)
            ''rw "@partnerOption="&iLine.FItemArray(16)
            ''rw "@partnerOptionName="&iLine.FItemArray(22)
            ''rw "@SellPrice="&iLine.FItemArray(18)
            ''rw "@RealSellPrice="&iLine.FItemArray(19)

            ''�ɼǰ��� �ִ°��.

            iLine.FItemArray(17) = Replace(iLine.FItemArray(17),",","")
            iLine.FItemArray(18) = Replace(iLine.FItemArray(18),",","")
            iLine.FItemArray(19) = Replace(iLine.FItemArray(19),",","")



            if (sellsite="dnshop") then '' 2014/01/15 interpark �߰�
                ''�Ǹ��� �ǸŰ�-���ΰ� ���� ���� 2014/03/10--------------------------- ����ݾ��� �ش� ��ǰ�ݾ� �հ谡 �ƴѵ� ��.
                if iLine.FItemArray(30)<>"" then
                    iLine.FItemArray(19) = iLine.FItemArray(18)-iLine.FItemArray(30)
                end if
                ''--------------------------------------------------------------------

            '''���� ���� ������ 1�� �̻��϶�  ''2011-06-29 �߰�
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                else
                    response.write "."&iLine.FItemArray(17)
                end if
            end if

            if (sellsite="interpark") then
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                else
                    response.write "."&iLine.FItemArray(19)
                end if
            end if

			if (SellSite="lotteCom") then
'	            IF (UBound(iLine.FItemArray)>30) and (iLine.FItemArray(18)<>"") then
'	                IF (iLine.FItemArray(30)="") then iLine.FItemArray(30)="0"
'	                IF (iLine.FItemArray(31)="") then iLine.FItemArray(31)="0"
'
'	                iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)) + CLNG(iLine.FItemArray(30)) ''?? �ּ�ó�� 2013/10 ������
'	            END IF
			end if

            iLine.FItemArray(3) = convPayTypeStr2Code(iLine.FItemArray(3))
            IF (iLine.FItemArray(3)="") then iLine.FItemArray(3)="50"                           ''PayType
            IF (iLine.FItemArray(2)="") then iLine.FItemArray(2)=iLine.FItemArray(1)            ''Paydate
            IF (iLine.FItemArray(19)="") then iLine.FItemArray(19)=iLine.FItemArray(18)         ''RealSellPrice

            '''�����ȣ�� �ּ�1�� ������� [ ]�� �����ȣ ���� = �ϳ�������.
            iLine.FItemArray(12) = TRIM(Replace(iLine.FItemArray(12),"  "," "))
            iLine.FItemArray(13) = TRIM(Replace(iLine.FItemArray(13),"  "," "))
            iLine.FItemArray(14) = TRIM(Replace(iLine.FItemArray(14),"  "," "))
            IF (iLine.FItemArray(12)=iLine.FItemArray(13)) then
                POS1 = InStr(iLine.FItemArray(12),"[")
                POS2 = InStr(iLine.FItemArray(12),"]")
                IF (POS1>0) and (POS2>0) then
                    iLine.FItemArray(12) = Mid(iLine.FItemArray(12),POS1+1,POS2-POS1-1)
                    iLine.FItemArray(12) = Trim (iLine.FItemArray(12))

                    IF (iLine.FItemArray(13)=iLine.FItemArray(14)) THEN
                        iLine.FItemArray(13) = TRIM(Mid(iLine.FItemArray(13),POS2+1,512))
                        iLine.FItemArray(14) = iLine.FItemArray(13)
                    ELSE
                        iLine.FItemArray(13) = TRIM(Mid(iLine.FItemArray(13),POS2+1,512))
                    END IF
                END IF
            END IF

            '''�ּҿ� ���ּҰ� ������� 3��° Blank���� ����.
            POS1 = 0
            POS2 = 0
            POS3 = 0
            IF (iLine.FItemArray(13)=iLine.FItemArray(14)) then
                POS1 = InStr(iLine.FItemArray(14)," ")
                ''rw "POS1="&POS1
                IF (POS1>0) then
                    POS2 = InStr(MID(iLine.FItemArray(14),POS1+1,512)," ")
                    ''rw "POS2="&POS2
                    IF POS2>0 then
                        POS3 = InStr(MID(iLine.FItemArray(14),POS1+POS2+1,512)," ")
                        IF POS3>0 then
                            iLine.FItemArray(13)=LEFT(iLine.FItemArray(14),POS1+POS2+POS3-1)
                            iLine.FItemArray(14)=MID(iLine.FItemArray(14),POS1+POS2+POS3+1,512)

                            'rw iLine.FItemArray(13)
                            'rw iLine.FItemArray(14)
                        END IF
                    END IF
                END IF
            END IF

			dim countryCode
			if (SellSite="cn10x10") or (SellSite="cnglob10x10") or (SellSite="cnhigo")  or (SellSite = "11stmy") or (SellSite = "cnugoshop") or (SellSite = "zilingo") or (SellSite = "etsy") then
				countryCode = iLine.FItemArray(33)
			end if

			if ucase(countryCode)="" then countryCode="KR"

			If Sellsite = "cnglob10x10" Then

			End If

			Dim replaceMemo
			If iLine.FItemArray(26) <> "" Then
				replaceMemo = Replace(iLine.FItemArray(26), "&amp;", "&")
				replaceMemo = Replace(replaceMemo, "amp;", "&")
				replaceMemo = Replace(replaceMemo, "&nbsp;", " ")
				replaceMemo = Replace(replaceMemo, "nbsp;", " ")
				replaceMemo = Replace(replaceMemo, "&lt;", "<")
				replaceMemo = Replace(replaceMemo, "lt;", "<")
				replaceMemo = Replace(replaceMemo, "&gt;", ">")
				replaceMemo = Replace(replaceMemo, "gt;", ">")
				replaceMemo = Replace(replaceMemo, "&quot;", """")
				replaceMemo = Replace(replaceMemo, "quot;", """")
				iLine.FItemArray(26) = replaceMemo
			End If

			'�ֹ��� / ������(�����̸�Case) ���� �涧 ������ Left ���ڿ� ó��
			If iLine.FItemArray(5) <> "" Then
				iLine.FItemArray(5) = LEFT(iLine.FItemArray(5), 28)
			End If

			If iLine.FItemArray(15) <> "" Then
				iLine.FItemArray(15) = LEFT(iLine.FItemArray(15), 28)
			End If

IF (application("Svr_Info")	= "Dev") or C_ADMIN_AUTH then
    rw "@SellSite="&SellSite
    rw "@OutMallOrderSerial="&iLine.FItemArray(0)
    rw "@SellDate="&iLine.FItemArray(1)
    rw "@PayType="&iLine.FItemArray(3)
    rw "@Paydate="&iLine.FItemArray(2)
    rw "@partnerItemID="&iLine.FItemArray(15)
    rw "@partnerItemName="&iLine.FItemArray(21)
    rw "@partnerOption="&iLine.FItemArray(16)
    rw "@partnerOptionName="&iLine.FItemArray(22)
    rw "@OrderUserID="&iLine.FItemArray(4)

    rw "@OrderName="&iLine.FItemArray(5)
    rw "@OrderEmail="&iLine.FItemArray(8)
    rw "@OrderTelNo="&iLine.FItemArray(6)
    rw "@OrderHpNo="&iLine.FItemArray(7)

    rw "@ReceiveName="&iLine.FItemArray(9)
    rw "@ReceiveTelNo="&iLine.FItemArray(10)
    rw "@ReceiveHpNo="&iLine.FItemArray(11)
    rw "@ReceiveZipCode="&iLine.FItemArray(12)
    rw "@ReceiveAddr1="&iLine.FItemArray(13)
    rw "@ReceiveAddr2="&iLine.FItemArray(14)

    rw "@SellPrice="&iLine.FItemArray(18)
    rw "@RealSellPrice="&iLine.FItemArray(19)
    rw "@ItemOrderCount="&iLine.FItemArray(17)
    rw "@OrgDetailKey="&iLine.FItemArray(25)

    rw "@deliverymemo="&iLine.FItemArray(26)
    rw "@requireDetail="&iLine.FItemArray(27)

    rw "@orderDlvPay="&iLine.FItemArray(28)
    if UBound(iLine.FItemArray)>=29 then
        rw "@etc1="&iLine.FItemArray(29)
    end if
    rw "@countryCode="&countryCode

    rw "@outMallGoodsNo="&iLine.FItemArray(23)
    IF (SellSite="shoplinker") THEN
        rw "@etc2(shoplinkermallname)="&iLine.FItemArray(30)
        rw "@etc3(shoplinkerPrdCode)="&iLine.FItemArray(31)
        rw "@etc4(shoplinkerOrderID)="&iLine.FItemArray(32)
        rw "@etc4(shoplinkerMallid)="&iLine.FItemArray(33)
    ENd IF
    rw "@overseasPrice="&overseasPrice
    rw "@overseasDeliveryPrice="&overseasDeliveryPrice
    rw "@overseasRealPrice="&overseasRealPrice
    rw "@reserve01="&reserve01
    rw "@beasongNum11st="&beasongNum11st
	rw "@outmalloptionno="&outmalloptionno
    rw "------------------------------------------------"
	' response.end
ENd IF

        IF (iLine.FItemArray(0)<>"") and (iLine.FItemArray(0)<>"20110430-927718") then
            IF (sellsite="lotteCom") then
                orderCsGbn = iLine.FItemArray(32)
            ELSEIF (sellsite="gseshop") then
                orderCsGbn = iLine.FItemArray(32)
				if (orderCsGbn <> "�ֹ�") then
					orderCsGbn = "3"
				else
					orderCsGbn = "0"
				end if
            ELSE
                orderCsGbn = "0"
            end if

            IF (SellSite="shoplinker") THEN  ''2013/09/16 �߰� ����Ŀ����
                outMallGoodsNo=iLine.FItemArray(23)
                shoplinkermallname=iLine.FItemArray(30)
                shoplinkerPrdCode=iLine.FItemArray(31)
                shoplinkerOrderID=iLine.FItemArray(32)
                shoplinkerMallID =iLine.FItemArray(33)
'

               '' rw shoplinkermallname&":"&shoplinkerMallID
            ELSE
                outMallGoodsNo=iLine.FItemArray(23)
                shoplinkermallname=""
                shoplinkerPrdCode=""
                shoplinkerOrderID=""
                shoplinkerMallID =""
            ENd IF

            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@SellSite" , adVarchar	, adParamInput, 32, SellSite)	_
    			,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,32, iLine.FItemArray(0))	_
    			,Array("@SellDate"	,adDate, adParamInput,, iLine.FItemArray(1)) _
    			,Array("@PayType"	,adVarchar, adParamInput,32, iLine.FItemArray(3)) _
    			,Array("@Paydate"	,adDate, adParamInput,, iLine.FItemArray(2)) _
    			,Array("@matchItemID"	,adInteger, adParamInput,, iLine.FItemArray(15)) _
    			,Array("@matchItemOption"	,adVarchar, adParamInput,4, iLine.FItemArray(16)) _
    			,Array("@partnerItemID"	,adVarchar, adParamInput,32, iLine.FItemArray(15)) _
    			,Array("@partnerItemName"	,adVarchar, adParamInput,128, iLine.FItemArray(21)) _
    			,Array("@partnerOption"	,adVarchar, adParamInput,128, iLine.FItemArray(16)) _
    			,Array("@partnerOptionName"	,adVarchar, adParamInput,1024, iLine.FItemArray(22)) _
    			,Array("@OrderUserID"	,adVarchar, adParamInput,32, iLine.FItemArray(4)) _
    			,Array("@OrderName"	,adVarchar, adParamInput,32, iLine.FItemArray(5)) _
    			,Array("@OrderEmail"	,adVarchar, adParamInput,100, iLine.FItemArray(8)) _
    			,Array("@OrderTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(6)) _
    			,Array("@OrderHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(7)) _
    			,Array("@ReceiveName"	,adVarchar, adParamInput,32, iLine.FItemArray(9)) _
    			,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(10)) _
    			,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(11)) _
    			,Array("@ReceiveZipCode"	,adVarchar, adParamInput,20, iLine.FItemArray(12)) _
    			,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, iLine.FItemArray(13)) _
    			,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, iLine.FItemArray(14)) _
    			,Array("@SellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(18)) _
    			,Array("@RealSellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(19)) _
    			,Array("@ItemOrderCount"	,adInteger, adParamInput,, iLine.FItemArray(17)) _
    			,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, iLine.FItemArray(25)) _
    			,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
    			,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
    			,Array("@deliverymemo"	,adVarchar, adParamInput,400, iLine.FItemArray(26)) _
    			,Array("@requireDetail"	,adVarchar, adParamInput,1024, iLine.FItemArray(27)) _
    			,Array("@orderDlvPay"	,adCurrency, adParamInput,, iLine.FItemArray(28)) _
    			,Array("@orderCsGbn"	,adInteger, adParamInput,, orderCsGbn) _
    			,Array("@countryCode"	,adVarchar, adParamInput,2, countryCode) _
                ,Array("@outMallGoodsNo"	,adVarchar, adParamInput,20, outMallGoodsNo) _
    			,Array("@shoplinkerMallName" ,adVarchar, adParamInput,64, shoplinkermallname) _
    			,Array("@shoplinkerPrdCode"	,adVarchar, adParamInput,16, shoplinkerPrdCode) _
    			,Array("@shoplinkerOrderID"	,adVarchar, adParamInput,16, shoplinkerOrderID) _
    			,Array("@shoplinkerMallID"	,adVarchar, adParamInput,32, shoplinkerMallID) _
    			,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
				,Array("@overseasPrice"	,adCurrency, adParamInput,, overseasPrice) _
				,Array("@overseasDeliveryPrice"	,adCurrency, adParamInput,, overseasDeliveryPrice) _
				,Array("@overseasRealPrice"	,adCurrency, adParamInput,, overseasRealPrice) _
				,Array("@reserve01"	,adVarchar, adParamInput,32, reserve01) _
				,Array("@beasongNum11st"	,adVarchar, adParamInput,16, beasongNum11st) _
				,Array("@outmalloptionno"	,adVarchar, adParamInput,32, outmalloptionno) _
    		)

			If sellsite <> "nvstorefarmclass" Then
				If ( Trim(iLine.FItemArray(13)) = "") AND ( Trim(iLine.FItemArray(14)) = "") Then
					RetErr = -1
					retErrStr = "�ּ� ���� " & iLine.FItemArray(0)
					rw retErrStr
					set iLine = Nothing
					dbget.rollbackTrans
					response.write "<script>alert('�ּ� ���� ���� �ֽ��ϴ�. �ٽ� Ȯ���ϼ���');</script>"
					dbget.close() : response.end
				End If
			End If

			If sellsite <> "cookatmall" and sellsite <> "cnglob10x10" and sellsite <> "cnhigo" and sellsite <> "11stmy" and sellsite <> "cnugoshop" and sellsite <> "kakaogift" and sellsite <> "etsy" and sellsite <> "nvstorefarmclass" Then
				If  (iLine.FItemArray(28) <> "") AND (isnumeric(iLine.FItemArray(28))) Then		'��ۺ� 5000�� ������ ƨ���..
					'// CInt => CLng, skyer9, 2018-01-22
					If CLng(iLine.FItemArray(28)) > 5000 Then
		                RetErr = -1
		                retErrStr = "��ۺ� 5000�� �ʰ� " & iLine.FItemArray(0) & " ��ǰ�ڵ� =" & iLine.FItemArray(15)&" �ɼǸ� = "&iLine.FItemArray(22)
		                rw retErrStr
		                set iLine = Nothing
		                IF (sellsite<>"interpark") then
		                dbget.rollbackTrans
		                end if
		                response.write "<script>alert('��ۺ� 5000���� �ѽ��ϴ�. �ٽ� Ȯ���ϼ���');</script>"
		                dbget.close() : response.end
		            End If
				End If
			End If

        	'//�����ȣ ���̿� "-" �� �������
        	if (SellSite<>"cn10x10") and (SellSite<>"cnglob10x10") and (SellSite<>"cnhigo") and (SellSite <> "11stmy") and (SellSite <> "cnugoshop") and (SellSite <> "zilingo") and (SellSite <> "nvstorefarmclass") then
				'�����ȣ ġȯ..2015-12-23 16:08 ������ �����ȣ�� 5�ڸ� �̸��� �� ƨ���..
        		If Len(iLine.FItemArray(12)) <= 4 Then	'wizwid�� �����ȣ�� 4�ڸ��� �Ѿ��..�����
	                RetErr = -1
	                retErrStr = "�����ȣ 5�ڸ� �̸�"
	                rw retErrStr
	                set iLine = Nothing
	                IF (sellsite<>"interpark") then
	                dbget.rollbackTrans
	                end if
	                response.write "<script>alert('�����ȣ�� 5�ڸ� �̸��Դϴ�. �ٽ� Ȯ���ϼ���');</script>"
	                dbget.close() : response.end
				Else
	        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
			        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
			                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
			        	end if
			        End If
        		End If

'        		If Len(iLine.FItemArray(12)) = 4 Then			'2015-12-23 15:18 ������ ����..wizwid�� �����ȣ�� 4�ڸ��� �Ѿ��..������ 0����
'        			iLine.FItemArray(12) = CStr("0"&iLine.FItemArray(12))
'        		End If

'        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 ������ ����..�����ȣ�� 5�ڸ��� �ƴҶ� �Ʒ� IF�� ����
'		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
'		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
'		        	end if
'		        End If
			end if

			If sellsite = "cnglob10x10" or sellsite = "cnhigo" or sellsite = "cnugoshop" Then
				If (len(iLine.FItemArray(10)) > 16) OR (len(iLine.FItemArray(11)) > 16) OR (len(iLine.FItemArray(6)) > 16) OR (len(iLine.FItemArray(7)) > 16) Then
	                RetErr = -1
	                retErrStr = "�ֹ��� OR ������ ��ȭ��ȣ ���� 16�ڸ� �ʰ�"
	                rw retErrStr
	                set iLine = Nothing
	                IF (sellsite<>"interpark") then
	                dbget.rollbackTrans
	                end if
	                response.write "<script>alert('�ֹ��� OR ������ ��ȭ��ȣ ���� 16�ڸ� �ʰ�');</script>"
	                dbget.close() : response.end
				End If
			End If

            If (SellSite="ezwel") Then
                If  (iLine.FItemArray(29) <> "����غ���") then
                    RetErr = -1
	                retErrStr = "ezwell ���� üũ �ֹ����� ����غ��� �� ����-"&iLine.FItemArray(0)&":"&iLine.FItemArray(29) ''ä���� ��û 2015/03/02 �߰�
	                rw retErrStr

	                IF (sellsite<>"interpark") then
		                dbget.rollbackTrans
	                end if
	                response.write "<script>alert('"&"ezwell ���� üũ �ֹ����� ����غ��� �� ����-"&iLine.FItemArray(0)&":"&iLine.FItemArray(29)&"');</script>"
	                dbget.close() : response.end
                end if
            end if

            if (iLine.FItemArray(16)<>"") and (iLine.FItemArray(15)<>"-1") and (iLine.FItemArray(15)<>"") then
                sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
                retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

                RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
                retErrStr  = GetValue(retParamInfo, "@retErrStr") ' ������

                if (RetErr<0) and (RetErr<>-1) then ''Break
                    rw retErrStr
                    set iLine = Nothing
                    IF (sellsite<>"interpark") then
                    dbget.rollbackTrans
                    end if
                    response.write "<script>alert('"&Replace("ERROR["&retErr&"]"& retErrStr,"'","")&"');</script>"
                    dbget.close() : response.end
                end if
            else
                RetErr = -1
                retErrStr = "��ǰ�ڵ� �Ǵ� �ɼ��ڵ�  ��Ī ����" & iLine.FItemArray(0) & " ��ǰ�ڵ� =" & iLine.FItemArray(15)&" �ɼǸ� = "&iLine.FItemArray(22)
                rw retErrStr
                set iLine = Nothing
                IF (sellsite<>"interpark") then
                dbget.rollbackTrans
                end if
                response.write "<script>alert('"&Replace("ERROR["&retErr&"]"& retErrStr,"'","")&"');</script>"
                dbget.close() : response.end
            end if

            IF RetErr=0 then
                okCNT = okCNT +1
            ELSE
                'rw "RetErr:"&RetErr&":"&iLine.FItemArray(0)&":"&shoplinkerMallID
                'rw "retErrStr:"&retErrStr
                errCNT = errCNT + 1
                totErrMsg = totErrMsg + retErrStr + VbCRLF
            end if

        END IF

            IF (retErr)<>0 then

            END IF

            set iLine = Nothing
        end if

    Next
IF (sellsite<>"interpark") then
    dbget.CommitTrans
end if

''ǰ��/���� ����üũ ---------------------------------------------
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
''-------------------------------------------------------------

IF errCNT<>0 then
    response.write "<script>alert('"&errCNT&"�� �Է¿���.\n\n"&Replace(totErrMsg,vbCRLF,"\n")&"')</script>"
end if
response.write "<script>alert('"&okCNT&"�� �ԷµǾ����ϴ�.')</script>"
response.end
response.write "<script>opener.location.reload();self.close();</script>"
'''====================================================================================

Class TXLRowObj
    public FItemArray

    public function setArrayLength(ln)
        Redim FItemArray(ln)
    end function
End Class

function convPayTypeStr2Code(oStr)
    SELECT CASE oStr
        CASE "�ſ�ī��" : convPayTypeStr2Code="100"
        CASE "�ſ�" : convPayTypeStr2Code="100"
        CASE "������" : convPayTypeStr2Code="7"
        CASE "�ǽð���ü" : convPayTypeStr2Code="20"
        CASE "�ڵ�������" : convPayTypeStr2Code="400"
        CASE "�޴�������" : convPayTypeStr2Code="400"
        CASE "�ڵ���" : convPayTypeStr2Code="400"
        CASE "�޴���" : convPayTypeStr2Code="400"
        CASE ELSE : convPayTypeStr2Code="50"

    END SELECT
end function

function IsSKipRow(ixlRow, skipCol0Str)
    if Not IsArray(ixlRow) then
        IsSKipRow = true
        Exit function
    end if

    if  LCASE(ixlRow(0))=LCASE(skipCol0Str) then
        IsSKipRow = true
        Exit function
    end if

    IsSKipRow = false
end function

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen)
    Dim conDB, Rs, strQry, iResult, i, J, iObj
    Dim irowObj, strTable
    '' on Error ���� ���� �ȵ�.. ���� ���ѷ��� ���µ�.

    Set conDB = Server.CreateObject("ADODB.Connection")
	conDB.Provider = "Microsoft.Jet.oledb.4.0"		'2017-10-30 ������ �ϴ����� ����
	'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"

	If SellSite = "gabangpop" or (SellSite="itsGabangpop") or SellSite = "musinsaITS" or (SellSite="itsMusinsa") Then
    	conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"		'ù����� ������(HDR), �ʵ�Ӽ�����(IMEX;����/�ؽ�Ʈ)
    Else
		conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 �߰� 2013/12/19
	End If

 ''   On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
			'/������ �˼� ���� ������ ������. "����ġ ���� ����. �ܺ� ��ü�� Ʈ�� ������ ����(C0000005)�� �߻��߽��ϴ�. ��ũ��Ʈ�� ��� ������ �� �����ϴ�"
			set conDB = nothing
            exit function
        End if
 ''  On Error Goto 0

    '' get First Sheet Name=============''��Ʈ�� �������ΰ�� ������ �� ����.
    Set Rs = conDB.OpenSchema(adSchemaTables)

    IF Not Rs.Eof Then
        aSheetName = Rs.Fields("table_name").Value
        ''rw "aSheetName="&aSheetName
    ENd IF
    Set Rs = Nothing
    ''==================================

    Set Rs = Server.CreateObject("ADODB.Recordset")

    ''strQry = "Select * From [sheet1$]"
    strQry = "Select * From ["&aSheetName&"]"

    ReDim xlRowALL(0)
    fnGetXLFileArray = true

''On Error Resume Next
    Rs.Open strQry, conDB
        IF (ERR) then
            fnGetXLFileArray=false
            Rs.Close
            Set Rs = Nothing
            Set conDB = Nothing
            exit function
        End if

        If Not Rs.Eof Then
            Do Until Rs.Eof
                IF (ERR) then
                    fnGetXLFileArray=false
                    Rs.Close
                    Set Rs = Nothing
                    Set conDB = Nothing
                    exit function
                End if

                set irowObj = new TXLRowObj
                irowObj.setArrayLength(iArrayLen)

                For i=0 to ArrayLen
					if Not IsArray(xlPosArr(i)) then
						'// ���� ����
						if (xlPosArr(i)<0) then
							irowObj.FItemArray(i) = ""
						else
							'2019-10-11 15:05 ������ gmartket1010 ���� �߰�
							If ((SellSite="gmarket1010") OR (SellSite="auction1010") OR (SellSite="hmall1010") OR (SellSite="gseshop") ) AND (i = 22) Then
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","��")
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
							End If
						end if
					else
						'// ���� �ʵ带 ���ľ� �� ��(�� : ezwel)
						tmpVal = 0
						for each tmpItem in xlPosArr(i)
							If (SellSite="ezwel") Then
								tmpVal = tmpVal + CLng(Trim(Replace(Replace(null2blank(Rs(tmpItem)),"*",""), "(��ǰ����)", "")))
							end if
						next
						irowObj.FItemArray(i) = tmpVal
					end if

                    ''rw irowObj.FItemArray(i)
                Next

                IF (Not IsSKipRow(irowObj.FItemArray,skipString)) then
                    ReDim Preserve xlRowALL(UBound(xlRowALL)+1)

                    set xlRowALL(UBound(xlRowALL)) =  irowObj
                    ''xlRowALL(UBound(xlRowALL)).arrayObj = xlRow

                END IF
                set irowObj = Nothing
                Rs.MoveNext
            Loop
       else
          fnGetXLFileArray=false
       end if

       ''''On Error Goto 0
        IF (ERR) then
            fnGetXLFileArray=false
        End if
    Rs.Close
''On Error Goto 0

    Set Rs = Nothing
    Set conDB = Nothing

    if Ubound(xlRowALL)< 1 then fnGetXLFileArray=false

End Function

Function AddTmpDbOrderData(ixlRowALL)
    AddTmpDbOrderData = false
end Function
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
