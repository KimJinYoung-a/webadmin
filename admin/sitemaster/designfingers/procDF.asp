<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������ΰŽ� DB ó��
' History : 2008.03.14 ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
Dim sMode, menupos, strSql, iCurrpage
Dim iDFSeq, ievt_code,sevt_link
Dim sDFType,sTitle,tContents,sPrizeDate,blnDisplay,blnOtherMall, sComment, blnMainDisplay
Dim arrItemid, sProdName, sProdSize, sProdColor, sProdJe, sProdGu, sProdSpe, sIsMovie, sOpenDate, sTag, edid, emktid

Dim iImgCode
Dim arrMainTop, arrTop, arrSmall, arrList, arr3dView, arrAdd, arrTxtAdd(30), intLoop, arrEventLeft, arrEventRight, arrplay
Dim arrInfo(10), tmp3dv
Dim sImgURL, strMsg

sMode=  requestCheckVar(request("sM"),1)
menupos = requestCheckVar(request("menupos"),10)
sOpenDate = request("opendate")
sTag = html2db(requestCheckVar(request("sTag"),100))
iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ

SELECT Case sMode
Case "I"
	'// ������
	sDFType 		= requestCheckVar(request("selDFT"),10)
	sTitle 			= html2db(requestCheckVar(request("sT"),64))
	tContents 		= html2db(request("txtC"))
	sPrizeDate 		= requestCheckVar(request("dPD"),10)
	sComment		= html2db(requestCheckVar(request("sCom"),200))
	sProdName		= html2db(requestCheckVar(request("sPdN"),100))
	sProdSize		= html2db(requestCheckVar(request("sPdS"),100))
	sProdColor		= html2db(requestCheckVar(request("sPdC"),100))
	sProdJe			= html2db(requestCheckVar(request("sPdJ"),100))
	sProdGu			= html2db(requestCheckVar(request("sPdG"),100))
	sProdSpe		= html2db(requestCheckVar(request("sPdP"),100))
	blnDisplay 		= requestCheckVar(request("rdoD"),1)
	blnMainDisplay 	= requestCheckVar(request("rdoMD"),1)
	blnOtherMall	= requestCheckVar(request("rdoOM"),1)

	edid  		= requestCheckVar(Request("selDId"),32)		'��� �����̳�
	emktid 		= requestCheckVar(Request("selMKTId"),32)		'��� MKT

	'//��ǰ
	arrItemid	= request("arrI")

	'// �̹���
	arrTop 		= split(request("imgtop"),",")	'//top Ÿ��
	arrplay 	= request("imgplay")			'//play Ÿ��
	arrSmall 	= request("imgsmall")			'//small Ÿ��  - �̹��� 1�� ���
	arrList 	= request("imglist")			'//list Ÿ�� 	- �̹��� 1�� ���
	arr3dView 	= split(request("img3dv"),",")	'//3dview Ÿ��
	arrAdd 		= split(request("imgadd"),",")	'//add Ÿ��
	arrMainTop	= request("imgmain_top")
	arrEventLeft	= request("imgeventLeft")
	arrEventRight	= request("imgeventRight")

	'//3dview �̹��� ����
	tmp3dv			=  split(request("sel3dv"),",")
	for intLoop = 0 to ubound(tmp3dv)
		arrInfo(intLoop) = trim(tmp3dv(intLoop))
	next

	sIsMovie = request("ismovie")

	'//add �̹��� ��
	for intLoop = 0 to 14
		arrTxtAdd(intLoop)	= html2db(request("tA"&(intLoop+1)))	'//add Ÿ�� ��ũ
	next

	dbget.beginTrans
	'// ������ ����
	strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers]([DFType], [Title], [Contents], [PrizeDate], [Comment],[IsDisplay] ,[IsOtherMall], [Userid], [IsMainDisplay] " & _
			"	, [ProdName], [ProdSize], [ProdColor], [ProdJe], [ProdGu], [ProdSpe], [IsMovie], [OpenDate], [Tag], [designerid], [partMKTid]) "&_
			" VALUES("&sDFType&",'"&sTitle&"','"&tContents&"','"&sPrizeDate&"','"&sComment&"',"&blnDisplay&","&blnOtherMall&",'"&session("ssBctId")&"','"&blnMainDisplay&"'," & _
			"	'"&sProdName&"', '"&sProdSize&"', '"&sProdColor&"', '"&sProdJe&"', '"&sProdGu&"', '"&sProdSpe&"', '" & sIsMovie & "', '" & sOpenDate & "', '" & sTag & "', '" & edid & "', '" & emktid & "')"
		dbget.execute strSql

		strSql =" select SCOPE_IDENTITY()"
		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			iDFSeq = rsget(0)
 		ELSE
 			dbget.RollBackTrans
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
 		END IF
 		rsget.close

	'//��ǰ ���� : �̺�Ʈ ��ǰ ���̺� ���� - �̺�Ʈ �ڵ�� �׻� 1�� ����
	IF arrItemid <> "" THEN
		ievt_code = 1
		sevt_link = html2db("\designfingers\designfingers.asp?fingerid="&iDFSeq)
		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem](evt_code, itemid, evtgroup_code, evtitem_linkurl)"&_
				" SELECT "&ievt_code&", itemid,"&iDFSeq&",'"&sevt_link&"' "&_
				"	FROM [db_item].[dbo].[tbl_item] WHERE itemid in ("&arrItemid&")"
		dbget.execute strSql
	END IF

	'//�̹������� : ��ġ�� �ڵ� �������� �̹�������
	' top- codeNo:2
	for intLoop = 0 To uBound(arrTop)
	if(trim(arrTop(intLoop)) <> "" ) THEN
		iImgCode = 2
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrTop(intLoop))&"')"
		dbget.execute strSql
	end if
	next

	' small - codeNo:3
	if(trim(arrSmall) <> "" ) THEN
		iImgCode = 3
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrSmall)&"')"
		dbget.execute strSql
	end if

	' list - codeNo:4
	if(trim(arrList) <> "" ) THEN
		iImgCode = 4
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrList)&"')"
		dbget.execute strSql
	end if

	' 3dview- codeNo:7
	for intLoop = 0 To uBound(arr3dView)
	if(trim(arr3dView(intLoop)) <> "" ) THEN
		iImgCode = 7
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[ImgDescCode])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arr3dView(intLoop))&"',"&arrInfo(intLoop)&")"
		dbget.execute strSql
	end if
	next

	' play - codeNo:8
	if(trim(arrplay) <> "" ) THEN
		iImgCode = 8
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrplay)&"')"
		dbget.execute strSql
	end if

	' add- codeNo:5
	for intLoop = 0 To uBound(arrAdd)
	if(trim(arrAdd(intLoop)) <> "" OR trim(arrTxtAdd(intLoop)) <> "" ) THEN
		iImgCode = 5
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[link])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrAdd(intLoop))&"','"&arrTxtAdd(intLoop)&"')"
		dbget.execute strSql
	end if
	next

	' main_top - codeNo:21
	if(trim(arrMainTop) <> "" ) THEN
		iImgCode = 21
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrMainTop)&"')"
		dbget.execute strSql
	end if

	' eventLeft - codeNo:22
	if(trim(arrEventLeft) <> "" ) THEN
		iImgCode = 22
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventLeft)&"')"
		dbget.execute strSql
	end if

	' eventRight - codeNo:23
	if(trim(arrEventRight) <> "" ) THEN
		iImgCode = 23
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventRight)&"')"
		dbget.execute strSql
	end if

	IF Err.Number = 0 THEN
		dbget.CommitTrans

		Call sbAlertMsg ("��ϵǾ����ϴ�.", "/admin/sitemaster/designfingers/listDF.asp?menupos="&menupos&"&iC="&iCurrpage&"", "self")
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
	END IF
	dbget.close()	:	response.End
Case "U"
	'// ������
	iDFSeq			=  requestCheckVar(request("iDFS"),10)
	sDFType 		= requestCheckVar(request("selDFT"),10)
	sTitle 			= html2db(requestCheckVar(request("sT"),64))
	tContents 		= html2db(request("txtC"))
	sPrizeDate 		= requestCheckVar(request("dPD"),10)
	sComment		= html2db(requestCheckVar(request("sCom"),200))
	sProdName		= html2db(requestCheckVar(request("sPdN"),100))
	sProdSize		= html2db(requestCheckVar(request("sPdS"),100))
	sProdColor		= html2db(requestCheckVar(request("sPdC"),100))
	sProdJe			= html2db(requestCheckVar(request("sPdJ"),100))
	sProdGu			= html2db(requestCheckVar(request("sPdG"),100))
	sProdSpe		= html2db(requestCheckVar(request("sPdP"),100))
	blnDisplay 		= requestCheckVar(request("rdoD"),1)
	blnMainDisplay 	= requestCheckVar(request("rdoMD"),1)
	blnOtherMall	= requestCheckVar(request("rdoOM"),1)

	edid  		= requestCheckVar(Request("selDId"),32)		'��� �����̳�
	emktid 		= requestCheckVar(Request("selMKTId"),32)		'��� MKT

	'//��ǰ
	arrItemid	= request("arrI")

	'// �̹���
	arrTop 		= split(request("imgtop"),",")	'//top Ÿ��
	arrplay 	= request("imgplay")			'//play Ÿ��
	arrSmall 	= request("imgsmall")			'//small Ÿ��  - �̹��� 1�� ���
	arrList 	= request("imglist")			'//list Ÿ�� 	- �̹��� 1�� ���
	arr3dView 	= split(request("img3dv"),",")	'//3dview Ÿ��
	arrAdd 		= split(request("imgadd"),",")	'//add Ÿ��
	arrMainTop	= request("imgmain_top")
	arrEventLeft	= request("imgeventLeft")
	arrEventRight	= request("imgeventRight")

	'//3dview �̹��� ����
	tmp3dv			=  split(request("sel3dv"),",")

	sIsMovie = request("ismovie")

	for intLoop = 0 to UBound(tmp3dv)
		arrInfo(intLoop) = trim(tmp3dv(intLoop))
	next

	'//add �̹��� ��
	for intLoop = 0 to 14
		arrTxtAdd(intLoop)	= html2db(request("tA"&(intLoop+1)))	'//add Ÿ�� ��ũ
	next

	dbget.beginTrans
	'// ������ ����
	strSql = "UPDATE [db_sitemaster].[dbo].[tbl_designfingers] SET [DFType] = "&sDFType&", [Title]= '"&sTitle&"', [Contents]= '"&tContents&"' "&_
			", [PrizeDate]= '"&sPrizeDate&"', Comment = '"&sComment&"' ,[IsDisplay] ="&blnDisplay&",[IsOtherMall] = "&blnOtherMall&", [Userid]= '"&session("ssBctId")&"' "&_
			", [IsMainDisplay] = '"&blnMainDisplay&"', [ProdName] = '"&sProdName&"', [ProdSize] = '"&sProdSize&"', [ProdColor] = '"&sProdColor&"'" & _
			", [ProdJe] = '"&sProdJe&"', [ProdGu] = '"&sProdGu&"', [ProdSpe] = '"&sProdSpe&"', [IsMovie] = '" & sIsMovie & "', [OpenDate] = '" & sOpenDate & "', [Tag] = '" & sTag & "' " & _
			", [designerid] = '" & edid & "', [partMKTid] = '" & emktid & "' " & _
			" WHERE DFSeq = "&iDFSeq
		dbget.execute strSql

	'//��ǰ ���� : �̺�Ʈ ��ǰ ���̺� ���� - �̺�Ʈ �ڵ�� �׻� 1�� ����
	IF arrItemid <> "" THEN
		ievt_code = 1
		sevt_link = html2db("\designfingers\designfingers.asp?fingerid="&iDFSeq)

		'������� ��ǰ ����
		strSql = "DELETE From [db_event].[dbo].[tbl_eventitem] WHERE evt_code ="&ievt_code&" AND evtgroup_code ="&iDFSeq
		dbget.execute strSql

		'Ÿ �ΰŽ� ��ϻ�ǰ���� �˻�
		strSql = "Select evtgroup_code, itemid From [db_event].[dbo].[tbl_eventitem] Where itemid in ("&arrItemid&") and evt_code=" & ievt_code
		rsget.Open strSql,dbget, 1
 		IF Not(rsget.EOF or rsget.BOF) THEN
 			Do Until rsget.EOF
 				strMsg = strMsg & "- [" & rsget("itemid") & "]��ǰ�� " & rsget("evtgroup_code") & "�� �ΰŽ��� ����.\n"
 			rsget.MoveNext
 			Loop
 			dbget.RollBackTrans
 			strMsg = strMsg & "\n�̹� �ٸ� ���� ��ϵǾ��־� ���� ����� �� �����ϴ�."
 			Call sbAlertMsg (strMsg, "back", "")
 			dbget.close()	:	response.End
 		END IF
 		rsget.close

		'��ǰ ����
		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem](evt_code, itemid, evtgroup_code, evtitem_linkurl)"&_
				" SELECT "&ievt_code&", itemid,"&iDFSeq&",'"&sevt_link&"' "&_
				"	FROM [db_item].[dbo].[tbl_item] WHERE itemid in ("&arrItemid&")"
		dbget.execute strSql

	END IF

	'//�����̹��� ����
		strSql = "Delete FRom [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = "&iDFSeq&" and DFCodeSeq NOT IN (6, 24, 25) "
		dbget.execute strSql

	'//�̹������� : ��ġ�� �ڵ� �������� �̹�������
	' top- codeNo:2
	for intLoop = 0 To uBound(arrTop)
	if(trim(arrTop(intLoop)) <> "" ) THEN
		iImgCode = 2
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrTop(intLoop))&"')"
		dbget.execute strSql
	end if
	next

	' small - codeNo:3
	if(trim(arrSmall) <> "" ) THEN
		iImgCode = 3
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrSmall)&"')"
		dbget.execute strSql
	end if

	' list - codeNo:4
	if(trim(arrList) <> "" ) THEN
		iImgCode = 4
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrList)&"')"
		dbget.execute strSql
	end if

	' 3dview- codeNo:7
	for intLoop = 0 To uBound(arr3dView)
	if(trim(arr3dView(intLoop)) <> "" ) THEN
		iImgCode = 7
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[ImgDescCode])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arr3dView(intLoop))&"',"&arrInfo(intLoop)&")"
		dbget.execute strSql
	end if
	next

	' play - codeNo:8
	if(trim(arrplay) <> "" ) THEN
		iImgCode = 8
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrplay)&"')"
		dbget.execute strSql
	end if

	' add- codeNo:5
	for intLoop = 0 To uBound(arrAdd)
	if(trim(arrAdd(intLoop)) <> "" OR trim(arrTxtAdd(intLoop)) <> "" ) THEN
		iImgCode = 5
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[link])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrAdd(intLoop))&"','"&arrTxtAdd(intLoop)&"')"
		dbget.execute strSql
	end if
	next

	' main_top - codeNo:21
	if(trim(arrMainTop) <> "" ) THEN
		iImgCode = 21
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrMainTop)&"')"
		dbget.execute strSql
	end if

	' eventLeft - codeNo:22
	if(trim(arrEventLeft) <> "" ) THEN
		iImgCode = 22
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventLeft)&"')"
		dbget.execute strSql
	end if

	' eventRight - codeNo:23
	if(trim(arrEventRight) <> "" ) THEN
		iImgCode = 23
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventRight)&"')"
		dbget.execute strSql
	end if

	IF Err.Number = 0 THEN
		dbget.CommitTrans

		Call sbAlertMsg ("�����Ǿ����ϴ�.", "/admin/sitemaster/designfingers/listDF.asp?menupos="&menupos&"&iC="&iCurrpage&"", "self")
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
	END IF
	dbget.close()	:	response.End
Case "B"	'//��ʵ��
	iImgCode	= 6
	iDFSeq		=  requestCheckVar(request("iDFS"),10)
	sImgURL		=  requestCheckVar(request("sIU"),100)

	'//�����̹��� ����
		strSql = "Delete FRom [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = "&iDFSeq&" and DFCodeSeq = 6 "
		dbget.execute strSql

	'//�̹��� ���
	strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&sImgURL&"')"
	dbget.execute strSql
	IF Err.Number = 0 THEN
		Call sbAlertMsg ("��ϵǾ����ϴ�.", "listbest.asp?menupos="&menupos, "opener")
	ELSE
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
	END IF

Case "C"	'//�ڵ����
	Dim iCodeSeq, sCodeDesc, iPCodeSeq,iCodeSort, blnUsing
	iCodeSeq 	= requestCheckVar(request("iCS"),10)
	sCodeDesc 	= requestCheckVar(request("sCD"),32)
	iPCodeSeq 	= requestCheckVar(request("selPCS"),10)
	iCodeSort 	= requestCheckVar(request("iCSort"),10)
	blnUsing 	= requestCheckVar(request("rdoU"),1)
	IF iCodeSeq ="" THEN
		strSql  ="INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_code]([PCodeSeq], [CodeDesc], [CodeSort], [IsUsing])"&_
				"VALUES("&iPCodeSeq&",'"&sCodeDesc&"',"&iCodeSort&","&blnUsing&")"
		dbget.execute strSql
	ELSE
		strSql  ="UPDATE [db_sitemaster].[dbo].[tbl_designfingers_code] SET PCodeSeq = "&iPCodeSeq&", CodeDesc = '"&sCodeDesc&"', CodeSort = "&iCodeSort&", IsUsing ="&blnUsing&_
				"	WHERE DFCodeSeq = "&iCodeSeq
		dbget.execute strSql
	END IF
	IF Err.Number = 0 THEN
		Call sbAlertMsg ("����Ǿ����ϴ�.", "/admin/sitemaster/designfingers/popManageCode.asp?sPCS="&iPCodeSeq, "self")
	ELSE
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
	END IF
Case Else
	Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
END SELECT
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->


