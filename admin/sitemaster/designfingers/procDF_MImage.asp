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
Dim sMode, menupos, strSql
Dim iDFSeq, ievt_code,sevt_link
Dim sDFType,sTitle,tContents,sPrizeDate,blnDisplay,blnOtherMall, sComment, blnMainDisplay
Dim arrItemid, sProdName, sProdSize, sProdColor, sProdJe, sProdGu, sProdSpe

Dim iImgCode
Dim arrMainTop, arrTop, arrSmall, arrList, arr3dView, arrAdd, arrTxtAdd(30), intLoop, arrEventLeft, arrEventRight
Dim arrInfo(10), tmp3dv
Dim sImgURL, strMsg, vCount

sMode	=  requestCheckVar(request("sM"),1)
iDFSeq			=  requestCheckVar(request("iDFS"),10)
vCount	= request("tempcount")


If iDFSeq = "" Then
	Response.Write "<script>alert('�߸��� ����Դϴ�.');window.close()</script>"
	dbget.close()
	Response.End
End If


	arrAdd 		= split(request("imgmobile"),",")	'//add Ÿ�� 

	If uBound(arrAdd) = "-1" Then
		arrAdd = split(",",",")
	End If

	'//add �̹��� ��	
	for intLoop = 0 to 14
		arrTxtAdd(intLoop)	= html2db(request("tA"&(intLoop+1)))	'//add Ÿ�� ��ũ	
	next

	dbget.beginTrans

	'//�����̹��� ����
		strSql = "Delete FRom [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = "&iDFSeq&" and DFCodeSeq = '24' "
		dbget.execute strSql	
		
	' Mobile add- codeNo:24
	Dim vMobileChk
	vMobileChk = "x"
	for intLoop = 0 To vCount-1	'uBound(arrAdd)
	if(trim(arrAdd(intLoop)) <> "" OR trim(arrTxtAdd(intLoop)) <> "" ) THEN	
		iImgCode = 24
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[link])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrAdd(intLoop))&"','"&arrTxtAdd(intLoop)&"')"									
		dbget.execute strSql
		
		vMobileChk = "o"
	end if
	
'"&arrTxtAdd(intLoop)&"
	next

	If vMobileChk = "o" Then
		strSql = "UPDATE [db_sitemaster].[dbo].[tbl_designfingers] SET IsMobile = 'Y' WHERE DFSeq = '" & iDFSeq & "' "
		dbget.execute strSql
	ElseIf vMobileChk = "x" Then
		strSql = "UPDATE [db_sitemaster].[dbo].[tbl_designfingers] SET IsMobile = 'N' WHERE DFSeq = '" & iDFSeq & "' "
		dbget.execute strSql
	End If

	IF Err.Number = 0 THEN
		dbget.CommitTrans
		
		Call sbAlertMsg ("�����Ǿ����ϴ�.", "/admin/sitemaster/designfingers/regDF_MImage.asp?iDFS="&iDFSeq, "self") 
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans	  	
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "") 
	END IF	
	dbget.close()	:	response.End

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

	
