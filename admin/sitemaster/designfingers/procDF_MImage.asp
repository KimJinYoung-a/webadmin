<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  디자인핑거스 DB 처리
' History : 2008.03.14 생성
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
	Response.Write "<script>alert('잘못된 경로입니다.');window.close()</script>"
	dbget.close()
	Response.End
End If


	arrAdd 		= split(request("imgmobile"),",")	'//add 타입 

	If uBound(arrAdd) = "-1" Then
		arrAdd = split(",",",")
	End If

	'//add 이미지 맵	
	for intLoop = 0 to 14
		arrTxtAdd(intLoop)	= html2db(request("tA"&(intLoop+1)))	'//add 타입 링크	
	next

	dbget.beginTrans

	'//기존이미지 삭제
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
		
		Call sbAlertMsg ("수정되었습니다.", "/admin/sitemaster/designfingers/regDF_MImage.asp?iDFS="&iDFSeq, "self") 
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans	  	
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "") 
	END IF	
	dbget.close()	:	response.End

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

	
