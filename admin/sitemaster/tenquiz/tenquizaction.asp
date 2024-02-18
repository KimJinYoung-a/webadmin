<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
'퀴즈
dim topTitle
dim backGroundImage
dim MWbackGroundImage
dim PCWbackGroundImage
dim quizDescription
dim questionHintNumber
dim totalMileage
dim quizStartDate
dim quizStartTime
dim quizEndDate
dim quizEndTime
dim chasu
dim totalQuestionCount
dim startDescription
dim quizStatus
dim productEvtNum
dim endAlertTxt
dim waitingAlertTxt
dim mode, sqlStr
dim deleteUserId

dim monthGroup
dim adminRegister
dim adminName
dim adminModifyer
dim adminModifyerName

dim idx

'문항
dim Qchasu
dim Qtype
dim QquestionNumber
dim Qquestion
dim QquestionType1Image1
dim QquestionType1Image2
dim QquestionType1Image3
dim QquestionType1Image4
dim QquestionExample1
dim QquestionExample2
dim QquestionExample3
dim QquestionExample4
dim Qtype2TextExample1
dim Qtype2TextExample2
dim Qtype2TextExample3
dim Qtype2TextExample4
dim Qanswer
dim QnumOfType1Image

'퀴즈 순서 변경 관련 파라미터
dim OrderChangedFlag
dim newSeq
dim originalSeq
dim i

'퀴즈 파라미터
topTitle			= Request("topTitle")		
backGroundImage		= Request("backgroundImage")	
MWbackGroundImage	= Request("MWbackGroundImage")	
PCWbackGroundImage	= Request("PCWbackGroundImage")	
quizDescription		= Request("quizDescription")
productEvtNum		= Request("productEvtNum")
questionHintNumber	= Request("questionHintNumber")	
totalMileage		= Request("totalMileage")	
quizStartDate		= Request("quizStartDate")& " " &Request("quizStartTime")
quizEndDate			= Request("quizEndDate")& " " &Request("quizEndTime")
chasu				= Request("chasu")
totalQuestionCount	= Request("totalQuestionCount")
startDescription	= Request("startDescription")
mode				= Request("mode")
quizStatus			= Request("quizStatus")
idx					= Request("idx")
waitingAlertTxt	 	= Request("waitingAlertTxt")
endAlertTxt			= Request("endAlertTxt")
deleteUserId 		= Request("deleteUserId")

'문항 파라미터
Qtype				 = Request("type")
QquestionNumber		 = Request("questionNumber")
Qquestion			 = Request("question")
QquestionType1Image1 = Request("questionType1Image1")
QquestionType1Image2 = Request("questionType1Image2")
QquestionType1Image3 = Request("questionType1Image3")
QquestionType1Image4 = Request("questionType1Image4")

If Qtype = 1 Then
	QquestionExample1	 = Request("questionExample1")
	QquestionExample2	 = Request("questionExample2")
	QquestionExample3	 = Request("questionExample3")
	QquestionExample4	 = Request("questionExample4")	
Else
	QquestionExample1	 = Request("questionExample1img")
	QquestionExample2	 = Request("questionExample2img")
	QquestionExample3	 = Request("questionExample3img")
	QquestionExample4	 = Request("questionExample4img")	
	Qtype2TextExample1	 = Request("type2TextExample1")
	Qtype2TextExample2	 = Request("type2TextExample2")
	Qtype2TextExample3	 = Request("type2TextExample3")
	Qtype2TextExample4	 = Request("type2TextExample4")		
End if	

Qanswer				 = Request("answer")
QnumOfType1Image	 = Request("numOfType1Image")

'기타
OrderChangedFlag   = Request("OrderChangedFlag")


if OrderChangedFlag = "1" then
	for i=1 to request("originalSeq").count
		originalSeq = request("originalSeq")(i)
		newSeq = request("newSeq")(i)		

sqlStr = "Update [db_sitemaster].[dbo].tbl_PlayingTenquizquestiondata " &_
		" Set questionNumber ='" & newSeq & "'" &_								
		" Where idx =" & originalSeq		

		dbget.Execute(sqlStr)		
	next			
end if

public Function GetAdminName(adminid)	
	If IsNull(adminid) Or adminid="" Then Exit Function
	On Error Resume Next
	dim SqlStr

	sqlStr = " Select top 1 username "
	sqlStr = sqlStr & " From db_partner.dbo.tbl_user_tenbyten "
	sqlStr = sqlStr & " where userid = '"& adminid &"'"
	rsget.CursorLocation = adUseClient
	rsget.CursorType=adOpenStatic
	rsget.Locktype=adLockReadOnly
	rsget.Open sqlStr, dbget

	If Not(rsget.bof Or rsget.eof) Then
		GetAdminName = rsget("username")
	End If
	rsget.close
	On Error goto 0
End Function	


'// 모드에 따른 분기
Select Case mode
	Case "deleteUserData"		
		if chasu = "" or deleteUserId = "" then
		%>
		<script language="javascript">
		<!--
			// 페이지 새로고침
			alert("차수나 사용자 아이디가 없습니다.");
			window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
			self.close();        // 팝업창 닫기
		//-->
		</script>		
		<%
		else
			sqlStr = "delete db_sitemaster.DBO.tbl_playingtenquizusermasterdata " &_
					" Where userid='" & deleteUserId & "'" &_		
					" and chasu = '" & chasu & "'"
			dbget.Execute(sqlStr)				
			sqlStr = "delete db_sitemaster.DBO.tbl_playingtenquizuserdetaildata " &_
					" Where userid='" & deleteUserId & "'" &_		
					" and chasu = '" & chasu & "'"
			dbget.Execute(sqlStr)								
		end if	
	'문항 추가
	Case "subadd"		
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_PlayingTenquizquestiondata " &_
					" (chasu , type , questionNumber, question , questionType1Image1, questionType1Image2, questionType1Image3, questionType1Image4, numOfType1Image " &_
					" , questionExample1, questionExample2, questionExample3, questionExample4, type2TextExample1, type2TextExample2, type2TextExample3, type2TextExample4, answer,registDate,lastupdate,isusing ) values "&_
					" ('" & chasu &"'" &_
					" ,'" & Qtype &"'" &_
					" ,'" & QquestionNumber &"'" &_
					" ,'" & Qquestion &"'" &_
					" ,'" & QquestionType1Image1 &"'" &_
					" ,'" & QquestionType1Image2 &"'" &_
					" ,'" & QquestionType1Image3 & "'" &_
					" ,'" & QquestionType1Image4 & "'" &_
					" , " & QnumOfType1Image &_							
					" ,'" & QquestionExample1 & "'" &_
					" ,'" & QquestionExample2 & "'" &_
					" ,'" & QquestionExample3 & "'" &_
					" ,'" & QquestionExample4 & "'" &_
					" ,'" & Qtype2TextExample1 & "'" &_
					" ,'" & Qtype2TextExample2 & "'" &_
					" ,'" & Qtype2TextExample3 & "'" &_
					" ,'" & Qtype2TextExample4 & "'" &_
					" ,'" & Qanswer & "'" &_
					" , getdate()" &_
					" , getdate()" &_
					" , 'Y'" &_									
					")"
		'response.write sqlStr					
		'response.end
		dbget.Execute(sqlStr)					
	'문항 수정	
	Case "submodify"
		sqlStr = "Update [db_sitemaster].[dbo].tbl_PlayingTenquizquestiondata " &_
				" Set type ='" & Qtype & "'" &_
				" 	,questionNumber ='" & QquestionNumber & "'" &_
				" 	,question ='" & Qquestion & "'" &_
				" 	,questionType1Image1 ='" & QquestionType1Image1 & "'" &_
				" 	,questionType1Image2 ='" & QquestionType1Image2 & "'" &_
				" 	,questionType1Image3 ='" & QquestionType1Image3 & "'" &_
				" 	,questionType1Image4 ='" & QquestionType1Image4 & "'" &_
				" 	,numOfType1Image =" & QnumOfType1Image &_					
				" 	,questionExample1 ='" & QquestionExample1 & "'" &_
				" 	,questionExample2 ='" & QquestionExample2 & "'" &_
				" 	,questionExample3 ='" & QquestionExample3 & "'" &_
				" 	,questionExample4 ='" & QquestionExample4 & "'" &_
				" 	,type2TextExample1 ='" & Qtype2TextExample1 & "'" &_
				" 	,type2TextExample2 ='" & Qtype2TextExample2 & "'" &_
				" 	,type2TextExample3 ='" & Qtype2TextExample3 & "'" &_
				" 	,type2TextExample4 ='" & Qtype2TextExample4 & "'" &_				
				" 	,answer ='" & Qanswer & "'" &_
				" 	,lastUpDate = getdate()" &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)
	Case "subdelete"
		sqlStr = "Update [db_sitemaster].[dbo].tbl_PlayingTenquizquestiondata " &_
				" Set isusing ='N' " &_
				" Where idx=" & idx		
		dbget.Execute(sqlStr)		
	Case "add"
		adminName = GetAdminName(session("ssBctId"))			

		'신규 등록
		monthGroup = Mid(chasu,1,6)
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_PlayingTenQuizData " &_
					" (chasu, monthGroup , topTitle , quizDescription, backGroundImage , questionHintNumber, totalMileage, quizStartDate, quizEndDate, endAlertTxt, waitingAlertTxt " &_
					" , totalQuestionCount, startDescription, adminRegister, adminName ,adminModifyer,adminModifyerName, registDate, modifydate, quizStatus, MWbackGroundImage, PCWbackGroundImage, productEvtNum, mileageDiv) values "&_
					" ('" & chasu &"'" &_
					" ,'" & monthGroup &"'" &_
					" ,'" & topTitle &"'" &_
					" ,'" & quizDescription &"'" &_
					" ,'" & backGroundImage & "'" &_
					" ,'" & questionHintNumber & "'" &_
					" ,'" & totalMileage * 10000 & "'" &_
					" ,'" & quizStartDate & "'" &_
					" ,'" & quizEndDate & "'" &_
					" ,'" & endAlertTxt & "'" &_
					" ,'" & waitingAlertTxt & "'" &_
					" ,'" & totalQuestionCount & "'" &_
					" ,'" & startDescription & "'" &_
					" ,'" & session("ssBctId") & "'" &_
					" ,'" & adminName & "'" &_
					" ,'" & session("ssBctId") & "'" &_
					" ,'" & adminName & "'" &_										
					" ,	getdate()" &_															
					" ,	getdate()" &_				
					" ,'" & quizStatus & "'" &_																					
					" ,'" & MWbackGroundImage & "'" &_	
					" ,'" & PCWbackGroundImage & "'" &_	
					" ,'" & productEvtNum & "'" &_						
					" ,	0" &_					
					")"		
		dbget.Execute(sqlStr)
	Case "modify"
		'내용 수정	
		monthGroup = Mid(chasu,1,6)
		adminModifyerName = GetAdminName(session("ssBctId"))			

		sqlStr = "Update [db_sitemaster].[dbo].tbl_PlayingTenQuizData " &_
				" Set chasu='" & chasu & "'" &_
				" 	,monthGroup='" & monthGroup & "'" &_
				" 	,topTitle='" & topTitle & "'" &_
				" 	,quizDescription='" & quizDescription & "'" &_
				" 	,backGroundImage='" & backGroundImage & "'" &_
				" 	,MWbackGroundImage='" & MWbackGroundImage & "'" &_				
				" 	,PCWbackGroundImage='" & PCWbackGroundImage & "'" &_				
				" 	,productEvtNum='" & productEvtNum & "'" &_				
				" 	,questionHintNumber='" & questionHintNumber & "'" &_
				" 	,totalMileage='" & totalMileage * 10000 & "'" &_
				" 	,quizStartDate='" & quizStartDate & "'" &_
				" 	,quizEndDate='" & quizEndDate & "'" &_
				" 	,totalQuestionCount='" & totalQuestionCount & "'" &_
				" 	,startDescription='" & startDescription & "'" &_
				" 	,adminModifyer='" & adminModifyer & "'" &_
				" 	,adminModifyerName='" & adminModifyerName & "'" &_
				" 	,modifydate=getdate()" &_
				" 	,quizStatus='" & quizStatus & "'" &_		
				" 	,endAlertTxt='" & endAlertTxt & "'" &_		
				" 	,waitingAlertTxt='" & waitingAlertTxt & "'" &_									
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)	
	Case "mileagediv"
		Dim objCmd, result, alertTxt
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_sitemaster].[dbo].sp_Tenquiz_mileage_division('"&chasu&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    result = objCmd(0).Value
		Set objCmd = Nothing				
		Select Case result
			case 0	'실패
				alertTxt = "시스템 오류입니다."
			case 1	'성공	
				alertTxt = "마일리지 분배가 완료되었습니다."
			case 2	'유효한 차수가 아님
				alertTxt = "유효한 차수가 아닙니다."
			case 3	'이미 분배한 차수
				alertTxt = "이미 분배한 차수입니다."
		end select 	
End Select
%>
<% If mode = "subadd"  Or mode = "submodify" then%>
<script>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
	 self.close();        // 팝업창 닫기
//-->
</script>
<% elseif mode = "subdelete" then %>
<script language="javascript">
<!--
	// 페이지 새로고침
	alert("삭제했습니다.");
	location.href = document.referrer;
//-->
</script>
<% elseif mode = "mileagediv" then %>
<script language="javascript">
<!--
	// 페이지 새로고침
	alert("<%=alertTxt%>");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
	 self.close();        // 팝업창 닫기
//-->
</script>
<% elseif mode = "deleteUserData" then %>
<script language="javascript">
<!--
	// 페이지 새로고침
	alert("삭제했습니다.");	
	window.opener.document.location.reload();    // 부모창 새로고침
	self.close();        // 팝업창 닫기
//-->
</script>
<% Else %>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "index.asp";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
