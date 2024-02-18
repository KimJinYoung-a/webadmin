<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim SqlStr
dim idx
dim giftcardImage
dim giftcardAlt
dim adminRegister
dim adminName
dim adminModifyer
dim adminModifyerName
dim registDate
dim lastUpDate
dim sortNumber
dim isusing

dim mode

idx					= request("idx")
mode 				= request("mode")
giftcardImage		= request("giftcardImage")	
giftcardAlt			= request("giftcardAlt")
adminRegister		= request("adminRegister")
adminName			= request("adminName")
adminModifyer		= request("adminModifyer")	
adminModifyerName	= request("adminModifyerName")			
registDate			= request("registDate")
lastUpDate			= request("lastUpDate")	
isusing				= request("isusing")

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
	Case "add"
		adminName = GetAdminName(session("ssBctId"))			

		'신규 등록		
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_giftcard_image " &_
					" (designId, giftcardImage, giftcardAlt, adminRegister, adminName, adminModifyer, adminModifyerName, registDate, lastUpDate, isusing, sortNumber ) values " &_					
					" ( " &_
					" ( SELECT MAX(DESIGNID) + 1 FROM [db_sitemaster].DBO.tbl_giftcard_image where designid <> 900)" &_
					" ,'" & giftcardImage & "'" &_
					" ,'" & giftcardAlt & "'" &_
					" ,'" & session("ssBctId") & "'" &_
					" ,'" & adminName & "'" &_
					" ,'" & session("ssBctId") & "'" &_
					" ,'" & adminName & "'" &_										
					" ,	getdate()" &_															
					" ,	getdate()" &_				
					" ,'" & isusing & "'" &_																					
					" ,(select max(sortNumber) + 1 from [db_sitemaster].DBO.tbl_giftcard_image where designid <> 900 ) " &_
					")"		
		dbget.Execute(sqlStr)
	Case "modify"
		'내용 수정			
		adminModifyerName = GetAdminName(session("ssBctId"))			

		sqlStr = "Update [db_sitemaster].[dbo].tbl_giftcard_image " &_
				" Set giftcardImage='" & giftcardImage & "'" &_				
				" 	,giftcardAlt='" & giftcardAlt & "'" &_
				" 	,adminModifyer='" & session("ssBctId") & "'" &_
				" 	,adminModifyerName='" & adminModifyerName & "'" &_
				" 	,lastUpDate= getdate() "&_
				" 	,isusing='" & isusing & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)	
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
