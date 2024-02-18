<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
dim msg, intloop, menupos
dim mode
dim code_mid, code_large , moveidx , arrmoveidx
dim SQL , retURL
Dim yyyy1, mm1


'// 내용 접수 및 처리
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)

code_mid	= RequestCheckvar(Request("code_mid"),3)
code_large	= RequestCheckvar(Request("code_large"),3)

moveidx		=	RequestCheckvar(request("moveidx"),10)
yyyy1		=	RequestCheckvar(request("yyyy1"),4)
mm1			=	RequestCheckvar(request("mm1"),2)

If InStr(moveidx,",") > 0 then
	arrmoveidx = split(moveidx,",")
Else  
	arrmoveidx = moveidx
End If 

'==============================================================================
'## 내용 저장(수정) 처리

if code_large="" then
	response.write	"<script language='javascript'>" &_
					"	alert('카테고리 구분이 없습니다.');" &_
					"	history.back();" &_
					"</script>"
	dbget.close()	:	response.End
end If

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode

	Case "modify"
		'@@ 수정처리

		If InStr(moveidx,",") > 0 then

			For intLoop = 0 To UBound(arrmoveidx)	
					SQL =	"Update db_academy.dbo.tbl_lec_item Set " &_
							"	newCate_Large = '" & code_large & "'" &_
							"	,newCate_mid = '" & code_mid & "'" &_
							" Where idx = '" & Trim(arrmoveidx(intloop)) & "'"
			
				'저장 처리
				dbACADEMYget.Execute(SQL)
			Next 
		
		Else
			
			SQL =	"Update db_academy.dbo.tbl_lec_item Set " &_
							"	newCate_Large = '" & code_large & "'" &_
							"	,newCate_mid = '" & code_mid & "'" &_
							" Where idx = '" & Trim(arrmoveidx) & "'"
			
			'저장 처리
			dbACADEMYget.Execute(SQL)
		
		End If 

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "pop_chg_lec.asp?menupos=" & menupos & "&yyyy1=" & yyyy1 & "&mm1=" & mm1

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->