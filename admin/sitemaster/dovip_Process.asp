<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : dovip_Process.asp
' Discription : 우수회원전용코너 처리페이지
' History : 2015.04.20 원승현 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, sqlStr, lp
dim evt_code, img1, img2, orderby
Dim orgsailprice, orgsailsuplycash, orgsailyn, isusing, idx

menupos		= Request("menupos")
mode		= Request("mode")

evt_code	= Request("evt_code")
orderby	= Request("orderby")
img1		= Request("image1")
img2		= Request("image2")
isusing	= Request("isusing")
idx	= Request("idx")

If isusing="" Then
	isusing="Y"
End If

If orderby="" Then
	orderby="99"
End If

'// 트랜젝션 시작
dbget.beginTrans

'// 모드에 따른 분기
Select Case mode
	Case "add"

		'데이터 저장
		sqlStr = "Insert Into db_sitemaster.dbo.tbl_vipcorner " &_
				" (evt_code,pcimg,maing,orderby,isusing,regname,regdate) values " &_
				" ('" & evt_code & "'" &_
				" ,'" & img1 &_
				"' ,'" & img2 &_
				"' ,'" & orderby &_
				"' ,'" & isusing &_
				"' ,'" & session("ssBctId") & "'" &_
				" ,getdate())"
				
		dbget.Execute(sqlStr)

	Case "edit"
		'// 내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_vipcorner " &_
				" Set evt_code='" & evt_code &_
				"' 	,pcimg='" & img1 &_
				"' 	,maing='" & img2 &_
				"' 	,orderby='" & orderby & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,modname='" & session("ssBctId") & "'" &_
				" 	,modifydate=getdate()" &_
				" Where idx='" & idx & "'"
		dbget.Execute(sqlStr)
	Case "delete"
		'// 삭제
		sqlStr = sqlStr & "delete [db_sitemaster].[dbo].tbl_vipcorner " &_
				" Where idx='" & idx & "';" & vbCrLf
		dbget.Execute(sqlStr)

End Select


'// 트랜젝션 검사 및 실행
If Err.Number = 0 Then
        dbget.CommitTrans
Else
        dbget.RollBackTrans
		Alert_return("데이타를 저장하는 도중에 에러가 발생하였습니다.")
		dbget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	<% if mode="delete" then %>
		alert("삭제했습니다.");
	<% else %>
		alert("저장했습니다.");
	<% end if %>
	opener.location.reload();
	self.close();
//	self.location = "vip.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->