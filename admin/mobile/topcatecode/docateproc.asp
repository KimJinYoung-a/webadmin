<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : domdpick.asp
' Discription : mdpick 처리 페이지
' History : 2013.12.16 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, sqlStr , totcnt
dim idx, isusing
Dim gcode ,  dcode

menupos	= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
idx			= getNumeric(Request("idx"))
gcode		= getNumeric(Request("gcode"))
dcode		= getNumeric(Request("dcode"))

'// 모드에 따른 분기
Select Case mode
	Case "add"

		SqlStr = "select count(*) "
        SqlStr = SqlStr + " from db_sitemaster.[dbo].[tbl_mobile_main_topsubcode] "
        SqlStr = SqlStr + " where dispcode=" + CStr(dcode) 
		rsget.Open SqlStr, dbget, 1
        if Not rsget.Eof then
            totcnt = rsget(0)
        end if
        rsget.close

		If totcnt = 0 then
			'신규 등록
			sqlStr = "Insert Into db_sitemaster.dbo.tbl_mobile_main_topsubcode " &_
						" (gnbcode, dispcode , adminid , isusing ) values " &_
						" ('" & gcode &"'" &_
						" ,'" & dcode &"'" &_
						" ,'" & session("ssBctId") &"'" &_
						" ,'" & isusing &"'" &_
						")"
			'response.write sqlStr
			dbget.Execute(sqlStr)
		Else
			Response.Write "<script>alert('이미 등록된 카테고리 입니다.'); history.back(-1);</script>"
			dbget.close() : Response.End
		End If 

	Case "modify"
		'내용 수정
		sqlStr = "Update db_sitemaster.dbo.tbl_mobile_main_topsubcode " &_
				" Set gnbcode='" & gcode & "'" &_
				" 	,dispcode='" & dcode & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)
End Select

%>
<script>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
	 self.close();        // 팝업창 닫기
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->