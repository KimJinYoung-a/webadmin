<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doappurl.asp
' Discription : appurl 처리 페이지
' History : 2014-08-19 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim idx , strMsg , tmpeCode , sqlStr
Dim urltitle , urldiv , urlcontent , isusing , catecode
Dim tempurl 

	IF application("Svr_Info") = "Dev" THEN
		tempurl = "http://testm.10x10.co.kr:8080/apps/link/?"
	Else
		tempurl = "http://m.10x10.co.kr/apps/link/?"
	End If

	idx = Request("idx")
	urltitle = Request("urltitle")
	urldiv = Request("urldiv")
	urlcontent = Request("urlcontent")
	isusing = Request("isusing")
	catecode = Request("catecode")

	If idx = "" then
		'신규 등록
		'트랜젝션
		dbget.beginTrans
			sqlStr = "Insert Into db_sitemaster.dbo.tbl_AppUrlList " &_
						" (urldiv , urltitle , urlcontent , isusing , catecode ) values " &_
						" ('" & urldiv &"'" &_
						" ,'" & urltitle &"'" &_
						" ,'" & Server.UrlEncode(urlcontent) &"'" &_
						" ,'" & isusing &"'" &_
						" ,'" & catecode &"'" &_
						")"
			dbget.Execute(sqlStr)

		IF Err.Number = 0 Then

			'sqlStr = "select SCOPE_IDENTITY() From db_sitemaster.dbo.tbl_AppUrlList "		'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
			sqlStr = "select SCOPE_IDENTITY()"
			rsget.Open sqlStr, dbget, 0
				tmpeCode = rsget(0)
			rsget.Close

			sqlStr = "Update db_sitemaster.dbo.tbl_AppUrlList " &_
				" Set urlcomplete='" & tempurl & tmpeCode & "'+ replace(convert(varchar(10),regdate,120),'-','')" &_
				" Where idx=" & tmpeCode
				'response.write sqlStr
				dbget.Execute(sqlStr)
			
				dbget.CommitTrans
				Response.write "<script>alert('신규 등록 완료.');</script>"
				Response.write "<script>opener.location='http://webadmin.10x10.co.kr/admin/appmanage/urlmanage/?menupos=1764';</script>"
				Response.write "<script>top.window.close();</script>"
				dbget.close()	:	response.End
		ELSE
			dbget.RollBackTrans
			Response.write "<script>alert('데이터 처리에 문제가 발생하였습니다.');</script>"
			Response.write "<script>window.opener.history.go(0);</script>"
			Response.write "<script>parent.self.close();</script>"
		END IF

		
	Else
		'내용 수정
		sqlStr = "Update db_sitemaster.dbo.tbl_AppUrlList " &_
				" Set urldiv='" & urldiv & "'" &_
				" 	,urltitle='" & urltitle & "'" &_
				" 	,urlcontent='" & Server.UrlEncode(urlcontent) & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,catecode='" & catecode & "'" &_
				" 	,lastupdate=getdate()" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

		Response.write "<script>alert('수정 완료');</script>"
		Response.write "<script>opener.location='http://webadmin.10x10.co.kr/admin/appmanage/urlmanage/?menupos=1764';</script>"
		Response.write "<script>top.window.close();</script>"
	End If

	'// 목록으로 복귀
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
