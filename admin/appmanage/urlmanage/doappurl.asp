<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doappurl.asp
' Discription : appurl ó�� ������
' History : 2014-08-19 ����ȭ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
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
		'�ű� ���
		'Ʈ������
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

			'sqlStr = "select SCOPE_IDENTITY() From db_sitemaster.dbo.tbl_AppUrlList "		'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
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
				Response.write "<script>alert('�ű� ��� �Ϸ�.');</script>"
				Response.write "<script>opener.location='http://webadmin.10x10.co.kr/admin/appmanage/urlmanage/?menupos=1764';</script>"
				Response.write "<script>top.window.close();</script>"
				dbget.close()	:	response.End
		ELSE
			dbget.RollBackTrans
			Response.write "<script>alert('������ ó���� ������ �߻��Ͽ����ϴ�.');</script>"
			Response.write "<script>window.opener.history.go(0);</script>"
			Response.write "<script>parent.self.close();</script>"
		END IF

		
	Else
		'���� ����
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

		Response.write "<script>alert('���� �Ϸ�');</script>"
		Response.write "<script>opener.location='http://webadmin.10x10.co.kr/admin/appmanage/urlmanage/?menupos=1764';</script>"
		Response.write "<script>top.window.close();</script>"
	End If

	'// ������� ����
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
