<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim keyval
dim point
keyval = requestCheckVar(request.Form("kv"),32)
point = requestCheckVar(request.Form("pt"),1)

dim strSQL
	strSQL =" UPDATE db_cs.dbo.tbl_myqna " &_
			" set md5key = null " &_
			" ,EvalPoint= '"& point &"'" &_
			" ,EvalDate = getdate() "&_
			" WHERE md5key='"& keyval &"' " 
	dbget.Execute(strSQL)
	
	Alert_close("감사합니다")
	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->