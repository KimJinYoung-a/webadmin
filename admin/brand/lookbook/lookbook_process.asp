<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.19 ������ ����
'			2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/lookbookCls.asp"-->
<%
dim idx, makerid, title, state, mainimg, isusing, sortNo, sqlStr, mode, adminid, comment, detailitemcnt
dim existsbrandcnt
	idx 	= request("idx")
	makerid	= request("makerid")
	title 	= request("title")
	state 	= request("state")
	mainimg	= request("mainimg")
	isusing	= request("isusing")
	sortNo 	= request("sortNo")
	comment 	= request("comment")
	mode 	= request("mode")
	menupos 	= request("menupos")
	
adminid = session("ssBctId")
detailitemcnt = 0
existsbrandcnt = 0

If mode = "I" Then
	sqlStr = "SELECT count(*) as cnt"
	sqlStr = sqlStr & " from db_user.dbo.tbl_user_c"
	sqlStr = sqlStr & " WHERE userid='"&makerid&"'"
	
	'response.write sqlStr & "<BR>"
	rsget.Open sqlStr, dbget, 1
    If Not rsget.Eof then
    	existsbrandcnt = rsget("cnt")
	End If
    rsget.Close

	If existsbrandcnt = 0 Then
		Response.Write  "<script language='javascript'>"
		Response.Write  "	alert('�ش�Ǵ� �귣�尡 �����ϴ�.');"
		Response.Write  "	location.replace('/admin/brand/lookbook/lookbookModify.asp?menupos="&menupos&"');"
		Response.Write  "</script>"
		dbget.close()	:	response.End
	End If
	

	if checkNotValidHTML(comment) or checkNotValidHTML(title) then
	%>
		<script>
			alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if
	
	sqlStr = "INSERT INTO db_brand.dbo.tbl_street_LookBook_Master (" + vbcrlf
	sqlStr = sqlStr & " makerid, title, state, mainimg, isusing, regadminid, lastadminid, comment)" + vbcrlf
	sqlStr = sqlStr & " 	select" + vbcrlf
	sqlStr = sqlStr & " 	c.userid as makerid, '"&html2db(title)&"', 3, '"&mainimg&"', '"&isusing&"','"&adminid&"'" + vbcrlf
	sqlStr = sqlStr & " 	,'"&adminid&"', '"&html2db(comment)&"'" + vbcrlf
	sqlStr = sqlStr & " 	from db_user.dbo.tbl_user_c c" + vbcrlf
	sqlStr = sqlStr & " 	where userid='"& makerid &"'"

	'response.write sqlStr & "<BR>"
	dbget.execute sqlStr

	sqlStr = "select IDENT_CURRENT('db_brand.dbo.tbl_street_LookBook_Master') as idx"
	rsget.Open sqlStr, dbget, 1

	If Not rsget.Eof then
		idx = rsget("idx")
	End If
	rsget.close

	Response.Write "<script language='javascript'>alert('����Ǿ����ϴ�');location.replace('/admin/brand/lookbook/lookbookModify.asp?idx="&idx&"&menupos="&menupos&"');</script>"

ElseIf mode = "U" Then
	if idx="" then
	%>
		<script>
			alert('IDX�� �����ϴ�.');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if
	

	if checkNotValidHTML(comment) or checkNotValidHTML(title) then
	%>
		<script>
			alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if
		
	sqlStr = "UPDATE db_brand.dbo.tbl_street_LookBook_Master SET" + VBCRLF
	sqlStr = sqlStr & " makerid = '"& makerid &"'" + VBCRLF
	sqlStr = sqlStr & " , title = '"& html2db(title) &"'" + VBCRLF
	sqlStr = sqlStr & " ,state = "&state&"" + VBCRLF
	sqlStr = sqlStr & " ,mainimg = '"&mainimg&"'" + VBCRLF
	sqlStr = sqlStr & " , isusing = '"&isusing&"'" + VBCRLF
	sqlStr = sqlStr & " ,sortNo = '"&sortNo&"'" + VBCRLF
	sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf
	sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf
	sqlStr = sqlStr & " , comment = '"& html2db(comment) &"'" + VBCRLF
	sqlStr = sqlStr & " where idx ='" & Cstr(idx) & "'"

	'response.write sqlStr & "<BR>"	
	dbget.execute sqlStr

	response.write "<script language='javascript'>"
	response.write "	alert('����Ǿ����ϴ�');"
	response.write "	location.replace('/admin/brand/lookbook/lookbookModify.asp?idx="&idx&"&menupos="&menupos&"');"
	response.write "</script>"	

'/���� ����
elseif mode="chstate" then
	if idx="" then
	%>
		<script>
			alert('���� �����ϴ�.');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if

	
	sqlStr = "SELECT count(*) as cnt"
	sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master as M"
	sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_LookBook_Detail as D"
	sqlStr = sqlStr & " 	on M.idx=D.masteridx"
	sqlStr = sqlStr & " WHERE m.idx="&idx&" and D.isusing='Y' "
	
	'response.write sqlStr & "<BR>"
	rsget.Open sqlStr, dbget, 1
    If Not rsget.Eof then
    	detailitemcnt = rsget("cnt")
	End If
    rsget.Close
	
	if state="3" or state="7" then
		If detailitemcnt = "0" Then
			Response.Write  "<script language='javascript'>"
			Response.Write  "	alert('LookBook���̹����� ��ϵǾ� ���� �ʽ��ϴ�.\n����Ͻð� �ٽ� �õ� �ϼ���.');"
			Response.Write  "	history.go(-1);"
			Response.Write  "</script>"
			dbget.close()	:	response.End		
		End If
	End If	

	sqlStr = "UPDATE db_brand.dbo.tbl_street_LookBook_Master SET" + VBCRLF
	sqlStr = sqlStr & " state = "&state&"" + VBCRLF
	sqlStr = sqlStr & " where idx ='" & Cstr(idx) & "'"

	'response.write sqlStr & "<BR>"	
	dbget.execute sqlStr

	response.write "<script language='javascript'>"
	response.write "	alert('����Ǿ����ϴ�');"
	response.write "	location.replace('/admin/brand/lookbook/lookbookModify.asp?idx="&idx&"&menupos="&menupos&"');"
	response.write "</script>"	
else
	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->