<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������Ʈ
' Hieditor : 2014.03.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/tenbyten/invalid/invalid_user_cls.asp"-->

<%
dim idx, gubun, invaliduserid, isusing, regdate, lastupdate, reguserid, lastuserid, comment
dim adminuserid, i, menupos, sql, mode
	idx = requestcheckvar(request("idx"),10)
	gubun = requestcheckvar(request("gubun"),12)
	invaliduserid = requestcheckvar(request("invaliduserid"),32)
	isusing = requestcheckvar(request("isusing"),1)
	menupos = requestcheckvar(request("menupos"),10)
	mode = requestcheckvar(request("mode"),32)
	comment = request("comment")

adminuserid=session("ssBctId")									

'//�ű�����
if mode = "edit" then
	if gubun="" then
		Response.Write "<script type='text/javascript'>alert('�ҷ��������� �����ϴ�.'); history.go(-1);</script>"
		dbget.close() : Response.End
	end if
	if invaliduserid="" then
		Response.Write "<script type='text/javascript'>alert('���̵� �����ϴ�.'); history.go(-1);</script>"
		dbget.close() : Response.End
	end if
	if isusing="" then
		Response.Write "<script type='text/javascript'>alert('��뿩�ΰ� �����ϴ�.'); history.go(-1);</script>"
		dbget.close() : Response.End
	end if
	if checkNotValidHTML(comment) then
		Response.Write "<script type='text/javascript'>alert('�ڸ�Ʈ�� ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���.'); history.go(-1);</script>"
		dbget.close() : Response.End
	end if

	if idx<>"" then
		sql = "update db_user.dbo.tbl_invalid_user set" + vbcrlf
		sql = sql & " gubun='"&gubun&"'" + vbcrlf
		sql = sql & " ,invaliduserid='"&invaliduserid&"'" + vbcrlf
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
		sql = sql & " ,lastupdate=getdate()" + vbcrlf
		sql = sql & " ,lastuserid='"&adminuserid&"'" + vbcrlf
		sql = sql & " ,comment='"&html2db(comment)&"'" + vbcrlf
		sql = sql & " where idx = "&idx&""
		
		'response.write sql
		dbget.execute sql
	else
		sql = "insert into db_user.dbo.tbl_invalid_user(" + vbcrlf
		sql = sql & " gubun, invaliduserid, isusing, regdate, lastupdate, reguserid, lastuserid, comment" + vbcrlf
		sql = sql & " )" + vbcrlf
		sql = sql & " 	select" + vbcrlf
		sql = sql & " 	'"&gubun&"', '"&invaliduserid&"', '"&isusing&"', getdate(), getdate(), '"&adminuserid&"', '"&adminuserid&"', '"&html2db(comment)&"'" + vbcrlf
		sql = sql & " 	from db_user.dbo.tbl_user_n n" + vbcrlf
		sql = sql & " 	left join db_user.dbo.tbl_invalid_user iu" + vbcrlf
		sql = sql & " 		on n.userid=iu.invaliduserid" + vbcrlf
		sql = sql & " 		and iu.isusing='Y'" + vbcrlf
		sql = sql & " 		and iu.gubun='"&gubun&"'" + vbcrlf
		sql = sql & " 		and iu.invaliduserid='"&invaliduserid&"'"		
		sql = sql & " 	where iu.idx is null" + vbcrlf
		sql = sql & " 	and n.userid='"&invaliduserid&"'"

		'response.write sql
		dbget.execute sql
	end if

else
	Response.Write "<script type='text/javascript'>alert('�����ڰ� �����ϴ�.'); history.go(-1);</script>"
	dbget.close() : Response.End
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->

<script type='text/javascript'>
	opener.location.reload();
	self.close();
</script>
