<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN(����������->�̽�����)
'	History		: 2014.07.09 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhikerCls.asp"-->

<%
dim idx, mode, sortnum, isusing, gubun, vol1, vol2, issueimg
dim evt_title, startdate, enddate, imghtmltext
	idx = request("idx")
	vol1 = request("vol1")
	vol2 = request("vol2")
	mode = request("mode")
	isusing = request("isusing")
	issueimg = request("issueimg")
	gubun = request("SearchGubun")
	sortnum =  requestcheckvar(request("sortnum"),5)
	evt_title = requestcheckvar(request("evt_title"),150)
	imghtmltext = request("imghtmltext")

	startdate	= Request("StartDate")& " " &Request("sTm")
	enddate	= Request("EndDate")& " " &Request("eTm")

	dim sqlstr, getdate
	if mode = "EDIT" then
		
		sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_list set " '��������϶� db������Ʈ
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,vol1 = '"& vol1 &"' "
		sqlstr = sqlstr & " ,vol2 = '"& vol2 &"' "
		sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
		sqlstr = sqlstr & " ,issueimg = '"& issueimg &"' "
		sqlstr = sqlstr & " ,hic_title = '"& html2db(evt_title) &"' "
		sqlstr = sqlstr & " ,imghtmltext = '"& html2db(imghtmltext) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbget.execute sqlstr

	elseif mode = "NEW" then
		
		'//html2db  ������
		sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_list (gubun, hic_title, sdate, edate, isusing, imghtmltext ,sortnum, vol1, vol2, issueimg, regdate)" ' �ű��Է¸��
		sqlstr = sqlstr & " values ('" & gubun & "' , '" & html2db(evt_title) & "' , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' , '" & html2db(imghtmltext) & "'"
		sqlstr = sqlstr & " ,'" &sortnum & "', '" &vol1 & "', '" &vol2 & "', '" &issueimg & "', getdate())"

		'response.write sqlstr
		dbget.execute sqlstr
	end if
%>
<script language = "javascript">
	alert("����Ǿ����ϴ�."); //����Ǿ����ϴ� ��� �޽������
	opener.location.reload(); //��â�� ��� �θ�â�� ���ε���
	self.close();			  //��â�� ����
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->