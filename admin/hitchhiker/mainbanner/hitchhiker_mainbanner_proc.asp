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
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_mainbannerCls.asp"-->

<%
dim idx, mode, isusing , gubun , sortnum , linkurl , layerpopurl , con_viewthumbimg
dim startdate, enddate
	idx = request("idx")
	mode = request("mode")
	isusing = request("isusing")
	gubun = request("SearchGubun")
	sortnum =  requestcheckvar(request("sortnum"),3)
	linkurl =  requestcheckvar(request("linkurl"),256)
	con_viewthumbimg = requestCheckvar(Request("con_viewthumbimg"),150)
	
	startdate	= Request("StartDate")& " " &Request("sTm")
	enddate	= Request("EndDate")& " " &Request("eTm")

	dim sqlstr, getdate
	if mode = "EDIT" then
		
		sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list set " '��������϶� db������Ʈ
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,gubun = '"& gubun &"' "
		sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
		sqlstr = sqlstr & " ,linkurl = '"& html2db(linkurl) &"' "
		sqlstr = sqlstr & " ,con_viewthumbimg = '"& html2db(con_viewthumbimg) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbget.execute sqlstr

	elseif mode = "NEW" then
		
		'//html2db  ������
		sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list (gubun, sdate, edate, isusing , linkurl, sortnum, con_viewthumbimg, regdate)" ' �ű��Է¸��
		sqlstr = sqlstr & " values ('" & gubun & "'  , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' , '" & html2db(linkurl) & "'"
		sqlstr = sqlstr & " ,'" &sortnum & "', '" & con_viewthumbimg & "' , getdate())"

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