<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ��ǰ�� ��� �귣�� ���� ���,���� ó��
'	History		: 2017.01.20 ���¿� ����
'				  2022.07.12 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/board/brand_noticeCls.asp"-->

<%
dim idx, mode, isusing , gubun , notice_title, notice_text, brandid, menupos, startdate, enddate, infiniteregyn, loginuserid
dim sqlstr, getdate
	menupos				=	requestcheckvar(request("menupos"),10)
	loginuserid			=	session("ssBctId")
	idx					=	requestcheckvar(request("idx"),10)
	mode				=	requestcheckvar(request("mode"),10)
	isusing				=	requestcheckvar(request("isusing"),2)
	notice_title		=	requestcheckvar(request("notice_title"),128)
	notice_text			=	requestcheckvar(request("notice_text"),256)
	gubun				=	requestcheckvar(request("SearchGubun"),2)
	brandid				=	requestcheckvar(request("brandid"),32)
	startdate			= Request("StartDate")& " " &Request("sTm")
	enddate				= Request("EndDate")& " " &Request("eTm")
	infiniteregyn		= requestcheckvar(request("infiniteregyn"),2)

	' �±״� �Է°����ϰ� ������. ��ü���� Ŭ���� ����.
	'if notice_title <> "" and not(isnull(notice_title)) then
	'	notice_title = ReplaceBracket(notice_title)
	'end If
	'if notice_text <> "" and not(isnull(notice_text)) then
	'	notice_text = ReplaceBracket(notice_text)
	'end If
	if (checkNotValidHTML(notice_title) = True) or (checkNotValidHTML(notice_text) = True)  then
		Alert_return("HTML�� ����Ͻ� �� �����ϴ�.")
		Response.Write "<script type='text/javascript'>self.close();</script>"
		dbget.close() : Response.End	
	End If

	if infiniteregyn <>"" then
		infiniteregyn = "Y"
	else
		infiniteregyn = "N"
	end if

	if mode = "EDIT" then

		sqlstr = " update db_board.dbo.tbl_brand_notice_list set " '��������϶� db������Ʈ
		sqlstr = sqlstr & " isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,edate = '"& enddate &"' "
		sqlstr = sqlstr & " ,sdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,gubun = '"& gubun &"' "
		sqlstr = sqlstr & " ,infiniteregyn = '"& infiniteregyn &"' "
		sqlstr = sqlstr & " ,notice_title = '"& html2db(notice_title) &"' "
		sqlstr = sqlstr & " ,notice_text = '"& html2db(notice_text) &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "

		'response.write sqlstr
		dbget.execute sqlstr

	elseif mode = "NEW" then
		sqlstr = "insert into db_board.dbo.tbl_brand_notice_list (gubun, sdate, edate, isusing , regdate, notice_title, notice_text, makerid, brandid, infiniteregyn)" ' �ű��Է¸��
		sqlstr = sqlstr & " values ('" & gubun & "'  , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' "
		sqlstr = sqlstr & " , getdate(), '" & html2db(notice_title) & "', '" & html2db(notice_text) & "', '" & loginuserid & "', '" & brandid & "', '" & infiniteregyn & "') "

		'response.write sqlstr
		'response.end
		dbget.execute sqlstr
	end if
%>

<script type='text/javascript'>
	alert("����Ǿ����ϴ�.");
	opener.location.reload();
	self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
