<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 상품상세 상단 브랜드 공지 등록,수정 처리
'	History		: 2017.01.20 유태욱 생성
'				  2022.07.12 한용민 수정(isms취약점보안조치, 표준코드로변경)
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

	' 태그는 입력가능하게 해주자. 업체에서 클레임 들어옴.
	'if notice_title <> "" and not(isnull(notice_title)) then
	'	notice_title = ReplaceBracket(notice_title)
	'end If
	'if notice_text <> "" and not(isnull(notice_text)) then
	'	notice_text = ReplaceBracket(notice_text)
	'end If
	if (checkNotValidHTML(notice_title) = True) or (checkNotValidHTML(notice_text) = True)  then
		Alert_return("HTML을 사용하실 수 없습니다.")
		Response.Write "<script type='text/javascript'>self.close();</script>"
		dbget.close() : Response.End	
	End If

	if infiniteregyn <>"" then
		infiniteregyn = "Y"
	else
		infiniteregyn = "N"
	end if

	if mode = "EDIT" then

		sqlstr = " update db_board.dbo.tbl_brand_notice_list set " '수정모드일때 db업데이트
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
		sqlstr = "insert into db_board.dbo.tbl_brand_notice_list (gubun, sdate, edate, isusing , regdate, notice_title, notice_text, makerid, brandid, infiniteregyn)" ' 신규입력모드
		sqlstr = sqlstr & " values ('" & gubun & "'  , '" & startdate & "' , '" & enddate & "' , '" & isusing & "' "
		sqlstr = sqlstr & " , getdate(), '" & html2db(notice_title) & "', '" & html2db(notice_text) & "', '" & loginuserid & "', '" & brandid & "', '" & infiniteregyn & "') "

		'response.write sqlstr
		'response.end
		dbget.execute sqlstr
	end if
%>

<script type='text/javascript'>
	alert("저장되었습니다.");
	opener.location.reload();
	self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
