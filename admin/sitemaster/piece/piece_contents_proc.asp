<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 피스 컨텐츠 등록 처리 페이지
' Hieditor : 2017.08.11 유태욱 생성
' Hieditor : 2017.09.05 원승현 추가/수정
' Hieditor : 2017-11-28 이종화 추가/수정
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim idx, fidx, isusing, usertype, makerid, gubun, noticeYN, nickname, etclink, mode
dim itemid, listtext, shorttext, listtitle, pieceidx, startdate, enddate, listimg, tagtext
dim tagtextarr, itemidarr, pieceidxarr, adminid, bannergubun , admintext , state , page
Dim SearchDeal , SearchOpen , SearchState

	mode		=	requestCheckvar(Request("mode"),4)	'NEW:신규, EDIT:수정, SORT:순서변경
	idx			=	requestCheckvar(Request("idx"),16)	'db key idx
	fidx		=	requestCheckvar(Request("fidx"),16)	'순서정렬용 idx
	isusing		=	requestCheckvar(Request("isusing"),1)	'사용여부
	usertype	=	requestCheckvar(Request("usertype"),1)	'(1 :관리자, 2:고객)
	gubun		=	requestCheckvar(Request("gubun"),1)	'1 : 조각, 2 : 파이, 3 : 베스트키워드, 4 : 배너, 5 : 회원조각
	noticeYN	=	requestCheckvar(Request("noticeYN"),2) '오프닝 여부
	etclink		=	requestCheckvar(Request("etclink"),256)	'기타 링크
	itemid		=	requestCheckvar(Request("itemid"),128)	'연관상품코드
	tagtext		=	requestCheckvar(Request("tagtext"),256) '태그
	listtext	=	requestCheckvar(Request("listtext"),500) ' 내용
	shorttext	=	Request("shorttext")	'여는말
	listtitle	=	requestCheckvar(Request("listtitle"),32)	'제목
	pieceidx	=	requestCheckvar(Request("pieceidx"),128)	'파이-연관조각
	startdate	=	requestCheckvar(Request("startdate"),10) & " " & requestCheckvar(Trim(Request("starttime")),8)	'시작일
	enddate		=	requestCheckvar(Request("enddate"),10)	'종료일
	listimg		=	requestCheckvar(Request("con_viewthumbimg"),150) '이미지
	adminid		=	requestCheckvar(Request("adminid"),50) '관리자 아이디
	bannergubun	=	requestCheckvar(Request("bannergubun"),50) '배너 구분값(1:텍스트, 2:이미지)

	admintext	=	requestCheckvar(Request("admintext"),2000) '작업자 지시사항
	state		=	requestCheckvar(Request("state"),1)	' 진행상태
	page		=	requestCheckvar(Request("page"),10)	' 글목록

	SearchDeal = requestCheckvar(request("SearchDeal"), 1) '// 검색 parameter
	SearchOpen = requestCheckvar(request("SearchOpen"), 1) '// 검색 parameter
	SearchState = requestCheckvar(request("SearchState"), 1) '// 검색 parameter

	if usertype="" or isNull(usertype) then
		usertype = 1
	end if

	if nickname="" or isNull(nickname) then
		nickname = "10x10"
	end if

	if enddate="" or isNull(enddate) then
		enddate = "2032-12-31"
	end if	

	if noticeYN="" or isNull(noticeYN) or noticeYN<>"Y" then
		noticeYN = "N"
	end if	
	
	dim sqlstr, getdate, i
	if mode = "EDIT" then
		sqlstr = " update db_sitemaster.dbo.tbl_piece set "
		sqlstr = sqlstr & " shorttext = '"& html2db(shorttext) &"' "
		sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
		sqlstr = sqlstr & " ,startdate = '"& startdate &"' "
		sqlstr = sqlstr & " ,enddate = '"& enddate &"' "
		sqlstr = sqlstr & " ,listtext = '"& html2db(listtext) &"' "
		sqlstr = sqlstr & " ,listtitle = '"& html2db(listtitle) &"' "
		sqlstr = sqlstr & " ,itemid = '"& itemid &"' "
		sqlstr = sqlstr & " ,pieceidx = '"& pieceidx &"' "
		sqlstr = sqlstr & " ,listimg = '"& listimg &"' "
		sqlstr = sqlstr & " ,etclink = '"& etclink &"' "
		sqlstr = sqlstr & " ,noticeYN = 'N' "
		sqlstr = sqlstr & " ,bannergubun = '"&bannergubun&"' "
		sqlstr = sqlstr & " ,lastupdate = getdate() "
		sqlstr = sqlstr & " ,admintext = '"& admintext &"' "
		sqlstr = sqlstr & " ,lastadminid = '"& adminid &"' "
		sqlstr = sqlstr & " ,state = '"& state &"' "
		sqlstr = sqlstr & " where idx = "& idx &" "
		'response.write sqlstr
		dbget.execute sqlstr

		If Trim(noticeYN)="Y" Then
			'// 오프닝으로 설정하여 등록시엔 기존 오프닝값을 전부 N으로 변경, 새로운 오프닝만 설정.
			sqlstr = " update db_sitemaster.dbo.tbl_piece set "
			sqlstr = sqlstr & " noticeYN = 'N' "
			'response.write sqlstr
			dbget.execute sqlstr

			sqlstr = " update db_sitemaster.dbo.tbl_piece set "
			sqlstr = sqlstr & " noticeYN = 'Y' "
			sqlstr = sqlstr & " where idx = "& idx &" "
			'response.write sqlstr
			dbget.execute sqlstr
		End If

		If Trim(tagtext)<>"" Then
			sqlstr = " Delete db_sitemaster.dbo.tbl_piece_tag Where pidx='"&idx&"' "
			dbget.execute sqlstr

			tagtextarr = split(tagtext,",")
			for i = 0 to ubound(tagtextarr)
				sqlstr = " if not exists(select top 1 * from db_sitemaster.dbo.tbl_piece_tag where pidx = '"& idx &"' and tagtext = '"& tagtextarr(i)&"') "
				sqlstr = sqlstr & " insert into db_sitemaster.dbo.tbl_piece_tag (pidx, tagtext)"
				sqlstr = sqlstr & " values (" & idx & " , '" & html2db(tagtextarr(i)) & "' )"
				'response.write sqlstr & "<br>"
			dbget.execute sqlstr
			Next
		End If

		If Trim(itemid)<>"" Then
			sqlstr = " Delete db_sitemaster.dbo.tbl_piece_item Where pidx='"&idx&"' "
			dbget.execute sqlstr

			itemidarr = Split(itemid, ",")
			for i = 0 to ubound(itemidarr)
				sqlstr = " if not exists(select top 1 * from db_sitemaster.dbo.tbl_piece_item where pidx = '"& idx &"' and itemid = '"& itemidarr(i)&"') "
				sqlstr = sqlstr & " insert into db_sitemaster.dbo.tbl_piece_item (pidx, itemid)"
				sqlstr = sqlstr & " values (" & idx & " , '" & itemidarr(i) & "' )"
				'response.write sqlstr & "<br>"
			dbget.execute sqlstr
			next
		End If

	elseif mode = "NEW" then
		sqlstr = "insert into db_sitemaster.dbo.tbl_piece (gubun, bannergubun, noticeYN, listimg, listtext, shorttext, listtitle, adminid, usertype, etclink, itemid, pieceidx, isusing, startdate, enddate, lastupdate, deleteyn , admintext,state)"
		sqlstr = sqlstr & " values (" & gubun & ",'" & bannergubun & "','N' ,'" & listimg & "' , '" & html2db(listtext) & "', '" & html2db(shorttext) & "' , '" & html2db(listtitle) & "' , '" & adminid & "' , " & usertype & ", '" & etclink & "', '" & itemid & "', '" & pieceidx & "', '" & isusing & "', '" & startdate & "', '" & enddate & "', getdate(), 'N' , '"& admintext &"' ,"& state &")"
'		response.write sqlstr
'		response.End
		dbget.execute sqlstr
		sqlstr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_piece') as idx "
		rsget.Open SqlStr, dbget, 1
		
		if Not rsget.Eof then
			idx = rsget("idx")
		end if
		rsget.Close

		sqlstr = " update db_sitemaster.dbo.tbl_piece set "
		sqlstr = sqlstr & " fidx = "& idx &" "
		sqlstr = sqlstr & " where idx = "& idx &" "
		'response.write sqlstr
		dbget.execute sqlstr

		If Trim(noticeYN)="Y" Then
			'// 오프닝으로 설정하여 등록시엔 기존 오프닝값을 전부 N으로 변경, 새로운 오프닝만 설정.
			sqlstr = " update db_sitemaster.dbo.tbl_piece set "
			sqlstr = sqlstr & " noticeYN = 'N' "
			'response.write sqlstr
			dbget.execute sqlstr

			sqlstr = " update db_sitemaster.dbo.tbl_piece set "
			sqlstr = sqlstr & " noticeYN = 'Y' "
			sqlstr = sqlstr & " where idx = "& idx &" "
			'response.write sqlstr
			dbget.execute sqlstr
		End If


		tagtextarr = split(tagtext,",")
		for i = 0 to ubound(tagtextarr)
			sqlstr = " if not exists(select top 1 * from db_sitemaster.dbo.tbl_piece_tag where pidx = '"& idx &"' and tagtext = '"& tagtextarr(i)&"') "
			sqlstr = sqlstr & " insert into db_sitemaster.dbo.tbl_piece_tag (pidx, tagtext)"
			sqlstr = sqlstr & " values (" & idx & " , '" & html2db(tagtextarr(i)) & "' )"
			'response.write sqlstr & "<br>"
		dbget.execute sqlstr
		Next
		itemidarr = Split(itemid, ",")
		for i = 0 to ubound(itemidarr)
			sqlstr = " if not exists(select top 1 * from db_sitemaster.dbo.tbl_piece_item where pidx = '"& idx &"' and itemid = '"& itemidarr(i)&"') "
			sqlstr = sqlstr & " insert into db_sitemaster.dbo.tbl_piece_item (pidx, itemid)"
			sqlstr = sqlstr & " values (" & idx & " , '" & itemidarr(i) & "' )"
			'response.write sqlstr & "<br>"
		dbget.execute sqlstr
		next		

	elseif mode = "SORT" then
	end If
	
	Dim pageParam
	If mode = "EDIT" Then 
		If page = "" Then page = 1
		pageParam = "?page="&page&"&deal="& SearchDeal &"&open="& SearchOpen &"&state="& SearchState &"&research=on"
	End If 
%>

<script language = "javascript">
    alert('저장되었습니다.');
    opener.location.href="/admin/sitemaster/piece/piece_terminal.asp<%=pageParam%>";
    window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->