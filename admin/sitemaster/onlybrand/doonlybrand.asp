<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : PC메인관리 온니브랜드
' History : 서동석 생성
'			2022.07.01 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, title, linkinfo
Dim qdate
Dim startdate , enddate , gubun , isusing, paramisusing, orderby

Dim listidx , itemid , sortnum
Dim subidx
Dim bannerlink, banneralt, maincopy, bannerImg, frequest
menupos		= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
bannerImg		= Request("bannerImg")
bannerlink	= Request("bannerlink")
banneralt	= Request("banneralt")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
title		= Newhtml2db(Request("title"))
linkinfo	= Request("linkinfo")

qdate		= Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate		= Request("EndDate")& " " &Request("eTm")

paramisusing= Request("paramisusing")

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("subItemid")
orderby		= Request("orderby")

maincopy	= Request("maincopy")
frequest	= Request("request")


if orderby="" then	orderby = Request("orderby")


If mode = "add" Then '//날짜 체크
	Dim itemcount
	sqlStr = "select count(*) from [db_sitemaster].[dbo].tbl_pc_main_onlyBrand_list where startdate = '"& startdate &"' and isusing = 'Y'" 
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		itemcount = rsget(0)
	end If
	rsget.Close

'	If itemcount > 0 Then 
	%>
	<script style="css/text">
//		alert("동일날짜 혹은 동일시간 대에 시작 하는 컨텐츠가 있습니다.");
//		self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
	</script>
	<%
'	Response.end
'	End If
End If 

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_onlyBrand_item " &_
					" (listidx, itemid , orderby) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & orderby &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_onlyBrand_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,orderby='" & orderby & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_onlyBrand_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		if bannerlink <> "" and not(isnull(bannerlink)) then
			bannerlink = ReplaceBracket(bannerlink)
		end If
		if banneralt <> "" and not(isnull(banneralt)) then
			banneralt = ReplaceBracket(banneralt)
		end If
		if maincopy <> "" and not(isnull(maincopy)) then
			maincopy = ReplaceBracket(maincopy)
		end If
		if frequest <> "" and not(isnull(frequest)) then
			frequest = ReplaceBracket(frequest)
		end If

		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_onlyBrand_list " &_
					" (bannerimg, bannerlink, banneralt, maincopy, orderby, startdate , enddate , adminid, request, isusing) values " &_
					" ('" & bannerimg &"'" &_
					" ,'" & html2db(bannerlink) & "'" &_
					" ,'" & html2db(banneralt) & "'" &_
					" ,'" & html2db(maincopy) & "'" &_
					" ,'" & orderby & "'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & html2db(frequest) &"'" &_
					" ,'" & isusing & "'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		if bannerlink <> "" and not(isnull(bannerlink)) then
			bannerlink = ReplaceBracket(bannerlink)
		end If
		if banneralt <> "" and not(isnull(banneralt)) then
			banneralt = ReplaceBracket(banneralt)
		end If
		if maincopy <> "" and not(isnull(maincopy)) then
			maincopy = ReplaceBracket(maincopy)
		end If
		if frequest <> "" and not(isnull(frequest)) then
			frequest = ReplaceBracket(frequest)
		end If

		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_onlyBrand_list " &_
				" Set bannerimg='" & bannerimg & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,bannerlink='" & html2db(bannerlink) & "'" &_
				" 	,banneralt='" & html2db(banneralt) & "'" &_
				" 	,maincopy='" & html2db(maincopy) & "'" &_
				" 	,orderby='" & orderby & "'" &_
				" 	,request='" & html2db(frequest) & "'" &_
				" Where idx=" & idx

		'response.write sqlStr
		'response.end
		dbget.Execute(sqlStr)

	Case "quickadd"
		'신규 등록
		Dim extratime1 , extratime2 , titledate
			
		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_just1day_list " &_
				" (title, startdate , enddate , adminid) values " &_
				" ('수정요망 --&gt;" & titledate &" 시작 mdpick 입니다'" &_
				" ,'" & extratime1 &"'" &_
				" ,'" & extratime2 &"'" &_
				" ,'" & session("ssBctId") &"')"
		dbget.Execute(sqlStr)

End Select

%>
<% If mode = "subadd"  Or mode = "submodify" then%>
<script>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
	 self.close();        // 팝업창 닫기
//-->
</script>
<% Else %>
<script type='text/javascript'>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
