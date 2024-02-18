<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : PC메인관리 위시베스트
' History : 서동석 생성
'			2022.07.04 한용민 수정(isms취약점조치)
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
Dim startdate , enddate , gubun , isusing, paramisusing

Dim listidx , itemid , sortnum
Dim subidx
Dim bannernameeng, bannernamekor, subcopy, linkurl, bannerImg, altname
Dim maincopy1, maincopy2



menupos		= requestCheckvar(getNumeric(Request("menupos")),10)
isusing		= requestCheckvar(Request("isusing"),1)
mode		= requestCheckvar(Request("mode"),32)
maincopy1		= Request("maincopy1")
maincopy2		= Request("maincopy2")
linkurl	= Request("linkurl")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= requestCheckvar(getNumeric(Request("idx")),10)
title		= Newhtml2db(Request("title"))
linkinfo	= Request("linkinfo")

qdate		= Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate		= Request("EndDate")& " " &Request("eTm")

paramisusing= Request("paramisusing")

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("subItemid")
sortnum		= requestCheckvar(getNumeric(Request("sortnum")),10)

if SortNo="" then	SortNo = Request("SortNo")


If mode = "add" Then '//날짜 체크
	Dim itemcount
	sqlStr = "select count(*) from [db_sitemaster].[dbo].tbl_pc_main_wishbest_list where startdate = '"& startdate &"' and isusing = 'Y'" 
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		itemcount = rsget(0)
	end If
	rsget.Close
End If 

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_wishbest_item " &_
					" (listidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_wishbest_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_wishbest_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		if maincopy1 <> "" and not(isnull(maincopy1)) then
			maincopy1 = ReplaceBracket(maincopy1)
		end If
		if maincopy2 <> "" and not(isnull(maincopy2)) then
			maincopy2 = ReplaceBracket(maincopy2)
		end If

		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_wishbest_list " &_
					" (maincopy1, maincopy2, linkurl, sortnum, startdate , enddate , adminid, isusing) values " &_
					" ('" & maincopy1 &"'" &_
					" ,'" & maincopy2 & "'" &_
					" ,'" & linkurl & "'" &_
					" ,'" & sortnum & "'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing & "'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		if maincopy1 <> "" and not(isnull(maincopy1)) then
			maincopy1 = ReplaceBracket(maincopy1)
		end If
		if maincopy2 <> "" and not(isnull(maincopy2)) then
			maincopy2 = ReplaceBracket(maincopy2)
		end If

		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_wishbest_list " &_
				" Set maincopy1='" & maincopy1 & "'" &_
				" 	,maincopy2='" & maincopy2 & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,linkurl='" & linkurl & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
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
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
