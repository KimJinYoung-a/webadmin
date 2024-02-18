<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : dogiftday_item.asp
' Discription : 기프트 데이 아이템 저장
' History : 2014.03.31 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, maintitle, linkinfo , subtitle
Dim qdate
Dim startdate , enddate , gubun , isusing

Dim listidx , itemid , sortnum , saleper
Dim subidx

menupos	= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun		= Request("gubun")
allusing		= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
maintitle	= html2db(Request("maintitle"))
subtitle		= html2db(Request("subtitle"))
linkinfo	 	= html2db(Request("linkinfo"))

qdate	=	Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate	= Request("EndDate")& " " &Request("eTm")

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("subItemid")
sortnum	= Request("sortnum")

if SortNo="" then	SortNo = html2db(Request("SortNo"))

''response.write gubun &"<br>"
'response.write mode &"<br>"
'response.write idx &"<br>"
'response.write itemid &"<br>"
'response.write subidx &"<br>"
'response.write sortnum &"<br>"
''response.write startdate &"<br>"
''response.write enddate &"<br>"
'response.End

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into db_board.dbo.tbl_giftday_master_item " &_
					" (masteridx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
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
	self.location = "giftday_item.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
