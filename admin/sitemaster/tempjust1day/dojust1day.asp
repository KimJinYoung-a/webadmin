<%@ language=vbscript %>
<% option explicit %>
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

Dim listidx , itemid , sortnum , is1day
Dim subidx
Dim todayban , extraurl
Dim subtitle , saleper '//주말or연휴특가용
Dim maxsaleper
Dim pcimage, mobileimage, pclinkurl, mobilelinkurl, price



menupos		= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun		= Request("gubun")
allusing	= Request("allusing")
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
itemid		= Request("Itemid")
sortnum		= Request("sortnum")
saleper		= Request("saleper")
maxsaleper		= Request("maxsaleper")
pcimage		= Request("pcimage")
mobileimage		= Request("mobileimage")
pclinkurl		= Request("pclinkurl")
mobilelinkurl		= Request("mobilelinkurl")
price		= Request("price")


if SortNo="" then	SortNo = Request("SortNo")


If mode = "add" Then '//날짜 체크
	Dim itemcount
	sqlStr = "select count(*) from [db_temp].[dbo].[tbl_temp_just1day_list] where startdate = '"& startdate &"' and isusing = 'Y'" 
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		itemcount = rsget(0)
	end If
	rsget.Close

	If itemcount > 0 Then 
	%>
	<script style="css/text">
		alert("동일날짜 혹은 동일시간 대에 시작 하는 컨텐츠가 있습니다.");
		self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
	</script>
	<%
	Response.end
	End If
End If 

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_temp].[dbo].[tbl_temp_just1day_item] " &_
					" (listidx, title, itemid, pclinkurl, mobilelinkurl, pcimage, mobileimage, price, saleper, adminid, isusing, sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & title &"'" &_
					" ,'" & itemid &"'" &_
					" ,'" & pclinkurl &"'" &_
					" ,'" & mobilelinkurl &"'" &_
					" ,'" & pcimage &"'" &_
					" ,'" & mobileimage &"'" &_
					" ,'" & price &"'" &_
					" ,'" & saleper &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_temp].[dbo].tbl_temp_just1day_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,title='" & title & "'" &_
				" 	,pclinkurl='" & pclinkurl & "'" &_
				" 	,mobilelinkurl='" & mobilelinkurl & "'" &_
				" 	,pcimage='" & pcimage & "'" &_
				" 	,mobileimage='" & mobileimage & "'" &_
				" 	,price='" & price & "'" &_
				" 	,saleper='" & saleper & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_temp].[dbo].tbl_temp_just1day_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_temp].[dbo].tbl_temp_just1day_list " &_
					" (title, startdate , enddate , adminid, isusing , maxsaleper) values " &_
					" ('" & title &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing & "'" &_
					" ,'" & maxsaleper & "'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_temp].[dbo].tbl_temp_just1day_list " &_
				" Set title='" & title & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,maxsaleper='" & maxsaleper & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
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
