<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/resending_reportcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
function drawSelectBoxCSCommComboALL(selectBoxName,selectedId,groupCode,onChangefunction)
   dim tmp_str,sqlStr
%>
     <select class="select" name="<%=selectBoxName%>" <%= onChangefunction %> onchange="divChange();">
     <option value='' <%if selectedId="" then response.write " selected" %> >선택</option>
<%
       sqlStr = " select comm_cd,comm_name "
       sqlStr = sqlStr + " from  "
       sqlStr = sqlStr + " [db_cs].[dbo].tbl_cs_comm_code "
       sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
       ''sqlStr = sqlStr + " and comm_isDel='N' "
       sqlStr = sqlStr + " order by comm_cd "

       rsget.Open sqlStr,dbget,1

       if  not rsget.EOF  then
           do until rsget.EOF
               if LCase(selectedId) = LCase(rsget("comm_cd")) then
                   tmp_str = " selected"
               end if
               response.write("<option value='" & rsget("comm_cd") & "' " & tmp_str & ">" + db2html(rsget("comm_name")) + " </option>")
               tmp_str = ""
               rsget.MoveNext
           loop
       end if
       rsget.close
%>
       </select>
<%
End function


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr
dim i,divcd
dim checkJumunSite, jumunsite


divcd = request("divcd")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())

checkJumunSite = request("checkJumunSite")
jumunsite = request("jumunsite")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = yyyy2 + "-" + mm2 + "-" + dd2

dim oreport

set oreport = new CReportMaster
oreport.FRectStart = startdateStr
oreport.FRectEnd =  nextdateStr
oreport.FRectDivcd =  divcd

if (checkJumunSite = "Y") then
	oreport.FRectJumunSite =  jumunsite
end if

if (divcd<>"") then
'    oreport.SearchReportByGubun
else
'    oreport.SearchReport
end if


'dim flashvar
'flashvar = "startdate=" + startdateStr + "&enddate=" + nextdateStr + "&divcd=" + divcd

dim totalcount
totalcount = 0


Dim rs
rs = oreport.getCSReport
%>

<script language='javascript'>
function divChange()
{
    frm.submit();
}

function popCsListView(divCd, finishDate, gubun01, gubun02)
{
	if (finishDate)
	{
		var finishDate1 = finishDate;
		var finishDate2 = finishDate;
	}
	else
	{
		var f = document.frm;
		var finishDate1 = f.yyyy1.value + "-" + f.mm1.value + "-" + f.dd1.value;
		var finishDate2 = f.yyyy2.value + "-" + f.mm2.value + "-" + f.dd2.value;
	}
	var url = "/admin/csreport/popCsListView.asp?divCd=" + divCd + "&finishDate1=" + finishDate1 + "&finishDate2=" + finishDate2 + "&gubun01=" + gubun01 + "&gubun02=" + gubun02;
	var popwin = window.open(url,"popCsListView","width=900, height=700, left=0, top=0, scrollbars=yes, resizable=yes, status=yes");
	popwin.focus();

}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30">
			기간(완료일) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			CS구분 : <% Call drawSelectBoxCSCommComboALL("divcd",divcd,"Z001","") %>
			&nbsp;
			<input type="checkbox" name="checkJumunSite" value="Y" <% if checkJumunSite="Y" then response.write "checked" %>>
			특정사이트 : <% DrawSelectExtSiteName "jumunsite", jumunsite %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30">
			* 1시간 지연 데이터입니다.
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

 <p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="120">
		<%If divCd <> "" Then %>
			날짜
		<%Else %>
			유형
		<%End If%>
		</td>
        <td colspan="3">공통</td>
        <td colspan="4">상품관련</td>
        <td colspan="5">물류관련</td>
        <td colspan="3">택배사관련</td>
		<td rowspan="2" width="50"><b>합계</b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<%
Dim trEndStart
Dim prevRow
Dim sumCnt, sumTotCnt
Dim totCnt(50), divLen, divNow
divLen = 15

If IsArray(rs) Then
	' Div명 출력
	For i=0 To UBound(rs,2)
		If i > 0 And rs(0,0) <> rs(0,i) Then
%>
    </tr>
	<tr>
<%
			trEndStart = true
			Exit For
		End If
		response.write "<td align='center'>" & rs(6,i) & "</td>"
	Next

	If Not trEndStart Then
%>
    </tr>
	<tr>
<%
	End If

	' 카운트 출력
	For i=0 To UBound(rs,2)
		If prevRow <> rs(0,i) Then	' 첫번째 로우가 바뀔 때
%>
    	<td><%=sumCnt%></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
<!--     	<td><%'=rs(0,i)%> <%'If rs(4,i) <> "" Then response.write "(" & Replace(rs(4,i),"\0x2F","/") & ")" End If %></td> -->
    	<td><%If divCd = "" Then response.write Replace(rs(4,i),"\0x2F","/") Else response.write rs(0,i)End If %></td>
<%
			prevRow = rs(0,i)
			sumCnt	= 0
		End If

		If divCd <> "" Then
			response.write "<td align='center'><a href=""javascript:popCsListView('"&divCd&"','"&rs(0,i)&"','"&rs(1,i)&"','"&rs(2,i)&"');"">" & rs(3,i) & "</a></td>"
		Else
			response.write "<td align='center'><a href=""javascript:popCsListView('"&rs(0,i)&"','','"&rs(1,i)&"','"&rs(2,i)&"');"">" & rs(3,i) & "</a></td>"
		End If

		sumCnt	  = sumCnt + CDbl(rs(3,i))
		sumTotCnt = sumTotCnt + CDbl(rs(3,i))

		divNow = i Mod divLen
		totCnt(divNow) = totCnt(divNow) + CDbl(rs(3,i))

	Next
%>
    	<td><%=sumCnt%></td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><b>합계</b></td>
    	<%For i=0 To divLen - 1 %>
    	<td align="center"><a href="javascript:popCsListView('<%=divCd%>','','<%=rs(1,i)%>','<%=rs(2,i)%>');"><b><%=totCnt(i)%></b></a></td>
		<%Next %>
		<td><b><%=sumTotCnt%></b></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><b>%</b></td>
    	<%For i=0 To divLen - 1 %>
    	<td align="center">
			<%
			If sumTotCnt > 0 Then
				response.write CInt(totCnt(i) * 100 / sumTotCnt)
			Else
				response.write "0"
			End If
			%>%
		</td>
		<%Next %>
		<td>100%</td>
    </tr>
<%
Else
	response.write "<td>검색결과가 없습니다.</td>" & vbCrLf
	response.write "</tr>" & vbCrLf
End If
%>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
