<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월별 면세 매출
' History : 2011.06.02 eastone 생성
'			2012.07.11 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/noTaxSummary.asp"-->
<%
Dim page, research ,yyyy1, yyyy2, mm1, mm2 ,makerid ,placeALL ,placegubun,placesub ,olist, i ,TTLcnt, TTLsum
dim pyyyymm, ColCnt
	page  = requestCheckvar(request("page"),10)
	research = requestCheckvar(request("research"),10)
	makerid = requestCheckvar(request("makerid"),32)
	placeALL = requestCheckvar(request("placeALL"),30)
	yyyy1 = requestCheckvar(request("yyyy1"),10)
	yyyy2 = requestCheckvar(request("yyyy2"),10)
	mm1   = requestCheckvar(request("mm1"),10)
	mm2   = requestCheckvar(request("mm2"),10)

if (yyyy1="") then
    yyyy1 = Left(CStr(DateAdd("m",-1,now())),4)
    yyyy2 = yyyy1
    mm1 = Mid(CStr(DateAdd("m",-1,now())),6,2)
    mm2 = mm1
end if

placegubun = Left(placeALL,5)
placesub = Mid(placeALL,6,255)

set olist = new CNoTaxList
	olist.FRectMakerid = makerid
	olist.FRectStYYYYMM = yyyy1+"-"+mm1
	olist.FRectEdYYYYMM = yyyy2+"-"+mm2
	olist.FRectplaceGubun = placegubun
	olist.FRectplaceSub   = placesub
	olist.getNoTaxListMonth

%>

<script language='javascript'>

function rsearch(pg){
    frm.page.value=pg;
    frm.submit();
}

function popDetail(yyyymm1,yyyymm2,pgn,ps,makerid){
    var rUrI = '/admin/datamart/mng/popmonthNoTaxDetail.asp?yyyymm1=' + yyyymm1 + '&yyyymm2=' + yyyymm2 + '&pgn='+pgn+'&ps='+ps+'&makerid='+makerid;
    
    var popwin = window.open(rUrI,'popNoTaxDetail','scrollbars=yes,resizable=yes,width=1024,height=768');
    popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간 :
		<% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="rsearch(1);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    구분     : <% call drawBoxNotaxPlaceGubun("placeALL",placeALL) %>
	    
		브랜드 ID : <% call drawSelectBoxDesigner("makerid",makerid) %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= olist.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= CHKIIF(olist.FTotalCount<1,0,page) %>/ <%= olist.FTotalPage %></b>
	</td>
</tr>

<%
ColCnt = -1
%>
<% if olist.FTotalCount>0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% for i=0 to olist.FResultCount-1 %>
    <% 
    if (pyyyymm<>olist.FItemList(i).Fyyyymm) and (i<>0) Then
        'response.write "<td>합계</td><td>상세</td>"
        'response.write "</tr>"
        Exit For
    Else
        if (i=0) then response.write "<td>구분</td>"
        response.write "<td>" & olist.FItemList(i).FplaceSubName& "</td>"
        pyyyymm = olist.FItemList(i).Fyyyymm
    end if
    %>
<% next %>
<% ColCnt = i %>
<td>합계</td><td>상세</td>
</tr>

<%
pyyyymm =""
Dim colSum(50)
Dim rowSum : rowSum = 0
Dim totSum : totSum = 0
Dim ColPos : ColPos = 0
%>
<tr align="center" bgcolor="#FFFFFF" >
	<% for i=0 to olist.FResultCount-1 %>
	    <% 
	    ''rw "ColCnt="&ColCnt
	    if (pyyyymm<>olist.FItemList(i).Fyyyymm) and (i<>0) Then
	        response.write "<td>"&Formatnumber(rowSum,0) &"</td><td><input type='button' onclick=""popDetail('"& pyyyymm & "','','" & placeGubun & "','"& placeSub &"','"& makerid &"');"" value='상세' class='button'></td></tr>"
	        totSum = totSum + rowSum
	        rowSum = 0
	        if (i<>olist.FResultCount-1) then 
	            response.write  "<tr align='center' bgcolor='#FFFFFF' >"
	            response.write "<td>"& olist.FItemList(i).FYYYYMM&"</td>"
	        end if
	    Else
	        if (i=0) then 
	            response.write "<td>"& olist.FItemList(i).FYYYYMM&"</td>"
	        end if
	        
	    end if
	    response.write "<td>" & FormatNumber(olist.FItemList(i).FnotaxPrice,0) & "</td>"
	    rowSum = rowSum + olist.FItemList(i).FnotaxPrice
	    ColPos = i mod ColCnt
	    colSum(ColPos) = colSum(ColPos) + olist.FItemList(i).FnotaxPrice
	    pyyyymm = olist.FItemList(i).Fyyyymm
	    %>
	<% next %>
	<% totSum = totSum + rowSum %>
	<td><%= FormatNumber(rowSum,0) %></td>
	<td>
		<input type='button' onclick="popDetail('<%= pyyyymm %>','','<%=  placeGubun%>','<%=  placeSub%>','<%= makerid %>');" value='상세' class='button'>
	</td>
</tr>
<tr align="center" bgcolor="#EEEEFF" >
    <td>합계</td>
    <% for i=0 to ColCnt-1 %>
    <td><%= FormatNumber(colSum(i),0) %></td>
    <% next %>
    <td><%= FormatNumber(totSum,0) %></td>
    <td>
    	<input type='button' onclick="popDetail('<%= yyyy1+"-"+mm1 %>','<%= yyyy2+"-"+mm2 %>','<%=  placeGubun%>','<%=  placeSub%>','<%= makerid %>');" value='상세' class='button'>
    </td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
    <td height="40" align="center">[ 검색 결과가 없습니다.] </td>
</tr>
<% end if %>
</table>

<%
set olist = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
