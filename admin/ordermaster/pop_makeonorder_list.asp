<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주문제작(텐배) 출고지시리스트
' History : 2007년 11월 29일 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	''date1 = dateAdd("d",-4,nowdate)
	date1 = nowdate
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(dateserial(yyyy2, mm2 , dd2)+1),10)
end if

dim objbaljumakeonorder, balju_code

balju_code	= requestCheckVar(request("balju_code"),10)

set objbaljumakeonorder = New COldMiSend
objbaljumakeonorder.FPageSize = 500

if (balju_code <> "") then
	objbaljumakeonorder.FRectBaljuCode = balju_code
else
	objbaljumakeonorder.FRectStartDate = yyyy1 + "-" + mm1 + "-" + dd1
	objbaljumakeonorder.FRectEndDate = Left(CStr(Edate),10)
end if

objbaljumakeonorder.GetBaljuListMakeOnOrder

dim i, tmp
dim orgitemno, makeonorderitemno
%>
<script language='javascript'>

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cOrderFin(detailidx){
    if (confirm('취소 처리 확인 하시겠습니까?')){
        var popwin = window.open("/admin/ordermaster/misendmaster_main_process.asp?detailidx=" + detailidx + "&mode=cancelFin","misendmaster_process","width=100 height=100 scrollbars=yes resizable=yes");
	    popwin.focus();
    }
}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			출고지시코드 :
			<input type="text" class="text" name="balju_code" value="<%= balju_code %>" size="10" maxlength="12">
			&nbsp;
			조회기간(출고지시일자) : <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> (출고지시코드 없는 경우)
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>

<p>

* 최대 500개까지 표시됩니다.

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%= objbaljumakeonorder.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	    <td width="70">주문번호</td>
        <td width="70">Site</td>
	    <td width="60">주문자</td>
	    <td width="60">수령인</td>
		<td width="100">상품구분</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="40">주문<br>수량</td>
	    <td>주문제작<br>문구</td>
		<td width="100">비고</td>
	</tr>
	<% if objbaljumakeonorder.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF" height="25">
	  	<td colspan="16" align="center">검색결과가 없습니다.</td>
	</tr>
	<% else %>
	<% for i=0 to objbaljumakeonorder.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
	    <td align="center">
	    <%
	    if (tmp <> objbaljumakeonorder.FItemList(i).FOrderSerial) then
	      tmp = objbaljumakeonorder.FItemList(i).FOrderSerial

		  orgitemno = 0
		  makeonorderitemno = 0
	    %>
			<a href="javascript:misendmaster('<%= objbaljumakeonorder.FItemList(i).FOrderSerial %>');"><%= objbaljumakeonorder.FItemList(i).FOrderSerial %></a>
	    <% end if %>
		<%
		if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderOrgItem) then
			orgitemno = orgitemno + objbaljumakeonorder.FItemList(i).FItemNo
		elseif (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then
			makeonorderitemno = makeonorderitemno + objbaljumakeonorder.FItemList(i).FItemNo
		end if
		%>
	    </td>
        <td><%= objbaljumakeonorder.FItemList(i).FSiteName %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FBuyName %></td>
    	<td><%= objbaljumakeonorder.FItemList(i).FReqName %></td>
	    <td align="left">
			<% if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderOrgItem) then %>
				<font color="blue">원상품</font>
			<% elseif (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then %>
				&nbsp; -&gt; <font color="green">주문제작</font>
			<% end if %>
		</td>
		<td><%= objbaljumakeonorder.FItemList(i).FItemId %></td>
		<td align="left">
			<%= objbaljumakeonorder.FItemList(i).FItemname %>
			<% if objbaljumakeonorder.FItemList(i).FItemOptionName<>"" then %>
			<font color="blue">[<%= objbaljumakeonorder.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= objbaljumakeonorder.FItemList(i).FItemNo %></td>
		<td>
			<% if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then %>
				<%= objbaljumakeonorder.FItemList(i).Frequiredetail %>
			<% end if %>
		</td>
	    <td>
			<% if (objbaljumakeonorder.FItemList(i).FisMakeOnOrderItem) then %>
				<% if (orgitemno = 0) then %>
					<font color="red"><b>원상품 없음</b></font>
				<% elseif (orgitemno > 1) or (orgitemno <> makeonorderitemno) then %>
					<font color="red">매칭 필요</font>
				<% end if %>
			<% end if %>
		</td>
	</tr>
  <% next %>
  <% end if %>
</table>


<%
set objbaljumakeonorder = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
