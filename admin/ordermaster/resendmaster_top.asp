<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/resendOrderCls.asp"-->

<%

dim oResend,  isCancel, itemid, itemoption, vSiteName, designer
itemid  = RequestCheckVar(request("itemid"),10)
itemoption  = RequestCheckVar(request("itemoption"),10)
isCancel = request("isCancel")
vSiteName		= requestCheckVar(request("sitename"),10)
designer		= requestCheckVar(request("designer"),32)

if isCancel="" then isCancel="A"

set oResend = New CReSend
oResend.FPageSize = 500
oResend.FRectIsCancel = isCancel
oResend.FRectSiteName = vSiteName

oResend.FRectMakerid = designer
oResend.FRectItemID = itemid
oResend.FRectItemOption = itemoption
oResend.GetResendOrderList

dim i, tmp

%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesigner "designer", designer %>
			&nbsp;
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="8" maxlength="10">
            &nbsp;
			상품코드 :
			<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="8" maxlength="10">
            &nbsp;
			Site :
			<select name="sitename" class="select">
				<option value="">-전체-</option>
				<option value="10x10" <%=CHKIIF(vSiteName="10x10","selected","")%>>텐바이텐</option>
				<option value="NOTTEN" <%=CHKIIF(vSiteName="NOTTEN","selected","")%>>제휴사</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="radio" name="isCancel" value="A" <% if (isCancel = "A") then response.write "checked" end if %>> 전체목록
			<input type="radio" name="isCancel" value="C" <% if (isCancel = "C") then response.write "checked" end if %>> 취소주문
		</td>
	</tr>
	</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="10">
			검색결과 : <b><%= oResend.FResultCount %></b> / 주문건수 : <b><%= oResend.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="70">주문번호</td>
        <td width="70">Site</td>
	    <td width="60">주문자</td>
	    <td width="60">수령인</td>
		<td width="50">상품코드</td>
		<td width="100">브랜드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="40">주문<br>수량</td>
		<td width="40">주문일</td>
		<td width="40">등록일</td>
	</tr>
	<% if oResend.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
	  	<td colspan="10" align="center">검색결과가 없습니다.</td>
	</tr>
	<% else %>

	<% for i=0 to oResend.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td align="center">
		<%
			if (tmp <> oResend.FItemList(i).FOrderSerial) then
				tmp = oResend.FItemList(i).FOrderSerial
		%>
			<%= oResend.FItemList(i).FOrderSerial %>
	    <% end if %>
	    </td>
        <td><%= oResend.FItemList(i).FSiteName %></td>
		<td><%= oResend.FItemList(i).FBuyName %></td>
    	<td><%= oResend.FItemList(i).FReqName %></td>
	    <td><%= oResend.FItemList(i).FItemId %></td>
		<td><%= oResend.FItemList(i).Fmakerid %></td>
		<td align="left">
			<%= oResend.FItemList(i).FItemname %>
			<% if oResend.FItemList(i).FItemOptionName<>"" then %>
			<font color="blue">[<%= oResend.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= oResend.FItemList(i).FItemNo %></td>
		<td><%= left(oResend.FItemList(i).FRegDate,10) %></td>
		<td><%= left(oResend.FItemList(i).FRegDate,10) %></td>
	</tr>
  <% next %>
  <% end if %>
</table>


<%
set oResend = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
