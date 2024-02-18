<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%

dim oldmisend,  inputyn
inputyn = request("inputyn")
if inputyn="" then inputyn="N"

set oldmisend = New COldMiSend
oldmisend.FPageSize = 500
oldmisend.FRectDelayDate = 0
'oldmisend.FRectNotInCludeUpcheCheck = notincludeupchecheck
oldmisend.FRectInCludeAlreadyInputed = inputyn
oldmisend.GetOldMisendListMasterCS


dim i, tmp
%>
<script language='javascript'>
</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	<input type="radio" name="inputyn" value="Y" <% if (inputyn = "Y") then response.write "checked" end if %>> 전체목록
				<input type="radio" name="inputyn" value="N" <% if (inputyn = "N") then response.write "checked" end if %>> 미처리목록
				<input type="radio" name="inputyn" value="1" <% if (inputyn = "1") then response.write "checked" end if %>> SMS완료
				<input type="radio" name="inputyn" value="2" <% if (inputyn = "2") then response.write "checked" end if %>> 안내Mail완료
				<input type="radio" name="inputyn" value="3" <% if (inputyn = "3") then response.write "checked" end if %>> 통화완료
				<input type="radio" name="inputyn" value="6" <% if (inputyn = "6") then response.write "checked" end if %>> CS완료
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="frmview" method="get">
  <input type="hidden" name="iid" value="">

  </form>
  <tr bgcolor="#FFFFFF">
  	<td colspan="17" align="left">총건수 : <%= oldmisend.FResultCount %></td>
  </tr>
  <tr bgcolor="DDDDFF" align="center">
    <td width="70" align="center">주문번호</td>
    <td width="70" align="center">주문일 /<br>결제일</td>
    <td width="60" align="center">사이트명</td>
    <td width="80" align="center">아이디</td>
    <td width="60" align="center">구매자 /<br>수령인</td>
    <td width="90" align="center">구매자전화 /<br>구매자핸드폰</td>
    <td width="50" align="center">결제금액</td>
    <td width="70" align="center">거래상태 /<br>송장No</td>
    <td width="44" align="center">상품</td>
    <td width="80" align="center">배송지연<br>사유</td>
    <td width="80" align="center">요청사항</td>
    <td width="80" align="center">처리결과</td>
    <td width="80" align="center">처리구분</td>
  </tr>
  <% if oldmisend.FResultCount<1 then %>
  <tr bgcolor="#FFFFFF">
  	<td colspan="17" align="center">검색결과가 없습니다.</td>
  </tr>
  <% else %>

  <% for i=0 to oldmisend.FResultCount -1 %>
  <tr bgcolor="#FFFFFF">
    <td align="center">
    <%
    if (tmp <> oldmisend.FItemList(i).FOrderSerial) then
      tmp = oldmisend.FItemList(i).FOrderSerial
    %>
      <a href="misendmaster_main.asp?orderserial=<%= oldmisend.FItemList(i).FOrderSerial %>" target="mainFrame"><%= oldmisend.FItemList(i).FOrderserial %></a>
    <% end if %>
    </td>
    <td align="center"><%= Left(oldmisend.FItemList(i).FRegdate,10) %><br><%= Left(oldmisend.FItemList(i).FIpkumDate,10) %></td>
    <td align="center"><%= oldmisend.FItemList(i).FSiteName %></td>
    <td align="center"><%= oldmisend.FItemList(i).FUserID %></td>
    <td align="center"><%= oldmisend.FItemList(i).FBuyName %><br><%= oldmisend.FItemList(i).FReqName %></td>
    <td align="center"><%= oldmisend.FItemList(i).FBuyPhone %><br><%= oldmisend.FItemList(i).FBuyHP %></td>
    <td align="right"><%= FormatNumber(oldmisend.FItemList(i).FSubTotalPrice,0) %></td>
    <td align="center"><font color="<%= oldmisend.FItemList(i).IpkumDivColor %>"><%= oldmisend.FItemList(i).IpkumDivName %></font>
    <br><%= oldmisend.FItemList(i).FDeliveryNo %>
    </td>
    <td align="center"><%= oldmisend.FItemList(i).FItemId %></td>
    <td align="center">
	<%= oldmisend.FItemList(i).getMiSendCodeName %><br><%= oldmisend.FItemList(i).getIpgoMayDay %>
    </td>
    <td><%= oldmisend.FItemList(i).FrequestString %></td>
    <td><%= oldmisend.FItemList(i).FfinishString %></td>
    <td align="center"><%= oldmisend.FItemList(i).GetStateString %></td>
  </tr>
  <% next %>
  <% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oldmisend = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->








