<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 입출고 내역서
' History : 2009.04.07 서동석 생성
'			2012.07.17 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%

dim idx
	idx = requestCheckVar(request("idx"),10)

dim oipchulmaster, oipchul
set oipchulmaster = new CShopIpChul
oipchulmaster.FRectIdx = idx
oipchulmaster.GetIpChulMasterList

set oipchul = new CShopIpChul
oipchul.FRectIdx = idx
oipchul.GetIpChulDetail

dim i

dim yyyymmdd,yyyy1,mm1,dd1
yyyymmdd = Left(CStr(oipchulmaster.FItemList(0).FScheduleDt),10)
yyyy1 = left(yyyymmdd,4)
mm1 = mid(yyyymmdd,6,2)
dd1 = mid(yyyymmdd,9,2)
%>

<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
  <tr valign="bottom">
    <td width="10" height="10" align="right" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif" bgcolor="#F3F3FF"></td>
    <td width="10" height="10" align="left" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
  </tr>
  <tr valign="top">
    <td height="20" background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td height="20" background="/images/tbl_blue_round_06.gif" bgcolor="#F3F3FF"><img src="/images/icon_star.gif" align="absbottom">
    <font color="red"><strong>매장 개별 입출고 내역</strong></font></td>
    <td height="20" background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td align="right" bgcolor="#F3F3FF">
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td bgcolor="#F3F3FF">
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>

  <tr valign="top" bgcolor="#F3F3FF">
    <td height="10" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td height="10" background="/images/tbl_blue_round_08.gif"></td>
    <td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
  </tr>
</table>
<p>
<table width="100%" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
<tr>
	<td width="100" bgcolor="#DDDDFF">공급처</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
		<%= oipchulmaster.FItemList(0).FChargeid %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">매장 </td>
	<td bgcolor="#FFFFFF">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
		<%= oipchulmaster.FItemList(0).FShopid %> (<%= oipchulmaster.FItemList(0).FShopname %>)
	</td>
</tr>
	<input type="hidden" name="divcode" value="006">
	<input type="hidden" name="vatcode" value="008">
<tr>
	<td width="100" bgcolor="#DDDDFF">총판매가</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
</tr>
<% if Not (C_IS_SHOP) then %>
	<tr>
		<td width="100" bgcolor="#DDDDFF">총매입가</td>
		<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSuplyCash,0) %></td>
	</tr>
	<% if Not (C_IS_Maker_Upche)  then %>
	<tr>
		<td width="100" bgcolor="#DDDDFF">총공급가</td>
		<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalShopBuyPrice,0) %></td>
	</tr>
	<% end if %>
<% end if %>
<tr>
	<td width="100" bgcolor="#DDDDFF">입고예정일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FScheduleDt %>
	택배사 : <%= oipchulmaster.FItemList(0).Fsongjangname %>
	송장번호:<%= oipchulmaster.FItemList(0).Fsongjangno %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">입고일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FexecDt %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">매장확인일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).Fshopconfirmdate %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">업체확인일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).Fupcheconfirmdate %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">등록일</td>
	<td bgcolor="#FFFFFF"><%= oipchulmaster.FItemList(0).FRegDate %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">상태</td>
	<td bgcolor="#FFFFFF"><font color="<%= oipchulmaster.FItemList(0).getStateColor %>"><%= oipchulmaster.FItemList(0).getStateName %></font></td>
</tr>
</table>
<br>
<br>
<table width="100%" cellspacing="0" class="a" >
<tr>
	<td align="right"></td>
</tr>
</table>
<table width="100%" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
	<% if oipchul.FresultCount>0 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="right">총건수: <%= oipchul.FTotalCount %> &nbsp;</td>
	</tr>
	<% end if %>
	<tr bgcolor="#DDDDFF" align="center">
		<td width="80">바코드</td>
		<td width="80">브랜드ID</td>
		<td width="100">상품명</td>
		<td width="100">옵션명</td>
		<td width="50">판매가</td>
		<% if Not (C_IS_SHOP) then %>
			<td width="50">텐바이텐<br>공급가</td>
			<% if Not (C_IS_Maker_Upche)  then %>
			<td width="50">매장<br>공급가</td>
			<% end if %>
		<% end if %>
		<td width="50">갯수</td>
		<td width="60">판매가합계</td>
	</tr>
	<% for i=0 to oipchul.FResultCount-1 %>

	<tr bgcolor="#FFFFFF">
		<td><%= oipchul.FItemList(i).GetBarCode %></td>
		<td><%= oipchul.FItemList(i).Fdesignerid %></td>
		<td><%= oipchul.FItemList(i).FItemName %></td>
		<td><%= oipchul.FItemList(i).FItemOptionName %></td>
		<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSellCash,0) %></td>
		<% if Not (C_IS_SHOP) then %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSuplyCash,0) %></td>
			<% if Not (C_IS_Maker_Upche)  then %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).FShopbuyprice,0) %></td>
			<% end if %>
		<% end if %>
		<td align="center"><%= oipchul.FItemList(i).Fitemno %></td>
		<td align="right"><%= ForMatNumber(oipchul.FItemList(i).Fitemno*oipchul.FItemList(i).FSellCash,0) %></td>
	</tr>
	</form>
	<% next %>
</table>

<%
set oipchulmaster = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->