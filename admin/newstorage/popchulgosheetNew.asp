<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim idx,itype
idx = request("idx")
itype = request("itype")

'==============================================================================
dim oipchul, oipchuldetail
set oipchul = new CIpChulStorage
oipchul.FRectId = idx
oipchul.GetIpChulMaster

set oipchuldetail = new CIpChulStorage
oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
oipchuldetail.GetIpChulDetail

'==============================================================================
dim obrand
set obrand = new CBrandShopInfoItem

obrand.FRectChargeId = oipchul.FOneItem.Fsocid
obrand.GetBrandShopInFo



dim i

dim executedate

if (oipchul.FOneItem.Fexecutedt <> "") then
	executedate = replace(Left(CstR(oipchul.FOneItem.Fexecutedt),10),"-","/")
else
	executedate = "<font color='red'>미출고</font>"
end if

dim ttlsellcash, ttlsuplycash, ttlcount
ttlsellcash = 0
ttlsuplycash  = 0
ttlcount    = 0
%>
<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + oipchul.FOneItem.Fsocid + Left(CStr(now()),10) + ".xls"
end if
%>






<!-- 표 상단바 시작-->

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="20">
		<td align="left">
			<font size="3"><b>출고내역서(<%= obrand.FChargeName %>)</b></font>
		</td>
		<td align="right">
			<b>텐바이텐 (<%= oipchul.FOneItem.Fcode %>)</b>
		</td>
	</tr>
	<tr height="1" bgcolor="<%= adminColor("tablebg") %>">
		<td colspan="2"></td>
	</tr>
</table>

<p>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top">
        <td width="48%">
        	<!-- 공급자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4"><b>공급자 정보</b></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>등록번호</td>
        			<td colspan="3">211-87-00620</td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td width="60">상호</td>
        			<td width="135">(주)텐바이텐</td>
        			<td width="60">대표자</td>
        			<td width="90"><%= TENBYTEN_CEONAME %></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>소재지</td>
        			<td colspan="3">(03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐</td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>업태</td>
        			<td>서비스,도소매 등</td>
        			<td>업종</td>
        			<td>전자상거래 등</td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>&nbsp;</td>
        			<td></td>
        			<td></td>
        			<td></td>
        		</tr>
        	</table>
        	<!-- 공급자정보 끝 -->
        </td>
        <td bgcolor="#FFFFFF">&nbsp;</td>
        <td width="48%">
        	<!-- 공급받는자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4"><b>공급받는자 정보</b></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>등록번호</td>
        			<td colspan="3"><%= obrand.FSocNo %></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td width="60">상호</td>
        			<td width="135"><b><%= obrand.FSocName %></b></td>
        			<td width="60">대표자</td>
        			<td width="90"><%= obrand.FCeoName %></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>소재지</td>
        			<td colspan="3"><%= obrand.FAddress %></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>업태</td>
        			<td><%= obrand.FUptae %></td>
        			<td>업종</td>
        			<td><%= obrand.FUpjong %></td>
        		</tr>
        		<tr align="center" height="23" bgcolor="#FFFFFF">
        			<td>담당자</td>
        			<td><%= obrand.FManagerName %></td>
        			<td>연락처</td>
        			<td><%= obrand.FManagerHp %></td>
        		</tr>
        	</table>
        	<!-- 공급받는자정보 끝 -->
        </td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="8">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td><img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<strong>출고상세내역</strong></td>
					<td align="right"><b>출고일자 : <%= executedate %></b></td>
				</tr>
			</table>
		</td>
	</tr>
    <tr align="center" height="23" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="120">상품코드</td>
        <td width="80">랙코드</td>
    	<td>상품명</td>
    	<td>옵션명</td>
    	<td width="60">소비자가</td>
    	<td width="60">공급가</td>
    	<td width="50">수량</td>
    	<td width="70">공급가합계</td>
    </tr>


	 <% for i=0 to oipchuldetail.FResultCount -1 %>
	 <%
	 	ttlsellcash = ttlsellcash + oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsellcash
	 	ttlsuplycash = ttlsuplycash + oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash
	 	ttlcount = ttlcount + oipchuldetail.FItemList(i).Fitemno
	 %>

	<tr height="23" align="center" bgcolor="#FFFFFF">
		<td><%= oipchuldetail.FItemList(i).Fiitemgubun %>-<b><%= CHKIIF(oipchuldetail.FItemList(i).FItemId>=1000000,Format00(8,oipchuldetail.FItemList(i).FItemId),Format00(6,oipchuldetail.FItemList(i).FItemId)) %></b>-<%= oipchuldetail.FItemList(i).FItemOption %>
		</td>
        <td align="center"><%= oipchuldetail.FItemList(i).FrackcodeByOption %></td>
		<td align="left"><%= oipchuldetail.FItemList(i).FIItemName %></td>
		<td><%= oipchuldetail.FItemList(i).FIItemoptionName %></td>
		<td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).Fsuplycash,0) %></td>
		<td><%= oipchuldetail.FItemList(i).Fitemno %></td>
		<td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash,0) %></td>
	<% next %>
	<tr height="23" align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="#FFFFFF">비고</td>
		<td colspan="4" align="left" bgcolor="#FFFFFF"><%= nl2br(oipchul.FOneItem.Fcomment) %></td>
		<td><b>총계</b></td>
		<td><b><%= ttlcount %></b></td>
		<td align="right"><b><%= ForMatNumber(ttlsuplycash,0) %></b></td>
	</tr>
</table>













<%
set obrand = Nothing
set oipchul = Nothing
set oipchuldetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
