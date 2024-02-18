<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  주문서관리
' History : 		   이상구 생성
'			2016.08.17 한용민 수정
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim idx,itype
idx = requestCheckVar(request("idx"),20)
itype = requestCheckVar(request("itype"),50)


dim oordersheetmaster, oordersheet
set oordersheetmaster = new COrderSheet
oordersheetmaster.FRectIdx = idx
oordersheetmaster.GetOneOrderSheetMaster

dim isFixed
isFixed = oordersheetmaster.FOneItem.IsFixed


set oordersheet = new COrderSheet
oordersheet.FrectisFixed = isFixed
oordersheet.FRectIdx = idx
oordersheet.GetOrderSheetDetail


dim obrand
set obrand = new CBrandShopInfoItem

obrand.FRectChargeId = oordersheetmaster.FOneItem.Ftargetid
obrand.GetBrandShopInFo


dim i

dim scheduleorexedate
if not IsNULL(oordersheetmaster.FOneItem.FScheduleDate) then
scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FScheduleDate),10),"-","/")
end if

dim ttlsellcash, ttlbuycash, ttlcount
ttlsellcash = 0
ttlbuycash  = 0
ttlcount    = 0

function getObjStr(v)
	dim reStr
	reStr = "<OBJECT" + vbCrlf
	reStr = reStr + "id=iaxobject" + vbCrlf
	reStr = reStr + "classid='clsid:A4F3A486-2537-478C-B023-F8CCC41BF29D'" + vbCrlf
	reStr = reStr + "codebase='http://partner.10x10.co.kr/cab/tenbarShow.cab#version=1,0,0,3'" + vbCrlf
	reStr = reStr + "width=100" + vbCrlf
	reStr = reStr + "height=20" + vbCrlf
	reStr = reStr + "align=bottom" + vbCrlf
	reStr = reStr + "hspace=0" + vbCrlf
	reStr = reStr + "vspace=0" + vbCrlf
	reStr = reStr + ">" + vbCrlf
	reStr = reStr + "</OBJECT>" + vbCrlf

	getObjStr = reStr
end function

%>

<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + oordersheetmaster.FOneItem.Ftargetid + Left(CStr(now()),10) + ".xls"
end if
%>



<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td colspan="3">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
			    	<td style="font-size:12pt; font-family:돋움, arial;"><b>거래명세서(<%= oordersheetmaster.FOneItem.Ftargetid %>)</b></td>
					<td align="right">
			    		<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=23&data=<%= oordersheetmaster.FOneItem.FBaljuCode %>&height=20&barwidth=1&TextAlign=2" <%=CHKIIF(LCASE(session("ssBctId"))="smlgroup","onClick='this.remove()'","") %>>
			    		&nbsp;&nbsp;&nbsp;&nbsp;
			    		<b>주문코드 (<%= oordersheetmaster.FOneItem.FBaljuCode %>)</b>
			<!--       	&nbsp;<%= getObjStr("oordersheetmaster.FOneItem.FBaljuCode") %>	-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="1">
		<td colspan="3" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<tr valign="top" style="padding:10 0 0 0">
        <td width="49%">
        	<!-- 공급자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4"><b>공급자 정보</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>등록번호</td>
        			<td colspan="3"><%= obrand.FSocNo %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="60">상호</td>
        			<td width="135"><b><%= obrand.FChargeName %></b></td>
        			<td width="60">대표자</td>
        			<td width="90"><%= obrand.FCeoName %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>소재지</td>
        			<td colspan="3"><%= obrand.FAddress %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>업태</td>
        			<td><%= obrand.FUptae %></td>
        			<td>업종</td>
        			<td><%= obrand.FUpjong %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>담당자</td>
        			<td><%= obrand.FManagerName %></td>
        			<td>연락처</td>
        			<td><%= obrand.FManagerHp %></td>
        		</tr>
        	</table>
        	<!-- 공급자정보 끝 -->
        </td>
        <td>&nbsp;</td>
        <td width="49%">
        	<!-- 공급받는자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4"><b>공급받는자 정보</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>등록번호</td>
        			<td colspan="3">211-87-00620</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="60">상호</td>
        			<td width="135">(주)텐바이텐</td>
        			<td width="60">대표자</td>
        			<td width="90">최은희</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>소재지</td>
        			<td colspan="3">서울시 종로구 동숭동 1-45 자유빌딩 2층</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>업태</td>
        			<td>서비스,도소매 등</td>
        			<td>업종</td>
        			<td>전자상거래 등</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>주문자</td>
        			<td><%= oordersheetmaster.FOneItem.Fregname %></td>
        			<td>연락처</td>
        			<td>1644-1851</td>
        		</tr>
        	</table>
        	<!-- 공급받는자정보 끝 -->
        </td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="15">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<strong>상세내역</strong>
			        	<b>(총액 : \<%= ForMatNumber(oordersheetmaster.FOneItem.FTotalBuycash,0) %>)</b>
			        </td>
			       	<td align="right">
			       		<b>주문일자 : <%= scheduleorexedate %></b>
			    	</td>
			    </tr>
			</table>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="90">상품코드</td>
    	<td>상품명</td>
    	<td>옵션명</td>
    	<td width="55">소비자가</td>
    	<td width="55">공급가</td>
    	<td width="50">수량</td>
    	<td width="70">공급가합계</td>
    </tr>

	<% for i=0 to oordersheet.FResultCount -1 %>
	<%
		ttlsellcash = ttlsellcash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FSellcash
		ttlbuycash = ttlbuycash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FBuycash
		ttlcount = ttlcount + oordersheet.FItemList(i).Frealitemno
	%>

	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oordersheet.FItemList(i).FItemGubun %>-<b><%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %></b>-<%= oordersheet.FItemList(i).FItemOption %></td>
		<td align="left"><%= left(oordersheet.FItemList(i).FItemName,35) %></td>
		<td><%= left(oordersheet.FItemList(i).FItemOptionName,15) %></td>
		<td align="right"><%= FormatNumber(oordersheet.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= FormatNumber(oordersheet.FItemList(i).FBuycash,0) %></td>
		<td><%= oordersheet.FItemList(i).FRealItemno %></td>
		<td align="right"><%= FormatNumber(oordersheet.FItemList(i).FRealItemno*oordersheet.FItemList(i).FBuycash,0) %></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="#FFFFFF">비고</td>
		<td colspan="3" align="left" bgcolor="#FFFFFF"><%= nl2br(oordersheetmaster.FoneItem.FComment) %></td>
		<td><b>총계</b></td>
		<td><b><%= ttlcount %></b></td>
		<td align="right"><b><%= ForMatNumber(ttlbuycash,0) %></b></td>
	</tr>
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="15">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td width="50%" align="left"><b>인계자 :</b></td>
			       	<td><b>인수자 :</b></td>
			    </tr>
			</table>
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			* 텐바이텐 물류센터 주소 : <b>[11154] 경기도 포천시 군내면 용정경제로2길 83 텐바이텐 물류센터</b> (연락처 : 1644-1851)</b>
		</td>
	</tr>
</table>


<script language='javascript'>
//iaxobject.ShowBarCode(30,'<%= oordersheetmaster.FOneItem.FBaljuCode %>',2);

function getOnLoad(){
   window.print();
}
window.onload=getOnLoad;
</script>

<%
set obrand = Nothing
set oordersheetmaster = Nothing
set oordersheet = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
