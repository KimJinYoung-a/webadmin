<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 리스트
' History : 2010.12.09 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/sale_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->

<%
Dim sCode, clsSaleItem ,iTotCnt, arrList,i , shopid ,mSPrice, mSBPrice, iSaleMargin, iOrgMargin ,sellpricemargin
Dim sTitle,isRate, isMargin, isStatus,eCode, egCode, dSDay, dEDay, isUsing, dOpenDay,isMValue ,adminvspos
Dim ix,page , sSearchTxt , iSerachType , sBrand ,designer ,itemid , itemname , smargin ,sshopmargin
	adminvspos = requestCheckVar(Request("adminvspos"),2)
	sCode = requestCheckVar(Request("sC"),10)
	page = requestCheckVar(request("page"),10)
	iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	isStatus		= requestCheckVar(Request("salestatus"),4)	'할인 상태
	shopid		= requestCheckVar(Request("shopid"),32)		'매장
	designer    = RequestCheckVar(request("designer"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)

if page = "" then page = 1

'검색부분이 번호만 받아야된다면 숫자만 접수
if iSerachType="1" or iSerachType="2" then
	sSearchTxt = getNumeric(sSearchTxt)
end if

iSaleMargin=0
iOrgMargin = 0

'할인 상품정보
set clsSaleItem = new CSaleItem
	clsSaleItem.FPageSize = 20
	clsSaleItem.FCurrPage = page
	clsSaleItem.FSearchType = iSerachType
	clsSaleItem.FSearchTxt  = sSearchTxt
	clsSaleItem.FBrand		= sBrand
	clsSaleItem.FSStatus	= isStatus
	clsSaleItem.frectshopid = 	shopid
	clsSaleItem.FRectDesigner = designer
	clsSaleItem.FRectItemId = 	itemid
	clsSaleItem.FRectItemName = 	itemname
	clsSaleItem.frectadminvspos = adminvspos
	clsSaleItem.fnGetSaleItemList()

'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
Dim arrsalemargin, arrsalestatus
	arrsalemargin = fnSetCommonCodeArr_off("salemargin",False)
	arrsalestatus= fnSetCommonCodeArr_off("salestatus",False)
%>

<script language="javascript">

//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function frmsubmit(page){
	if(frmSearch.itemid.value!=''){
		if (!IsDouble(frmSearch.itemid.value)){
			alert('상품코드는 숫자만 가능합니다.');
			frmSearch.itemid.focus();
			return;
		}
	}

	frmSearch.page.value=page;
	frmSearch.submit();
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="get" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<select name="selType">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>할인코드</option>
			<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>이벤트코드</option>
			<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>할인명</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>
		상품코드 : <input type="text" name="itemid" value="<%=itemid%>" size="10" maxlength="10">
		상품명 : <input type="text" name="itemname" value="<%=itemname%>" size="20" maxlength="20">
		<br>상태:
		<% sbGetOptCommonCodeArr_off "salestatus", isStatus, True, False,"onChange='javascript:document.frmSearch.submit();'"%>
		매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<input type="checkbox" name="adminvspos" value="ON" <% if adminvspos = "ON" then response.write " checked" %>>어드민포스할인가격서로상이
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>

</form>
</table>
<!---- /검색 ---->

<br>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
	<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right"></td>
	</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="left">검색결과 : <b><%=clsSaleItem.ftotalcount%></b>&nbsp;&nbsp;페이지 : <b><%=page%> / <%=clsSaleItem.FTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
	<td align="center" >이미지</td>
	<td align="center">
		상품ID<br>할인코드(이벤트코드)
	</td>

	<td align="center">브랜드</td>
	<td align="center">
		상품명<br>할인명
	</td>
	<td align="center">매장</td>
	<td align="center">할인상태</td>
	<td align="center">현재판매가</td>
	<td align="center">현재매입가<br>현재샾공급가</td>
	<td align="center">현재매입마진<br>현재샾공급마진</td>
	<td align="center">할인판매가</td>
    <td align="center">할인마진</td>
	<td align="center">할인매입가<br>할인샾공급가</td>
	<td align="center">할인매입마진<br>할인샾공급마진</td>
	<td align="center">적립포인트</td>
</tr>
<% IF clsSaleItem.fresultcount > 0 THEN %>
<% For i = 0 To clsSaleItem.fresultcount -1 %>
<%
mSPrice  =clsSaleItem.FItemList(i).forgsellprice - (clsSaleItem.FItemList(i).forgsellprice*(isRate/100))
mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,clsSaleItem.FItemList(i).forgsellprice,clsSaleItem.FItemList(i).fshopsuplycash,mSPrice,clsSaleItem.FItemList(i).fcomm_cd)
if mSPrice<>0 then iSaleMargin =  100-fix(mSBPrice/mSPrice*10000)/100
 if clsSaleItem.FItemList(i).forgsellprice<>0 then iOrgMargin= 100-fix(clsSaleItem.FItemList(i).fshopsuplycash/clsSaleItem.FItemList(i).forgsellprice*10000)/100

'/할인마진
sellpricemargin = 0
if clsSaleItem.FItemList(i).fshopitemprice<>0 then
	sellpricemargin = 100-fix(clsSaleItem.FItemList(i).fsaleprice/clsSaleItem.FItemList(i).fshopitemprice*10000)/100
end if
%>
<% if cint(clsSaleItem.FItemList(i).fsaleItem_status) <> 8 then %>
	<tr align="center" bgcolor="#FFFFFF">
<% else %>
	<tr align="center" bgcolor="silver">
<% end if %>
<form name="frmBuyPrc_<%=clsSaleItem.FItemList(i).fitemid%>" >
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
    <td>
    	<%IF clsSaleItem.FItemList(i).fsmallimage <> "" THEN%>
    		<img src="<%=clsSaleItem.FItemList(i).fsmallimage%>" width=50 height=50>
    	<%END IF%>
    </td>
    <td>
    	<%=clsSaleItem.FItemList(i).fitemgubun%>-<%=clsSaleItem.FItemList(i).fitemid%>-<%=clsSaleItem.FItemList(i).fitemoption%>
    	<br><%=clsSaleItem.FItemList(i).fsale_code%>
    	<% if clsSaleItem.FItemList(i).fevt_code <> 0 then response.write " ("& clsSaleItem.FItemList(i).fevt_code & ")" %>
    </td>

    <td>
    	<%=db2html(clsSaleItem.FItemList(i).fmakerid)%><br><%= fnColor(clsSaleItem.FItemList(i).fcentermwdiv,"mw") %>&nbsp;<%=clsSaleItem.FItemList(i).fcentermwdiv%>
    </td>
    <td>
    	<%=db2html(clsSaleItem.FItemList(i).fshopitemname)%><br><%=db2html(clsSaleItem.FItemList(i).fsale_name)%>
    </td>
    <td>
    	<%= clsSaleItem.FItemList(i).fshopid %>
    </td>
    <td>
    	포스할인 :
    	<% if isStatus = "8" and clsSaleItem.FItemList(i).fsaleyn = "Y" then %>
    		<font color="red"><%=clsSaleItem.FItemList(i).fsaleyn%> (타할인)</font>
    	<% elseif clsSaleItem.FItemList(i).fsaleyn = "Y" then %>
    		<font color="red"><%=clsSaleItem.FItemList(i).fsaleyn%></font>
    	<% else %>
    		<font color="blue"><%=clsSaleItem.FItemList(i).fsaleyn%></font>
    	<% end if %>

    	<Br>어드민상태(<%=clsSaleItem.FItemList(i).fsaleItem_status%>) : <font color="blue"><%=fnGetCommCodeArrDesc_off(arrsalestatus,clsSaleItem.FItemList(i).fsaleItem_status)%></font>
    </td>
    <td align="right">
    	<!--현재판매가-->
    	<%=formatnumber(clsSaleItem.FItemList(i).fshopitemprice,0)%>
    </td>
    <td align="right">
    	<%=formatnumber(clsSaleItem.FItemList(i).fshopsuplycash,0)%><!--현재매입가-->
    	<br><%=formatnumber(clsSaleItem.FItemList(i).fshopbuyprice,0)%><!--현재샾판매가-->
    </td>
    <td align="right">
    	<% if clsSaleItem.FItemList(i).fshopitemprice<>0 then %><!--현재마진율-->
			<%= 100-fix(clsSaleItem.FItemList(i).fshopsuplycash/clsSaleItem.FItemList(i).fshopitemprice*10000)/100 %>%
		<% end if %>

    	<% if clsSaleItem.FItemList(i).fshopitemprice<>0 then %><!--현재샾판매마진율-->
			<br><%= 100-fix(clsSaleItem.FItemList(i).Fshopbuyprice/clsSaleItem.FItemList(i).fshopitemprice*10000)/100 %>%
		<% end if %>
	</td>
	<%IF cint(clsSaleItem.FItemList(i).fsaleItem_status) = 8 or  cint(clsSaleItem.FItemList(i).fsaleItem_status) = 9 THEN%>
		<td align="right">0<Br>0</td>
		<td align="right">0%</td>
		<td align="right">0<Br>0</td>
	    <td align="right">0%</td>
	    <td align="right">0%</td>
	<%ELSE%>
	    <td align="right">
			<%=formatnumber(clsSaleItem.FItemList(i).fsaleprice,0)%>
	    	<%
	    	if clsSaleItem.FItemList(i).fsale_status = "6" and clsSaleItem.FItemList(i).fsaleItem_status = "6" and clsSaleItem.FItemList(i).fpossaleprice <> "" then
	    	%>
	    		<br><font color="red">현재포스적용가격 : <%=formatnumber(clsSaleItem.FItemList(i).fpossaleprice,0)%></font>
	    	<% end if %>
	    </td>
        <td align="right">
			<%= sellpricemargin %>%
		</td>
        <td align="right">
	    	<%=formatnumber(clsSaleItem.FItemList(i).fsalesupplycash,0)%>
	    	<br><%=formatnumber(clsSaleItem.FItemList(i).fsaleshopsupplycash,0)%>
	    </td>
	    <td align="right">
	    	<%
	    	if clsSaleItem.FItemList(i).fsaleprice<>0 then smargin= 100-fix(clsSaleItem.FItemList(i).fsalesupplycash/clsSaleItem.FItemList(i).fsaleprice*10000)/100
	    	if clsSaleItem.FItemList(i).fsaleprice<>0 then sshopmargin= 100-fix(clsSaleItem.FItemList(i).fsaleshopsupplycash/clsSaleItem.FItemList(i).fsaleprice*10000)/100
	    	%>
			<%=smargin%>%
			<br><%=sshopmargin%>%
	    </td>
	    <td align="right">
	    	<%= clsSaleItem.FItemList(i).fpoint_rate %>%
	    </td>
	<%END IF%>
</form>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if clsSaleItem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= clsSaleItem.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + clsSaleItem.StartScrollPage to clsSaleItem.StartScrollPage + clsSaleItem.FScrollCount - 1 %>
			<% if (i > clsSaleItem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(clsSaleItem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if clsSaleItem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="FFFFFF">
	<td colspan="20" align="center">
		검색된 결과가 없습니다
	</td>
</tr>
<% END IF %>
</table>

<%
set clsSaleItem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
