<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim mwdiv, isusing, itemid
dim page, research
dim sType, finType

itemid  = RequestCheckVar(request("itemid"),9)
isusing = "Y"
page    = RequestCheckVar(request("page"),9)
research= RequestCheckVar(request("research"),9)
mwdiv   = RequestCheckVar(request("mwdiv"),9)
sType   = RequestCheckVar(request("sType"),9)
finType = RequestCheckVar(request("finType"),9)

if (page="") then page=1
if (research="") then
    if (isusing="") then isusing="Y"
    if (mwdiv="") then mwdiv="MW"
    if (sType="") then sType="S"
    if (finType="") then finType="on"
end if

'상품코드 유효성 검사(2008.08.04;허진원)
if trim(itemid)<>"" then
	if Not(isNumeric(trim(itemid))) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim oStatSList
set oStatSList = new CSummaryItemStock
oStatSList.FCurrPage    = page
oStatSList.FPageSize        = 30
oStatSList.FRectMakerid     = session("ssBctID")
oStatSList.FRectItemID      = itemid
oStatSList.FRectOnlyIsUsing = isusing
oStatSList.FRectMWDiv       = mwdiv
oStatSList.FRectState       = finType

oStatSList.GetImsiSoldOutList

dim i
%>
<script language='javascript'>
function popDanjongSet(iitemid, itemoption, actType){
    var popwin = window.open('/common/popitemdanjongSet.asp?itemid=' + iitemid + '&itemoption=' + itemoption + '&actType=' + actType,'popitemdanjongSet','width=900, height=400, scrollbars=yes, resizable=yes');
	popwin.focus();
}


function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();

}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
	    <td rowspan="2" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
        <td align="left">
            상품코드: <input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
            &nbsp;&nbsp;
        	배송구분:
        	   <select class="select" name="mwdiv">
               <option value="MW" <%= ChkIIF(mwdiv="MW","selected","") %> >매입+특정</option>
               <option value="W"  <%= ChkIIF(mwdiv="W","selected","") %> >특정</option>
               <option value="M"  <%= ChkIIF(mwdiv="M","selected","") %> >매입</option>
               </select>
            &nbsp;&nbsp;
            검색구분:
               <input type="radio" name="sType" value="S" <%= chkIIF(sType="S","checked","") %> > 일시품절상품
               <!-- <input type="radio" name="sType" value="N" <%= chkIIF(sType="N","checked","") %> > 판매중지상품 -->
            &nbsp;&nbsp;
            <input type="checkbox" name="finType" <%= ChkIIF(finType="on","checked","") %> >(단종/재입고 설정)미처리내역만
        	<br>
        </td>
        <td rowspan="2" width="50" bgcolor="#EEEEEE">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a><br>
        </td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<!-- 액션 시작 -->
<!--
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="right">
			<input type="button" class="button" value="전체선택" onClick="">
			&nbsp;
			마진율 : <input type="text" class="text" name="" size="3" maxlength="5">
			<input type="button" class="button" value="선택상품적용" onClick="">
			&nbsp;
			<input type="button" class="button" value="저장" onClick="">
		</td>

	</tr>
</table>
-->
<!-- 액션 끝 -->
<p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25">
			검색결과 : <b><%= oStatSList.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oStatSList.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#E6E6E6">
    	<!-- <td>선택</td> -->
		<td width="50">이미지</td>
		<!-- <td width="70">브랜드</td> -->
		<td width="40">상품<br>코드</td>
		<td>상품명<br>(옵션명)</td>
		<td width="35">배송<br>구분</td>
        <td width="25">전체<br>입고<br>반품</td>
        <td width="25">전체<br>판매<br>반품</td>
        <td width="25">전체<br>출고<br>반품</td>
        <td width="25">기타<br>출고<br>반품</td>
        <td width="25">CS<br>출고<br>반품</td>
		<td width="25">총<br>불량</td>
        <td width="25">총<br>실사<br>오차</td>
        <td width="25">실사<br>재고</td>
        <td width="25">총<br>상품<br>준비</td>
        <td width="25">재고<br>파악<br>재고</td>
        <td width="25">ON<br>결제<br>완료</td>
        <td width="25">ON<br>주문<br>접수</td>
        <td width="25">한정<br>비교<br>재고</td>
		<td width="40">판매<br>여부</td>
		<td width="50">한정<br>여부</td>
        <td width="35">단종<br>여부</td>
        <td width="60">재입고<br>예정일</td>
        <td width="35">단종<br>처리</td>
    </tr>
<% for i=0 to oStatSList.FresultCount-1 %>
    <% if oStatSList.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<!-- <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td> -->
		<td><img src="<%= oStatSList.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<!-- <td align="left"><%= oStatSList.FItemList(i).FMakerID %></td> -->
		<td>
          <%= oStatSList.FItemList(i).FItemID %>
        </td>
		<td align="left">
          <%= oStatSList.FItemList(i).FItemName %>
        <% if (oStatSList.FItemList(i).FItemOptionName <> "") then %>
          <br>(<font color="#3333CC"><%= oStatSList.FItemList(i).FItemOptionName %></font>)
        <% end if %>
        </td>
        <td><font color="<%= mwdivColor(oStatSList.FItemList(i).Fmwdiv) %>"><%= mwdivName(oStatSList.FItemList(i).Fmwdiv) %></font></td>
		<td><%= oStatSList.FItemList(i).Ftotipgono %></td>
		<td><%= -1*oStatSList.FItemList(i).Ftotsellno %></td>
		<td><%= oStatSList.FItemList(i).Foffchulgono + oStatSList.FItemList(i).Foffrechulgono %></td>
        <td><%= oStatSList.FItemList(i).Fetcchulgono + oStatSList.FItemList(i).Fetcrechulgono %></td>
        <td><%= oStatSList.FItemList(i).Ferrcsno %></td>
        <td><%= oStatSList.FItemList(i).Ferrbaditemno %></td>
        <td><%= oStatSList.FItemList(i).Ferrrealcheckno %></td>
        <td><b><%= oStatSList.FItemList(i).Frealstock %></b></td>
        <td><%= oStatSList.FItemList(i).Fipkumdiv5 + oStatSList.FItemList(i).Foffconfirmno %></td>
        <td><b><%= oStatSList.FItemList(i).GetCheckStockNo %></b></td>
        <td><%= oStatSList.FItemList(i).Fipkumdiv4 %></td>
        <td><%= oStatSList.FItemList(i).Fipkumdiv2 %></td>
        <td><b><%= oStatSList.FItemList(i).GetLimitStockNo %></b></td>
        <td>
        	<%= ChkIIF(oStatSList.FItemList(i).Fsellyn="S","일시<br>품절",ChkIIF(oStatSList.FItemList(i).Fsellyn="N","품절","")) %>
        </td>

        <td>
        <% if (oStatSList.FItemList(i).Flimityn = "Y") then %>
          	한정(<%= oStatSList.FItemList(i).GetLimitStr %>)
            <% if (oStatSList.FItemList(i).Foptlimityn = "Y") then %>
            <br>(<%= oStatSList.FItemList(i).Foptlimitno %>/<%= oStatSList.FItemList(i).Foptlimitsold %>)
            <% else %>
            <br>(<%= oStatSList.FItemList(i).FLimitNo %>/<%= oStatSList.FItemList(i).FLimitSold %>)
          	<% end if %>
        <% end if %>
        </td>
        <td><%= oStatSList.FItemList(i).getDanjongNameHTML %></td>
        <td>
            <% if (Not IsNull(oStatSList.FItemList(i).Fstockreipgodate)) then %>
            <a href="javascript:popDanjongSet('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>','R');"><%= oStatSList.FItemList(i).Fstockreipgodate %></a>
            <% else %>
            <a href="javascript:popDanjongSet('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>','R');"><img src="/images/icon_arrow_link.gif" width="14" border="0"></a>
            <% end if %>
        </td>
        <td>
            <% if (oStatSList.FItemList(i).FDanjongyn<>"M") and (oStatSList.FItemList(i).FDanjongyn<>"Y") then %>
            <a href="javascript:popDanjongSet('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>','D');"><img src="/images/icon_arrow_link.gif" width="14" border="0"></a>
            <% end if %>
        </td>

	</tr>
<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
		<% if oStatSList.HasPreScroll then %>
    		<a href="javascript:NextPage('<%= oStatSList.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oStatSList.StartScrollPage to oStatSList.FScrollCount + oStatSList.StartScrollPage - 1 %>
    		<% if i>oStatSList.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oStatSList.HasNextScroll then %>
    		<a href="javascript:NextPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
	    </td>
	</tr>
</table>

<%
set oStatSList = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
