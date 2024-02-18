<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%

dim itemid, makerid, itemname, waititemid
dim sellyn, isusing, danjongyn, limityn, mwdiv
dim page, cdl, cdm, cds, dispCate
dim infodivYn, itemdiv,overseaYN

itemid  = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)
itemname = RequestCheckVar(request("itemname"),32)

sellyn  = RequestCheckVar(request("sellyn"),10)
isusing = RequestCheckVar(request("isusing"),10)
danjongyn = RequestCheckVar(request("danjongyn"),10)
limityn = RequestCheckVar(request("limityn"),10)
mwdiv = RequestCheckVar(request("mwdiv"),10)

page = RequestCheckVar(request("page"),10)

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)
infodivYn  = requestCheckvar(request("infodivYn"),10)
waititemid = requestCheckvar(request("waititemid"),10)
itemdiv = requestCheckvar(request("itemdiv"),2)
overseaYN= requestCheckvar(request("overseaYN"),1)
if (sellyn="") then sellyn="A"

if (page="") then page=1

''if (isusing="") then isusing="Y"
''사용하는 상품만 표시로 변경
isusing="Y"

'상품코드 유효성 검사(2008.08.01;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemid = itemid
oitem.FRectItemName = itemname
oitem.FRectDanjongyn = danjongyn
oitem.FRectLimityn = limityn
oitem.FRectMWDiv = mwdiv
oitem.FPageSize = 30
oitem.FCurrPage = page
oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectInfodivYn    = infodivYn
oitem.FRectSellReserve = "Y"
oitem.FRectwaititemid  = waititemid
oitem.FRectItemDiv  = itemdiv
oitem.FRectdeliverOverseas = overseaYN

if (sellyn <> "A") then
    oitem.FRectSellYN = sellyn
end if

if (isusing <> "A") then
    oitem.FRectIsUsing = isusing
end if


oitem.GetProductList

dim i

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(ipage){
	document.frm.page.value= ipage;
	SubmitSearch();
}
function SubmitSearch(){
	document.frm.action = "/designer/itemmaster/upche_item_list.asp";
	document.frm.target = "";

	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("상품코드에는 숫자만 입력이 가능합니다.");
	    document.frm.itemid.focus();
	    return;
    }
	document.frm.submit();
}


// ============================================================================
// 기본정보수정
function editItemInfo(itemid) {

	var param = "itemid=" + itemid;
	popwin = window.open('upche_item_infomodify.asp?' + param ,'editItemInfoPop','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// 옵션수정
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('upche_item_optionmodify.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function editSimpleItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_upche_simpleitemedit.asp?' + param ,'editSimpleItemOption','width=500,height=650,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// 이미지수정
function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('upche_item_imagemodify.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//엑셀다운
function nowListExcelDown()
{
	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("상품코드에는 숫자만 입력이 가능합니다.");
	    document.frm.itemid.focus();
	    return;
    }

	document.frm.action = "/designer/itemmaster/upche_item_list_XL.asp";
	document.frm.target = "XLdown";
	document.frm.submit();
}

//엑셀다운_옵션포함
function nowListExcelDownOption(){
	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("상품코드에는 숫자만 입력이 가능합니다.");
	    document.frm.itemid.focus();
	    return;
    }

	document.frm.action = "/designer/itemmaster/upche_item_list_option_XL.asp";
	document.frm.target = "XLdown";
	document.frm.submit();
}

//품목정보 일괄변경 팝업
function popUploadXLSItemInfo() {
	popwin = window.open('pop_item_infoUploadFile.asp','popInfoUpload','width=520,height=300,scrollbars=no,resizable=no');
	popwin.focus();
}

//안전인증정보 일괄변경 팝업
function popUploadXLSSafetyInfo() {
	popwin = window.open('./itemInfoFile/pop_item_safetyinfoUploadFile.asp','popInfoUpload','width=520,height=270,scrollbars=no,resizable=no');
	popwin.focus();
}

//해외배송정보 일괄변경 팝업
function popUploadXLSOverSeaInfo(){
    popwin =  window.open('./itemInfoFile/pop_item_overseainfoUploadFile.asp','popInfoUpload','width=520,height=270,scrollbars=no,resizable=no');
	popwin.focus();
}
</script>


<!-- 표 상단바 시작-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) SubmitSearch();">
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="20" onKeyPress="if (event.keyCode == 13) SubmitSearch();">
			<br>
			관리<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			&nbsp; 전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->
			<input type="hidden" name="waititemid" value=""> <!-- for play auto -->
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:SubmitSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
			단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	한정:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
	     	<font color="red">품목정보입력여부</font>
	     	<select class="select" name="infodivYn">
            <option value="">전체</option>
            <option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >입력이전</option>
            <option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >입력완료</option>
            </select>
	     	&nbsp;
			상품구분:<% drawSelectBoxItemDiv "itemdiv", itemdiv %>
			&nbsp;
			<font color="red">해외배송여부</font>
			<select class="select" name="overseaYN">
            <option value="">전체</option>
            <option value="N" <%= CHKIIF(overseaYN="N","selected","") %> >N</option>
            <option value="Y" <%= CHKIIF(overseaYN="Y","selected","") %> >Y</option>
            </select>

		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" class="a" >
<tr>
	<td align="left" style="padding-top:5px;">
		<input type="button" class="button" style="width:240px;background-color:#F8DFF0;" value="[상품정보고시관련] 추가정보 일괄등록" onclick="popUploadXLSItemInfo()" title="Excel파일을 업로드하여 [상품정보고시관련] 추가정보 일괄등록합니다." /> &nbsp;
		<input type="button" class="button" style="width:190px;background-color:#DFF8F0;" value="[안전인증대상]정보 일괄등록" onclick="popUploadXLSSafetyInfo()" title="Excel파일을 업로드하여 [안전인증대상]정보 일괄등록합니다." />
		<input type="button" class="button" style="width:190px;" value="[해외배송]정보 일괄등록" onclick="popUploadXLSOverSeaInfo()" title="Excel파일을 업로드하여 [해외배송]정보 일괄등록합니다." />
	</td>
	<td align="right" style="padding:5 0 5 0;">
	    <img src="/images/btn_excel.gif" style="cursor:pointer;display:inline;position:relative;top:5px" onClick="nowListExcelDown()" alt="상품목록엑셀" title="상품목록 다운로드">(상품목록)
	    &nbsp;
	    <img src="/images/btn_excel.gif" style="cursor:pointer;display:inline;position:relative;top:5px" onClick="nowListExcelDownOption()" alt="옵션포함상품목록엑셀" title="상품목록 다운로드(옵션포함)">(옵션포함)
	</td>
</tr>
</table>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr bgcolor="#FFFFFF">
	        <td colspan="14" align="right">총건수 : <%= oitem.FTotalCount %> </td>
	    </tr>
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="60">상품코드</td>
			<td width="50">이미지</td>
			<td width="100">브랜드ID</td>
			<td>상품명</td>
			<td width="60">판매가</td>
			<td width="60">공급가</td>
			<td width="40">마진</td>
			<td width="30">거래<br>구분</td>
			<td width="30">판매여부</td>
			<td width="40">한정<br>여부</td>
			<td width="40">해외배송<br>여부</td>
			<td width="50">기본<br>정보</td>
			<td width="50">이미지</td>
			<td width="70">옵션/한정<br>판매관련</td>
	    </tr>
<% if oitem.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="14" align="center">[검색결과가 없습니다.]</td>
	    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
    	<% if (oitem.FItemList(i).Fisusing = "N") then %>
    	<tr class="a" height="25" bgcolor="<%= adminColor("gray") %>">
		<% else %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center"><a href="http://www.10x10.co.kr/<%= oitem.FItemList(i).Fitemid %>" target="_blank"><%= oitem.FItemList(i).Fitemid %></a></td>
			<td align="center"><img src="<%= oitem.FItemList(i).FImgSmall %>" width="50" height="50" border="0" alt=""></td>
			<td align="center"><%= oitem.FItemList(i).Fmakerid %></td>
			<td align="left"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><font color="blue">(확인하기)</font></a></td>
			<td align="right">
			    <%= FormatNumber(oitem.FItemList(i).Forgprice,0) %>
			    <%
			    '할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				end Select
			end if
			    %>
			</td>
			<td align="right"><%= FormatNumber(oitem.FItemList(i).Forgsuplycash,0) %>
			    <%
			    '할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
					if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
					end if
				end if
			end if
			    %>
			</td>
			<td align="right">
			<%
			Response.Write fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1)
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & fnPercent(oitem.FItemList(i).Fsailsuplycash,oitem.FItemList(i).Fsailprice,1) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						end if
					Case "2"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						end if
				end Select
			end if
		%>
	        </td>
			<td align="center">
				<font color="<%= mwdivColor(oitem.FItemList(i).Fmwdiv) %>"><%= mwdivName(oitem.FItemList(i).Fmwdiv) %></font>
			</td>

			<td align="center">
				<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
			<%IF oitem.FItemList(i).Fsellreservedate <>"" THEN%><div>오픈예약: <%=oitem.FItemList(i).Fsellreservedate%></div><%END IF%>
			</td>
			<td align="center">
        		<% if (oitem.FItemList(i).Flimityn = "Y") then %>
             		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
             		<br>(<%= (oitem.FItemList(i).Flimitno - oitem.FItemList(i).Flimitsold) %>)
        		<% else %>
              		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
       			<% end if %>
			</td>
			<td align="center"><%=fnColor(oitem.FItemList(i).FdeliverOverseas,"yn")%>
		    </td>
		    <td align="center">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom" onClick="editItemInfo('<%= oitem.FItemList(i).FItemId %>');" style="cursor:pointer">
		    </td>
		    <td align="center">
		    	<a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>')">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		    	</a>
		    </td>
		    <td align="center">
        <% if (oitem.FItemList(i).Fmwdiv = "U") then %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		      	</a>
        <% else %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<b>[</b>수정요청<b>]</b>
		      	</a>
        <% end if %>

		    </td>
		</tr>
		<% next %>
	</table>
<% end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<iframe id="XLdown" name="XLdown" src="about:blank" frameborder="0" width="110" height="110"></iframe>

<% set oitem = nothing %>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
