<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim makerid, deliverytype, isMapping, oCoupang
Dim page, i
page    		= request("page")
makerid			= request("makerid")
deliverytype	= request("deliverytype")
isMapping		= request("isMapping")
If page = "" Then page = 1

SET oCoupang = new CCoupang
	oCoupang.FCurrPage					= page
	oCoupang.FPageSize					= 30
	oCoupang.FRectMakerId				= makerid
	oCoupang.FRectDeliveryType			= deliverytype
	oCoupang.FRectIsMapping				= isMapping
	oCoupang.getTenCoupangBrandDeliveryList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function popBrandDeliverMap(v, m){
    var pb = window.open("/admin/etc/coupang/popCoupangBrandMap.asp?id="+v+"&maeipdiv="+m,"popOptionAddPrc","width=800,height=400,scrollbars=yes,resizable=yes");
	pb.focus();
}
function CoupangSelectRegDeliveryProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 브랜드가 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('쿠팡에 선택하신 ' + chkSel + '개 브랜드의 출고지를 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGDELIVERY";
        document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		배송구분 :
		<select name="deliverytype" class="select">
			<option value="">전체</option>
			<option value="MW" <%= Chkiif(deliverytype = "MW", "selected", "") %>>텐바이텐배송</option>
			<option value="U" <%= Chkiif(deliverytype = "U", "selected", "") %>>업체배송</option>
		</select>&nbsp;
		매칭여부 :
		<select name="isMapping" class="select">
			<option value="">전체</option>
			<option value="Y" <%= Chkiif(isMapping = "Y", "selected", "") %>>매칭완료</option>
			<option value="N" <%= Chkiif(isMapping = "N", "selected", "") %>>미매칭</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="RIGHT">
		<input class="button" type="button" id="btnRegSel" value="출고지등록" onClick="CoupangSelectRegDeliveryProcess()();">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oCoupang.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCoupang.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="150">브랜드ID</td>
	<td>브랜드명(한글)</td>
	<td>브랜드명(영문)</td>
	<td>담당자</td>
	<td>택배사</td>
	<td>주소</td>
	<td>배송구분</td>
	<td>출고지코드</td>
	<td>수정</td>
</tr>
<% If oCoupang.FResultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oCoupang.FResultCount - 1
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF( (oCoupang.FItemList(i).FOutboundShippingPlaceCode <> "") ,"#FFFFFF","#CCCCCC") %>">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oCoupang.FItemList(i).FId %>"></td>
	<td><%= oCoupang.FItemList(i).FId %></td>
	<td><%= oCoupang.FItemList(i).FSocname_kor %></td>
	<td><%= oCoupang.FItemList(i).FSocname %></td>
	<td><%= oCoupang.FItemList(i).FDeliver_name %></td>
	<td><%= oCoupang.FItemList(i).FDivname %></td>
	<td><%= oCoupang.FItemList(i).FReturn_zipcode %>&nbsp;<%= oCoupang.FItemList(i).FReturn_address %>&nbsp;<%= oCoupang.FItemList(i).FReturn_address2 %></td>
	<td><%= ChkIIF(oCoupang.FItemList(i).FMaeipdiv="U","업체배송","텐바이텐배송") %></td>
	<td><%= oCoupang.FItemList(i).FOutboundShippingPlaceCode %></td>
	<td><input type="button" class="button" value="수정" onClick="popBrandDeliverMap('<%= oCoupang.FItemList(i).FId %>', '<%= oCoupang.FItemList(i).FMaeipdiv %>')"></td>
</tr>
<%
		Next
	End If
%>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oCoupang.HasPreScroll then %>
		<a href="javascript:goPage('<%= oCoupang.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oCoupang.StartScrollPage to oCoupang.FScrollCount + oCoupang.StartScrollPage - 1 %>
    		<% if i>oCoupang.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oCoupang.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% SET oCoupang = nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
