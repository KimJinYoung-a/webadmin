<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 상품 카테고리 관리
' Hieditor : 2013.05.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<!-- #include virtual="/lib/classes/ithinkso/category/category_cls_ithinkso.asp"-->

<%
Dim oitem, i, page, itemid, itemname, makerid,CateSeq0, CateSeq1, CateSeq2, CateSeq3, sellyn, usingyn
dim menupos
	menupos = request("menupos")
	page = request("page")
	itemid		= request("itemid")
	itemname	= request("itemname")
	makerid		= request("makerid")
	sellyn		= request("sellyn")
	usingyn		= request("usingyn")
	CateSeq0 = request("CateSeq0")
	CateSeq1 = request("CateSeq1")
	CateSeq2 = request("CateSeq2")
	CateSeq3 = request("CateSeq3")

if (page = "") then page = 1

if CateSeq0 = "" or CateSeq1 = "" or CateSeq2 = "" then	
	response.write "<script language='javascript'>"
	response.write "	alert('카테고리가 지정되지 않았습니다.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
end if

set oitem = new ccategory_ithinkso
	oitem.FPageSize         = 50
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.frectCateTypeSeq   = CateSeq0
	oitem.FRectCateSeq1   = CateSeq1
	oitem.FRectCateSeq2     = CateSeq2
	oitem.FRectCateSeq3   = CateSeq3
	oitem.frectcountryCd = "ITSWEB"
	oitem.getitemlist
%>
<script type="text/javascript">

	function frmsubmit(page){
		frmSearch.page.value=page;
		frmSearch.submit();
	}
	
	//카테고리 신규저장 
	function categoryitemreg(upfrm){

		if (!CheckSelected()){
				alert('선택아이템이 없습니다.');
				return;
			}	
			var frm;
				for (var i=0;i<document.forms.length;i++){
					frm = document.forms[i];
					if (frm.name.substr(0,9)=="frmBuyPrc") {
						if (frm.cksel.checked){
							upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ','
								
						}
					}
				}
		upfrm.action='category_item_process_ithinkso.asp';
		upfrm.submit();
	}

</script>
			
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="post">
<input type="hidden" name="itemidarr">
<input type="hidden" name="CateSeq0" value="<%=CateSeq0%>">
<input type="hidden" name="CateSeq1" value="<%=CateSeq1%>">
<input type="hidden" name="CateSeq2" value="<%=CateSeq2%>">
<input type="hidden" name="CateSeq3" value="<%=CateSeq3%>">
<input type="hidden" name="menupos" value="<%= Request("menupos") %>">
<input type="hidden" name="mode" value="categoryitemreg">
<input type="hidden" name="page" >
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		<p>
		* 상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
		&nbsp;&nbsp;
		* 상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* 판매여부:
	   <select class="select" name="sellyn">
		   <option value="">전체</option>
		   <option value="Y"  <%=CHKIIF(sellyn="Y","selected","")%>>판매</option>
		   <option value="S"  <%=CHKIIF(sellyn="S","selected","")%>>일시품절</option>
		   <option value="N"  <%=CHKIIF(sellyn="N","selected","")%>>품절</option>
		   <option value="YS"  <%=CHKIIF(sellyn="YS","selected","")%>>판매+일시품절</option>
	   </select>
     	&nbsp;&nbsp;
     	* 사용여부:
	   <select class="select" name="usingyn">
		   <option value="">전체</option>
		   <option value="Y"  <%=CHKIIF(usingyn="Y","selected","")%>>사용함</option>
		   <option value="N"  <%=CHKIIF(usingyn="N","selected","")%>>사용안함</option>
	   </select>
	</td>
</tr>
</form>
</table>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" onclick="categoryitemreg(frmSearch);" class="button" value="선택신규저장">
    </td>
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				검색결과 : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
			</td>
			<td align="right">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>itemID</td>
	<td> 이미지</td>
	<td width="100">브랜드ID</td>
	<td>상품명</td>
	<td>판매가</td>
	<td>판매<br>여부</td>
	<td>사용<br>여부</td>
	<td>관리</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
<form action="" name="frmBuyPrc<%=i%>" method="get">
	<td align="center" width=20><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center" width=60>
		<input type="hidden" name="itemid" value="<%= oitem.FItemList(i).Fitemid %>">
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">				
		<%= oitem.FItemList(i).Fitemid %></a>
		</td>
	<td align="center" width=50><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
	<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
	<td align="right" width=80>
	<%
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
		'할인가
		if oitem.FItemList(i).Fsailyn="Y" then
			Response.Write "<br><font color=#F08050>(할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
		end if
		'쿠폰가
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
				Case "2"
					'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
			end Select
		end if
	%>
	</td>
	<td align="center" width=30><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
	<td align="center" width=30><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
    <td align="center" width=30>
    </td>
</form>    
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:frmsubmit('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:frmsubmit('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->