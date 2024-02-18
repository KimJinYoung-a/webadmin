<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim diffItem, page, i, workdt, itemid, isusing, sellyn, notinitemid, notinmakerid, gubun
page			= requestCheckvar(request("page"),10)
workdt			= request("workdt")
itemid			= request("itemid")
isusing			= request("isusing")
sellyn			= request("sellyn")
notinitemid		= request("notinitemid")
notinmakerid	= request("notinmakerid")
gubun			= request("gubun")

If page = "" Then page = 1
If workdt = "" Then workdt = date

'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

Set diffItem = new epShop
	diffItem.FCurrPage			= page
	diffItem.FRectWorkdt		= workdt
	diffItem.FRectItemid		= itemid
	diffItem.FRectIsusing		= isusing
	diffItem.FRectSellyn		= sellyn
	diffItem.FRectNotinitemid	= notinitemid
	diffItem.FRectNotinmakerid	= notinmakerid
	diffItem.FRectGubun			= gubun
	diffItem.FPageSize	= 100
    diffItem.diffItemItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		체크일 :
		<input id="workdt" name="workdt" value="<%=workdt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="workdt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "workdt", trigger    : "workdt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		사용여부 :
		<select name="isusing" class="select">
			<option value="" <%= CHkIIF(isusing="","selected","") %>>전체</option>
			<option value="Y" <%= CHkIIF(isusing="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(isusing="N","selected","") %>>N</option>
		</select>
		&nbsp;
		판매여부 :
		<select name="sellyn" class="select">
			<option value="" <%= CHkIIF(sellyn="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
		</select>&nbsp;
		등록제외상품 :
		<select name="notinitemid" class="select">
			<option value="" <%= CHkIIF(notinitemid="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(notinitemid="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(notinitemid="N","selected","") %> >N
		</select>&nbsp;
		등록제외브랜드 :
		<select name="notinmakerid" class="select">
			<option value="" <%= CHkIIF(notinmakerid="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(notinmakerid="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(notinmakerid="N","selected","") %> >N
		</select>&nbsp;
		구분 :
		<select name="gubun" class="select">
			<option value="" <%= CHkIIF(gubun="","selected","") %> >전체
			<option value="A" <%= CHkIIF(gubun="A","selected","") %> >추가
			<option value="D" <%= CHkIIF(gubun="D","selected","") %> >삭제
		</select>&nbsp;
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(diffItem.FTotalPage,0) %> 총건수: <%= FormatNumber(diffItem.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>체크일</td>
    <td>상품코드</td>
    <td>판매가격</td>
    <td>사용여부</td>
    <td>판매여부</td>
	<td>등록제외상품</td>
	<td>등록제외브랜드</td>
	<td>상품최종수정일</td>
	<td>최종수정일<br/>월 차이</td>
	<td>최근판매수량</td>
	<td>구분</td>
</tr>
<% For i=0 to diffItem.FResultCount - 1 %>

<tr bgcolor="#FFFFFF" height="20" align="center">
	<td><%= diffItem.FItemList(i).FWorkdt %></td>
    <td><%= diffItem.FItemList(i).FItemid %></td>
    <td><%= FormatNumber(diffItem.FItemList(i).FSellprice,0) %></td>
    <td><%= diffItem.FItemList(i).FIsusing %></td>
    <td>
        <% if diffItem.FItemList(i).IsSoldOut then %>
            <% if diffItem.FItemList(i).FSellyn="N" then %>
            <font color="red">품절</font>
            <% else %>
            <font color="red">일시<br>품절</font>
            <% end if %>
        <% end if %>
    </td>
	<td><%= diffItem.FItemList(i).FNotinitemid %></td>
	<td><%= diffItem.FItemList(i).FNotinmakerid%></td>
	<td><%= diffItem.FItemList(i).FLastupdate %></td>
	<td><%= diffItem.FItemList(i).FDiffMonth %></td>
	<td><%= diffItem.FItemList(i).FRecentsellcount %></td>
	<td>
		<% if diffItem.FItemList(i).FGubun="D" then %>
			<font color="red">삭제</font>
		<% else %>
			<font color="blue">추가</font>
		<% end if %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if diffItem.HasPreScroll then %>
		<a href="javascript:goPage('<%= diffItem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + diffItem.StartScrollPage to diffItem.FScrollCount + diffItem.StartScrollPage - 1 %>
    		<% if i>diffItem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if diffItem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
