<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim oitem
dim itemid
dim page, iColorCD, lp
Dim sColorName, sColorIcon, iSortNo, sIsUsing

iColorCD	= requestCheckVar(Request("iCD"),10)
itemid      = requestCheckVar(request("itemid"),10)
page = requestCheckVar(request("page"),10)

if (page="") then page=1

'==== 상품목록 접수 =======================================================
set oitem = new CItemColor
oitem.FRectColorCD	= iColorCD
oitem.FRectItemId	= itemid
oitem.FRectMakerId	= session("ssBctID")
oitem.FPageSize		= 20
oitem.FCurrPage		= page
oitem.FRectUsing	= "Y"
oitem.GetColorItemList
%>
<script language="javascript">
document.domain='10x10.co.kr';

//페이지 이동
function GoPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

//색상코드 선택
function selColorChip(cd) {
	document.frm.iCD.value= cd;
	document.frm.submit();
}

//상품색상 등록/수정
function jsItemColorReg(ccd,iid) {
	var winCItem;
	if(iid=="") {
		winCItem = window.open('/designer/itemmaster/popItemColorReg.asp?iCD='+ccd+'&iid='+iid,'popItemColor','width=580,height=250,scrollbars=yes');
	} else {
		winCItem = window.open('/designer/itemmaster/popItemColorReg.asp?iCD='+ccd+'&iid='+iid,'popItemColor','width=580,height=400,scrollbars=yes');
	}
	winCItem.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="iCD" value="<%=iColorCd%>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left"><%=FnSelectColorBar(iColorCd,25)%></td>
	</tr>
    </form>
</table>
<table width="100%" align="center" class="a">
<tr>
	<td align="right">
		<input type="button" value="신규등록" onclick="jsItemColorReg('','');"  class="button">
	</td>
</tr>
</table>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="9">
			검색결과 : <b><%= oitem.FTotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	</form>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>상품이미지</td>
		<td>컬러칩</td>
		<td>상품코드</td>
		<td>상품명</td>
		<td>브랜드</td>
		<td>계약구분</td>
		<td>판매여부</td>
		<td>한정여부</td>
		<td>등록일시</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="9" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for lp=0 to oitem.FresultCount-1 %>
	<tr align="center">
		<td bgcolor="#FFFFFF"><a href="javascript:jsItemColorReg(<%=oitem.FItemList(lp).FcolorCode%>,<%=oitem.FItemList(lp).FitemId%>);"><img src="<%=oitem.FItemList(lp).FsmallImage%>" border="0" width="50"></a></td>
		<td bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="1" bgcolor="#dddddd"><tr><td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="12" height="12" hspace="2" vspace="2"></td></tr></table></td>
		<td bgcolor="#FFFFFF"><a href="javascript:jsItemColorReg(<%=oitem.FItemList(lp).FcolorCode%>,<%=oitem.FItemList(lp).FitemId%>);"><%=oitem.FItemList(lp).FitemId%></a></td>
		<td bgcolor="#FFFFFF"><a href="javascript:jsItemColorReg(<%=oitem.FItemList(lp).FcolorCode%>,<%=oitem.FItemList(lp).FitemId%>);"><%=oitem.FItemList(lp).Fitemname%></a></td>
		<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FmakerId%></td>
		<td bgcolor="#FFFFFF"><%=fnColor(oitem.FItemList(lp).Fmwdiv,"mw")%></td>
		<td bgcolor="#FFFFFF"><%=fnColor(oitem.FItemList(lp).Fsellyn,"yn")%></td>
		<td bgcolor="#FFFFFF"><%=fnColor(oitem.FItemList(lp).Flimityn,"yn")%></td>
		<td bgcolor="#FFFFFF"><%=left(oitem.FItemList(lp).Fregdate,10)%></td>
    </tr>
	<% next %>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="9" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:GoPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for lp=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if lp>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(lp) then %>
    			<font color="red">[<%= lp %>]</font>
    			<% else %>
    			<a href="javascript:GoPage('<%= lp %>')">[<%= lp %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:GoPage('<%= lp %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
	
</table>
<% end if %>


<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->