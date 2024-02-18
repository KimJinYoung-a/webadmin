<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epItemManageCls.asp"-->
<%
Dim oFixedItem,page, i
page		= requestCheckvar(request("page"),10)

Dim itemid :  itemid = requestCheckvar(request("itemid"),10)
Dim research : research = requestCheckvar(request("research"),10)
Dim showimage : showimage = requestCheckvar(request("showimage"),10)
Dim makerid : makerid = requestCheckvar(request("makerid"),32)
Dim sellyn : sellyn = requestCheckvar(request("sellyn"),10)
Dim mwdiv : mwdiv = requestCheckvar(request("mwdiv"),10) 
Dim useyn : useyn = requestCheckvar(request("useyn"),10) 

If page = "" Then page = 1
''if (research="") and (showimage="") then showimage="on"

' itemidarr = replace(itemidarr,"'","")
' itemidarr = replace(itemidarr,vbCRLF,",")
' itemidarr = replace(itemidarr,vbCR,",")
' itemidarr = replace(itemidarr,vbLf,",")

if NOT isNumeric(itemid) then itemid=""

Dim oEpAdItem
SET oEpAdItem = new CNvEpAdList
	oEpAdItem.FCurrPage		= page
	oEpAdItem.FPageSize		= 50
    oEpAdItem.FRectitemid   = itemid
	

	if (itemid<>"") then
		oEpAdItem.getEpAdGetOneItem
	end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">

<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        상품코드 : <input type="text" name="itemid" value="<%=itemid%>" size="10" maxlength="10">
        <% if (FALSE) then %>
        &nbsp;&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        <% end if %>
		
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<% if (FALSE) then %>
<tr bgcolor="#FFFFFF">
	<td>
        판매상태 : 
            <select name="sellyn" class="select">
                <option value="" <%= CHkIIF(sellyn="","selected","") %> >전체
                <option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
                <option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
            </select>&nbsp;
        &nbsp;&nbsp;
        매입구분 : 
        <% Call drawSelectBoxMWU("mwdiv",mwdiv) %>
        &nbsp;&nbsp;
        사용여부 : 
            <select name="useyn" class="select">
                <option value="" <%= CHkIIF(useyn="","selected","") %> >전체
                <option value="Y" <%= CHkIIF(useyn="Y","selected","") %> >사용
                <option value="N" <%= CHkIIF(useyn="N","selected","") %> >사용않함
            </select>&nbsp;
        &nbsp;&nbsp;
            
        &nbsp;&nbsp;
        <input type="checkbox" name="showimage" <%=CHKIIF(showimage="on","checked","")%> >이미지표시
	</td>
</tr>
<% end if %>
</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= FormatNumber(oEpAdItem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEpAdItem.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if (showimage="on") then %>
	<td width="50">이미지</td>
	<% end if %>

	<td width="50">광고계정</td>
	
    <td width="100">캠페인 아이디</td>
    <td width="120">캠페인명</td>
	<td width="100">광고그룹 아이디</td>

	<td width="90">광고그룹명</td>
    <td width="90">소재 아이디</td>

	<td width="40">On/Off</td>

	<td width="70">네이버 상품코드</td>
	<td width="70">자사 상품코드</td>
    <td width="90">자사 상품명</td>
    <td width="80">광고 상품명</td>
    
    <td width="30">비고</td>

</tr>
<% if oEpAdItem.FResultCount<1 then %>
<tr align="center" bgcolor="#FFFFFF">
    <% if itemid="" then %>
    <td colspan="12">상품코드를 입력하세요.</td>
    <% else %>
    <td colspan="12">검색결과가 없습니다.</td>
    <% end if %>
</tr>
<% else %>
<% For i=0 to oEpAdItem.FResultCount - 1 %>
<tr align="center" bgcolor="<%=CHKIIF(LCASE(oEpAdItem.FItemList(i).FOnOff)<>"on","#DDDDDD","#FFFFFF")%>">
	<% if (showimage="on") then %>
	<td><img src="<%= oEpAdItem.FItemList(i).FImageSmall%>" width="50"></td>
	<% end if %>
    <td><%= oEpAdItem.FItemList(i).FAccountId %></td>
    <td><%= oEpAdItem.FItemList(i).FCampaignId %></td>
	<td><%= oEpAdItem.FItemList(i).FCampaignNm %></td>
    <td>
		<%= oEpAdItem.FItemList(i).FAdGroupId %>
	</td>
	<td><%= oEpAdItem.FItemList(i).FAdGroupNm %></td>
	<td><%= oEpAdItem.FItemList(i).FAdId %></td>
    <td><%= oEpAdItem.FItemList(i).FOnOff %></td>
    <td><%= oEpAdItem.FItemList(i).FProductNo %></td>
    <td><%= oEpAdItem.FItemList(i).FProductNoMall %></td>
    <td><%= oEpAdItem.FItemList(i).FProductNm %></td>

	<td><%= oEpAdItem.FItemList(i).FAdProductNm %></td>
	<td></td>
</tr>
<% Next %>
<% end if %>
<% if (FALSE) then %>
<tr height="20">
    <td colspan="12" align="center" bgcolor="#FFFFFF">
        <% if oEpAdItem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEpAdItem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEpAdItem.StartScrollPage to oEpAdItem.FScrollCount + oEpAdItem.StartScrollPage - 1 %>
    		<% if i>oEpAdItem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEpAdItem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<% end if %>
</table>
<%
SET oEpAdItem = Nothing
%>
<p>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->