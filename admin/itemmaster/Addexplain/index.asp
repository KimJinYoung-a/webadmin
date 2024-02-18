<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
dim showminusmagin, marginup, margindown
dim page, research
Dim mduserid , noinsert 
dim cdl, cdm, cds
itemid      = requestCheckvar(request("itemid"),255)
itemname    = request("itemname")
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)
mduserid    = RequestCheckVar(request("mduserid"),32) '담당MD
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)

page = requestCheckvar(request("page"),10)
noinsert = requestCheckvar(request("noinsert"),10)
research = requestCheckvar(request("research"),10)
If sellyn = "" Then sellyn = "Y"

if (page="") then page=1
if (research="") then noinsert="Y"

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize			= 50
oitem.FCurrPage			= page
oitem.FRectMakerid     = makerid
oitem.FRectItemid			= itemid
oitem.FRectItemName  = itemname

oitem.FRectMWDiv        = mwdiv

oitem.FRectSellYN			= sellyn
oitem.FRectMduserid	= mduserid
oitem.FRectcheckYN		= noinsert
oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds

If noinsert = "K" Then ''필드누락 상품존재 브랜드
    oitem.FPageSize = 100                   ''느림 다뿌림.
	oitem.GetItemNotAddexplain_FieldBrand
Else
	oitem.GetItemNotAddexplainList
End If 

dim i

Dim addParameter
addParameter = "&sellYN="&sellyn&"&mwdiv="&mwdiv&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds

%>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
function popBrandlist(makerid,infodivYn)
{
	var popwin = window.open("pop_brandlist.asp?makerid=" + makerid + "&infodivYn="+ infodivYn +"<%=addParameter%>" ,"popitemContImage","width=1024 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="research" value="on" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			&nbsp;
			거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
			<br>
			담당자ON : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
			<input class="button" type="button" value="Me" onClick="this.form.mduserid.value='<%=session("ssBctId")%>'">
			<!-- 상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"> -->
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
			<input type="radio" name="noinsert" value="" <%=chkiif(noinsert="","checked","")%> />전체 브랜드 
			&nbsp;&nbsp;
			<input type="radio" name="noinsert" value="Y" <%=chkiif(noinsert="Y","checked","")%> />미입력 상품존재 브랜드 
			&nbsp;&nbsp;
			<input type="radio" name="noinsert" value="K" <%=chkiif(noinsert="K","checked","")%> />필드누락 상품존재 브랜드 
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
		판매지수(ItemScore)=최근판매(2일)*10 + (최근판매(2일) * (판매가/10,000))/4 + 최근위시리스트(5일)*2 + (최근후기포인트(7일)/5) + 총판매(러프)/30
	    </td>
	</tr>
    </form>
</table>

<!-- 리스트 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%= oitem.FTotalCount%>개 브랜드</b> 
			<% If noinsert <> "K" Then %>
			(상품 :<b><%=oitem.FtotitemCnt%></b>) (등록 완료상품 :<b><%=oitem.FtotFinCnt%></b>) (미등록 상품:<b><%=oitem.FtotNoFinCnt%></b>)
			<% Else %>
			(미등록 상품 :<b><%=oitem.FtotitemCnt%></b>) 
			<% End If %>
			&nbsp;
			페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<%  If noinsert ="K" Then   %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%">브랜드ID</td>
		<td width="25%">미등록 상품수</td>
		<td width="25%">담당자ON</td>
    </tr>
	<% Else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%">브랜드ID</td>
		<td width="16%">검색 상품수</td>
		<td width="16%">등록 완료된 상품수</td>
		<td width="16%">미등록된 상품수</td>
		<td width="16%">담당자ON</td>
		<td width="11%">판매지수평균<%=CHKIIF(noinsert="","<br>(전체 기준)","<br>(미등록 기준)")%></td>
		<td width="8%">정산내역</td>
    </tr>
	<% End If %>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
	<% If noinsert ="K" Then %>
		<% for i=0 to oitem.FresultCount-1 %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','K')" title="브랜드 리스트보기"><%= oitem.FItemList(i).Fmakerid	%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','K')" title="미등록된 상품수"><%= oitem.FItemList(i).Fitemcnt%></a></td>
			<td align="center"><%= oitem.FItemList(i).Fmdname%></td>
		</tr>
		<% next %>
	<% Else %>
		<% for i=0 to oitem.FresultCount-1 %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','')" title="브랜드 리스트보기"><%= oitem.FItemList(i).Fmakerid	%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','')" title="등록된 상품수"><%= oitem.FItemList(i).Fitemcnt%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','Y')" title="등록 완료된 상품수"><%= oitem.FItemList(i).Ffincnt%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','N')" title="미등록된 상품수"><%=oitem.FItemList(i).Fitemcnt - oitem.FItemList(i).Ffincnt%></a></td>
			<td align="center"><%= oitem.FItemList(i).Fmdname%></td>
			<td align="center"><%= formatnumber(oitem.FItemList(i).FAvgScore,2) %></td>
			<td align="center"><a href="javascript:PopBrandAdminUsingChange('<%= oitem.FItemList(i).Fmakerid %>');">보기</a></td>
		</tr>
		<% next %>
	<% End If %>
	<!-- paging -->		
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
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
	</tr>
	
</table>
<% end if %>

<%
SET oitem = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->