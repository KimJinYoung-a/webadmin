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
mduserid    = RequestCheckVar(request("mduserid"),32) '���MD
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

If noinsert = "K" Then ''�ʵ崩�� ��ǰ���� �귣��
    oitem.FPageSize = 100                   ''���� �ٻѸ�.
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

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="research" value="on" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			&nbsp;
			�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
			<br>
			�����ON : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
			<input class="button" type="button" value="Me" onClick="this.form.mduserid.value='<%=session("ssBctId")%>'">
			<!-- ��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"> -->
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
			<input type="radio" name="noinsert" value="" <%=chkiif(noinsert="","checked","")%> />��ü �귣�� 
			&nbsp;&nbsp;
			<input type="radio" name="noinsert" value="Y" <%=chkiif(noinsert="Y","checked","")%> />���Է� ��ǰ���� �귣�� 
			&nbsp;&nbsp;
			<input type="radio" name="noinsert" value="K" <%=chkiif(noinsert="K","checked","")%> />�ʵ崩�� ��ǰ���� �귣�� 
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
		�Ǹ�����(ItemScore)=�ֱ��Ǹ�(2��)*10 + (�ֱ��Ǹ�(2��) * (�ǸŰ�/10,000))/4 + �ֱ����ø���Ʈ(5��)*2 + (�ֱ��ı�����Ʈ(7��)/5) + ���Ǹ�(����)/30
	    </td>
	</tr>
    </form>
</table>

<!-- ����Ʈ ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= oitem.FTotalCount%>�� �귣��</b> 
			<% If noinsert <> "K" Then %>
			(��ǰ :<b><%=oitem.FtotitemCnt%></b>) (��� �Ϸ��ǰ :<b><%=oitem.FtotFinCnt%></b>) (�̵�� ��ǰ:<b><%=oitem.FtotNoFinCnt%></b>)
			<% Else %>
			(�̵�� ��ǰ :<b><%=oitem.FtotitemCnt%></b>) 
			<% End If %>
			&nbsp;
			������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<%  If noinsert ="K" Then   %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%">�귣��ID</td>
		<td width="25%">�̵�� ��ǰ��</td>
		<td width="25%">�����ON</td>
    </tr>
	<% Else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%">�귣��ID</td>
		<td width="16%">�˻� ��ǰ��</td>
		<td width="16%">��� �Ϸ�� ��ǰ��</td>
		<td width="16%">�̵�ϵ� ��ǰ��</td>
		<td width="16%">�����ON</td>
		<td width="11%">�Ǹ��������<%=CHKIIF(noinsert="","<br>(��ü ����)","<br>(�̵�� ����)")%></td>
		<td width="8%">���곻��</td>
    </tr>
	<% End If %>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
	<% If noinsert ="K" Then %>
		<% for i=0 to oitem.FresultCount-1 %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','K')" title="�귣�� ����Ʈ����"><%= oitem.FItemList(i).Fmakerid	%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','K')" title="�̵�ϵ� ��ǰ��"><%= oitem.FItemList(i).Fitemcnt%></a></td>
			<td align="center"><%= oitem.FItemList(i).Fmdname%></td>
		</tr>
		<% next %>
	<% Else %>
		<% for i=0 to oitem.FresultCount-1 %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','')" title="�귣�� ����Ʈ����"><%= oitem.FItemList(i).Fmakerid	%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','')" title="��ϵ� ��ǰ��"><%= oitem.FItemList(i).Fitemcnt%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','Y')" title="��� �Ϸ�� ��ǰ��"><%= oitem.FItemList(i).Ffincnt%></a></td>
			<td align="center"><a href="javascript:popBrandlist('<%= oitem.FItemList(i).Fmakerid %>','N')" title="�̵�ϵ� ��ǰ��"><%=oitem.FItemList(i).Fitemcnt - oitem.FItemList(i).Ffincnt%></a></td>
			<td align="center"><%= oitem.FItemList(i).Fmdname%></td>
			<td align="center"><%= formatnumber(oitem.FItemList(i).FAvgScore,2) %></td>
			<td align="center"><a href="javascript:PopBrandAdminUsingChange('<%= oitem.FItemList(i).Fmakerid %>');">����</a></td>
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
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->