<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  [OFF]����_��ǰ����>>�Ż�ǰ����
' History : 2008.04.17 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer, page, itemid, locationid, datesearch, sdate, edate, itemname, IsOnlineItem, imageList, offmain
Dim sort, i, offlist, offsmall, iTotCnt, itemgubun, isusing, vParam, inc3pl
dim yyyy1, mm1, dd1, yyyy2, mm2, dd2, datefg, fromDate, toDate, cdl, cdm, cds
	designer = requestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	itemid = requestCheckVar(request("itemid"),10)
	datesearch = requestCheckVar(request("datesearch"),10)
	sdate = requestCheckVar(request("sdate"),10)
	edate = requestCheckVar(request("edate"),10)
	itemname = requestCheckVar(request("itemname"),124)
	itemgubun = requestCheckVar(request("itemgubun"),2)
	isusing = requestCheckVar(request("isusing"),1)
	sort = requestCheckVar(request("sort"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
If page = "" Then page = 1
If sort = "" Then sort = "itemregdate"

if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/�����ϰ�� ���� ���常 ��밡��
if (C_IS_SHOP) then
	
	'/���α��� ���� �̸�
	'if getlevel_sn("",session("ssBctId")) > 6 then
		locationid = C_STREETSHOPID
	'end if

else
	if (C_IS_Maker_Upche) then
		locationid = session("ssBctID")
	else
		locationid = request("locationid")
	end if
end if

vParam = "&locationid="&locationid&"&designer="&designer&"&itemid="&itemid&"&datesearch="&datesearch&"&sdate="&sdate&"&edate="&edate&"&itemgubun="&itemgubun&"&isusing="&isusing&"&itemname="&itemname&"&sort="&sort&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&datefg="&datefg&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&inc3pl="&inc3pl

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
	ioffitem.FRectShopid = locationid
	ioffitem.FRectDesigner = designer
	ioffitem.FRectDateSearch = datesearch
	ioffitem.FRectSDate = sdate
	ioffitem.FRectEDate = edate
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemName = itemname
	ioffitem.FRectItemId = itemid
	ioffitem.FRectIsusing = isusing
	ioffitem.FRectSorting = sort
	ioffitem.frectdatefg = datefg	
	ioffitem.FRectStartDay = fromDate
	ioffitem.FRectEndDay = toDate
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectInc3pl = inc3pl
	
	If locationid <> "" Then
		ioffitem.GetOffLineNewItemList
	End If

iTotCnt = ioffitem.FTotalCount
%>

<script language='javascript'>

function popOffItemEdit(ibarcode){
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function goSort(a){
	document.frm.sort.value = document.getElementById("tmpsort").value;
	document.frm.submit();
}

function goExcelDown(){
	var ExcelDown = window.open('offitemlist_xls.asp?1=1<%=vParam%>','ExcelDown','width=600,height=400,scrollbars=yes,resizable=yes');
	ExcelDown.focus();
}

function pop_ipgomaechul(shopid, extbarcode, yyyy1, mm1, dd1, yyyy2, mm2, dd2){
	var pop_ipgomaechul = window.open('/admin/offshop/dayitemsellsum.asp?shopid='+shopid+'&extbarcode='+extbarcode+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2,'pop_ipgomaechul','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_ipgomaechul.focus();
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value=1>
<input type="hidden" name="sort" value="<%=sort%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(locationid) <> "1" and locationid <> "" then %>
				* ShopID : <%=locationid%><input type="hidden" name="locationid" value="<%= locationid %>">
			<% else %>
				* ShopID : <% drawSelectBoxOffShopNotUsingAll "locationid",locationid %>
			<% end if %>
		<% else %>
			* ShopID : <% drawSelectBoxOffShopNotUsingAll "locationid",locationid %>
		<% end if %>
		&nbsp;&nbsp;
		* �Ⱓ : 
		<select name="datesearch" class="select">
			<option value="" <%=CHKIIF(datesearch="","selected","")%>>-����-</option>
			<option value="itemregdate" <%=CHKIIF(datesearch="itemregdate","selected","")%>>��ǰ�����</option>
			<option value="ipgodate" <%=CHKIIF(datesearch="ipgodate","selected","")%>>�귣�������԰���</option>
			
			<% if locationid <> "" then %>
				<option value="stockipgodate" <%=CHKIIF(datesearch="stockipgodate","selected","")%>>��ǰ�����԰���</option>
			<% end if %>
		</select>
		<input type="text" name="sdate" size="10" maxlength=10 value="<%=sdate%>">
		<a href="javascript:calendarOpen(frm.sdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="edate" size="10" maxlength=10 value="<%=edate%>">
		<a href="javascript:calendarOpen(frm.edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;&nbsp;
		* ��뿩�� : 
		<input type="radio" name="isusing" value="Y" <%=CHKIIF(isusing="Y","checked","")%>>Y&nbsp;
		<input type="radio" name="isusing" value="N" <%=CHKIIF(isusing="N","checked","")%>>N
	</td>
	<td rowspan="4" class="a" align="center" valign="middle">
		<input type="button" onClick="frmsubmit('');" value="�˻�" class="button">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* �귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;&nbsp;
		* ��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8">
		&nbsp;&nbsp;
		* ��ǰ�� : <input type="text" name="itemname" value="<%= itemname %>" size="30">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* ��ǰ���� : 
		<input type="checkbox" name="itemgubun" value="10" <%=CHKIIF(InStr(itemgubun,"10")>0,"checked","")%>>�¶��λ�ǰ(10)&nbsp;
		<input type="checkbox" name="itemgubun" value="90" <%=CHKIIF(InStr(itemgubun,"90")>0,"checked","")%>>������ �����ǰ(90)&nbsp;
		<input type="checkbox" name="itemgubun" value="70" <%=CHKIIF(InStr(itemgubun,"70")>0,"checked","")%>>�Ҹ�ǰ(70)&nbsp;
		<input type="checkbox" name="itemgubun" value="80" <%=CHKIIF(InStr(itemgubun,"80")>0,"checked","")%>>����ǰ(80)&nbsp;
		<input type="checkbox" name="itemgubun" value="60" <%=CHKIIF(InStr(itemgubun,"60")>0,"checked","")%>>���α�(60)&nbsp;
		<input type="checkbox" name="itemgubun" value="00" <%=CHKIIF(InStr(itemgubun,"00")>0,"checked","")%>>������Ի�ǰ(00)&nbsp;
		<input type="checkbox" name="itemgubun" value="95" <%=CHKIIF(InStr(itemgubun,"95")>0,"checked","")%>>���������������Ǹ�(95)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
        &nbsp;&nbsp;        
		* <!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
</table>
<br>
<% If locationid = "" Then %>
	<center><font color="red"><b>�� ShopID(����)�� �����ϼž� �����Ͱ� ��Ÿ���ϴ�.</b></font></center><br>
<% End If %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� �˻��� ����� ��ܺ��� �ִ� 5õ�Ǳ����� �޾����ϴ�.<br>
		<input type="button" onClick="goExcelDown();" value="�����ٿ�" class="button">
	</td>
	<td align="right">	
		<% drawmaechuldatefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>�˻���� : <b><%= FormatNumber(iTotCnt,0) %></b></td>
			<td align="right" valign="bottom">
				���� :
				<select name="tmpsort" id="tmpsort" class="select" style="margin-bottom:3px;" onChange="goSort(this.value);">
					<option value="itemregdate" <%=CHKIIF(sort="itemregdate","selected","")%>>��ǰ�����</option>
					<option value="ipgodate" <%=CHKIIF(sort="ipgodate","selected","")%>>�԰��ϼ�</option>
				</select>			
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>ITEMID</td>
	<td>IMAGE</td>
	<td>BRANDID</td>
	<td>��ǰ��[�ɼǸ�]</td>
	<td>���<br>����</td>
	<td>��ǰ�����</td>
	<td>����������Ʈ��</td>
	<td>�귣��<Br>�����԰���</td>
	
	<% if locationid <> "" then %>
		<td>��ǰ<Br>�����԰���</td>
	<% end if %>

	<td>
		�����
	</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			���԰�
		</td>
	<% end if %>
	
	<td>�Ǹ�<br>����</td>
	
	<td width="60">���</td>
</tr>
<%
If ioffitem.FResultCount > 0 Then

For i=0 To ioffitem.FResultCount -1
%>
<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" align="center">
	<td width=80><%=ioffitem.FItemList(i).Fitemgubun%><%=ioffitem.FItemList(i).Fshopitemid%><%=ioffitem.FItemList(i).Fitemoption%></td>	
	<td width=50>
		<%
			If ioffitem.FItemList(i).Fitemgubun = "10" Then
				Response.Write "<img src='" & ioffitem.FItemList(i).FimageSmall & "' width='50' height='50'>"
			Else
   				If ioffitem.FItemList(i).FOffimgSmall <> "" Then
   					Response.Write "<img src='" & ioffitem.FItemList(i).FOffimgSmall & "' width='50' height='50'>"
   				ElseIf ioffitem.FItemList(i).FOffimgSmall = "" Then
	   				If ioffitem.FItemList(i).FOffimgMain <> "" Then
	   					Response.Write "<img src='" & ioffitem.FItemList(i).FOffimgMain & "' width='50' height='50'>"
	   				ElseIf ioffitem.FItemList(i).FOffimgMain = "" Then
		   				If ioffitem.FItemList(i).FOffimgList <> "" Then
		   					Response.Write "<img src='" & ioffitem.FItemList(i).FOffimgList & "' width='50' height='50'>"
		   				End If
		   			End If
   				End If
	   		End If
		%></td>
	<td><%=ioffitem.FItemList(i).Fmakerid%></td>
	<td align="left">
		<%=ioffitem.FItemList(i).Fshopitemname%>
		
		<% if ioffitem.FItemList(i).Fshopitemoptionname <> "" then %>
		[<%= ioffitem.FItemList(i).Fshopitemoptionname %>]
		<% end if %>
	</td>
	<td width=30><%=ioffitem.FItemList(i).Fisusing%></td>
	<td width=140><%=ioffitem.FItemList(i).Fregdate%></td>
	<td width=140><%=ioffitem.FItemList(i).Fupdt%></td>
	<td width=80><%=ioffitem.FItemList(i).Ffirstipgodate%></td>

	<% if locationid <> "" then %>
		<td width=140>
			<%=ioffitem.FItemList(i).fstockregdate%>
			
			<% if ioffitem.FItemList(i).fstockregdate<>"" then %>
				<!--<Br><a href="javascript:pop_ipgomaechul('<%= locationid %>','<%=ioffitem.FItemList(i).Fitemgubun & Format00(6,ioffitem.FItemList(i).Fshopitemid) & ioffitem.FItemList(i).Fitemoption%>','<%= left(left(ioffitem.FItemList(i).fstockregdate,10),4) %>','<%= mid(left(ioffitem.FItemList(i).fstockregdate,10),6,2) %>','<%= right(left(ioffitem.FItemList(i).fstockregdate,10),2) %>','','','');" onfocus="this.blur()">
				��¥���󼼸��⺸��</a>-->
			<% end if %>
		</td>
	<% end if %>

	<td align="right" bgcolor="#E6B9B8" width=80><%= FormatNumber(ioffitem.FItemList(i).fsellsum,0) %></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right" width=80><%= FormatNumber(ioffitem.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>
	
	<td align="right" width=60><%= FormatNumber(ioffitem.FItemList(i).fitemcnt,0) %></td>
	<td width=60>
		<input type="button" onclick="popShopCurrentStock('<%=ioffitem.FItemList(i).FShopID%>','<%=ioffitem.FItemList(i).Fitemgubun%>','<%=ioffitem.FItemList(i).Fshopitemid%>','<%=ioffitem.FItemList(i).Fitemoption%>');" value="���" class="button">
	</td>
</tr>
<%
Next
%>

<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
	<% if ioffitem.HasPreScroll then %>
		<a href="javascript:frmsubmit('<%= ioffitem.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
		<% if i>ioffitem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:frmsubmit('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ioffitem.HasNextScroll then %>
		<a href="javascript:frmsubmit('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<%
Else
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="center" colspan="20">�˻��� ��ǰ�� �����ϴ�.</td>
</tr>
<% End If %>
</table>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->