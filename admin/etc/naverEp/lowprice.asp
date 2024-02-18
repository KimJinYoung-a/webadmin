<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim makerid, itemid, itemname, orderby, priceCompare, getday, regdate
Dim page, olow, i, suplycash, twentyhigh, sorting
Dim dispCate
page    				= requestCheckvar(request("page"),10)
itemid  				= request("itemid")
makerid					= requestCheckvar(request("makerid"),32)
itemname				= requestCheckvar(request("itemname"),100)
priceCompare			= requestCheckvar(request("priceCompare"),100)
regdate					= requestCheckvar(request("regdate"),10)
orderby					= requestCheckvar(request("orderby"),32)
dispCate 				= requestCheckvar(request("disp"),16)
suplycash				= requestCheckvar(request("suplycash"),4)
twentyhigh				= requestCheckvar(request("twentyhigh"),4)
sorting					= requestCheckvar(request("sorting"),16)

research = requestCheckvar(request("research"),10)
if (research="") and (regdate="") then
    regdate=LEFT(dateadd("d",-1,now()),10)
end if

If page = "" Then page = 1
If itemid <> "" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp) 
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

If sorting <> "" Then orderby = ""
If orderby <> "" Then sorting = ""

SET olow = new epShop
	olow.FCurrPage				= page
	olow.FPageSize				= 20
	olow.FRectMakerid			= makerid
	olow.FRectItemid			= itemid
	olow.FRectItemname			= itemname
	olow.FRectPriceCompare		= priceCompare
	olow.FRectRegdate			= regdate
	olow.FRectCDL				= request("cdl")
	olow.FRectCDM				= request("cdm")
	olow.FRectCDS				= request("cds")
	olow.FRectDispCate			= dispCate
	olow.FRectOrderby			= orderby
	olow.FRectSorting			= sorting
	olow.FRectsuplycash			= suplycash
	olow.FRecttwentyhigh		= twentyhigh
    olow.getNaverLowpriceList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function pop_detail(midx, myrank){
	var popNvD = window.open('/admin/etc/naverEp/pop_lowprice_detail.asp?midx='+midx+'&myrank='+myrank,'notinItem','width=500,height=800,scrollbars=yes,resizable=yes');
	popNvD.focus();
}
function pop_cause(){
	var popCau = window.open('/admin/etc/naverEp/pop_cause.asp','cause','width=800,height=400,scrollbars=yes,resizable=yes');
	popCau.focus();
}
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;&nbsp;
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>&nbsp;&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"><br><br>
		����<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;&nbsp;���� ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		<br><br>
		���԰��� : 
		<select name="suplycash" class="select">
			<option value="">-Choice-</option>
			<option value="high" <%= Chkiif(suplycash = "high", "selected", "") %> >���԰� > ������</option>
			<option value="low" <%= Chkiif(suplycash = "low", "selected", "") %> >���԰� < ������</option>
		</select>
		&nbsp;
		�ǸŰ��� : 
		<select name="twentyhigh" class="select">
			<option value="">-Choice-</option>
			<option value="high" <%= Chkiif(twentyhigh = "high", "selected", "") %> >���������� �ǸŰ��� 20%�̻�</option>
			<option value="low" <%= Chkiif(twentyhigh = "low", "selected", "") %> >���������� �ǸŰ��� 20%�̸�</option>
		</select>
		<br><br>
		��¥ : <input type="text" name="regdate" id="regdate" size="10" value="<%=regdate%>" onClick="jsPopCal('regdate');" style="cursor:hand;">&nbsp;&nbsp;
		<!--
		���̹� �ǸŰ� : 
		<select name="priceCompare" class="select">
			<option value="">-Choice-</option>
			<option value="T" <%= Chkiif(priceCompare = "T", "selected", "") %> >���̹��ǸŰ� > ������</option>
			<option value="N" <%= Chkiif(priceCompare = "N", "selected", "") %> >���̹��ǸŰ� < ������</option>
			<option value="S" <%= Chkiif(priceCompare = "S", "selected", "") %> >���̹��ǸŰ� = ������</option>
		</select>&nbsp;&nbsp;
		-->
		���ı��� : 
		<select name="orderby" class="select">
			<option value="">-Choice-</option>
			<option value="best" <%= Chkiif(orderby = "best", "selected", "") %> >����Ʈ������</option>
			<option value="wish" <%= Chkiif(orderby = "wish", "selected", "") %> >�α� ���ü�</option>
			<!-- <option value="myL" <%= Chkiif(orderby = "myL", "selected", "") %> >�ٹ����ټ�����</option> -->
		</select>
		<input type="hidden" name="sorting" value="<%=sorting%>">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
</p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		1.���̹����ο� �����Ǿ� ������, ���ݺ������� ���� ��Ī�� ��ǰ�� �˻� �����մϴ�.<br>
		2.����Ʈ �߿� <font color="red">��</font>�� ���� �������� �����ϰ��� Ȯ�� �����մϴ�.<br>
		3.�ٹ����� ������ 100�������� Ȯ�� �����ϸ�, �� ������ ��� ������������ ǥ��˴ϴ�.<br>
		4.�ǸŰ����� ���� �򰹼��� �ش� ��ǰ�� �ٹ����ٿ� ���� ��ϵ� ���� ���� ������� ���������Դϴ�.<br>
		5.���� ����Ʈ�� ���� 11�� 10�а濡 ������Ʈ �˴ϴ�.<br>
		6.<input type="button" class="button" value="���ݺ� ���� �Ұ��׸�" onclick="javascrtip:pop_cause();">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(olow.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(olow.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td>��</td>
	<!-- <td width="100">���̹�<br>�ǸŰ�</td> -->
	<td width="100" onClick="jstrSort('sellcash'); return false;" style="cursor:pointer;">
		�ٹ�����<br>�ǸŰ�
		<img src="/images/list_lineup<%=CHKIIF(sorting="sellcashD","_bot","_top")%><%=CHKIIF(instr(sorting,"sellcash")>0,"_on","")%>.png" id="imgsellcash">
	</td>
	<td width="100" onClick="jstrSort('samecashCnt'); return false;" style="cursor:pointer;">
		���ϰ�<br>�ǸŸ�
		<img src="/images/list_lineup<%=CHKIIF(sorting="samecashCntD","_bot","_top")%><%=CHKIIF(instr(sorting,"samecashCnt")>0,"_on","")%>.png" id="imgsamecashCnt">
	</td>
	<td width="100" onClick="jstrSort('lowcash'); return false;" style="cursor:pointer;">
		������
		<img src="/images/list_lineup<%=CHKIIF(sorting="lowcashD","_bot","_top")%><%=CHKIIF(instr(sorting,"lowcash")>0,"_on","")%>.png" id="imglowcash">
	</td>
	<td width="100">Rank2�ǸŰ�</td>
	<td width="100">Rank3�ǸŰ�</td>
	<td width="100" onClick="jstrSort('myrank'); return false;" style="cursor:pointer;">
		�ٹ�����<br>����
		<img src="/images/list_lineup<%=CHKIIF(sorting="myrankD","_bot","_top")%><%=CHKIIF(instr(sorting,"myrank")>0,"_on","")%>.png" id="imgmyrank">
	</td>
	<td width="100" onClick="jstrSort('sellcount'); return false;" style="cursor:pointer;">
		�ǸŰ���
		<img src="/images/list_lineup<%=CHKIIF(sorting="sellcountD","_bot","_top")%><%=CHKIIF(instr(sorting,"sellcount")>0,"_on","")%>.png" id="imgsellcount">
	</td>
	<td width="100" onClick="jstrSort('favcount'); return false;" style="cursor:pointer;">
		���� �򰹼�
		<img src="/images/list_lineup<%=CHKIIF(sorting="favcountD","_bot","_top")%><%=CHKIIF(instr(sorting,"favcount")>0,"_on","")%>.png" id="imgfavcount">
	</td>
	<td width="100" onClick="jstrSort('buycash'); return false;" style="cursor:pointer;">
		���԰�
		<img src="/images/list_lineup<%=CHKIIF(sorting="buycashD","_bot","_top")%><%=CHKIIF(instr(sorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
	</td>
	<td width="100" onClick="jstrSort('margin'); return false;" style="cursor:pointer;">
		������
		<img src="/images/list_lineup<%=CHKIIF(sorting="marginD","_bot","_top")%><%=CHKIIF(instr(sorting,"margin")>0,"_on","")%>.png" id="imgmargin">
	</td>
	<!--
	<td width="100">������Rank</td>
	<td width="100">�ֻ���Rank</td>
	<td width="100">�ְ�</td>
	-->
	<td width="100">������Ʈ<br>��¥</td>
</tr>
<% For i = 0 To olow.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center">
		<a href="<%=wwwURL%>/<%=olow.FItemList(i).FItemID%>" target="_blank"><%= olow.FItemList(i).FItemID %></a><br>
	</td>
	<td align="left"><%= olow.FItemList(i).FMakerid %><br><%= olow.FItemList(i).FItemName %></td>
	<td align="center">
		<% If datediff("d", olow.FItemList(i).FRegdate, now()) < 7 Then  %>
			<input type="button" class="button" value="Ȯ��" onclick="javascrtip:pop_detail('<%= olow.FItemList(i).FIdx %>', '<%=olow.FItemList(i).FMyrank%>')">
		<% Else %>
			�����Ұ�
		<% End If %>
	</td>
	<!-- <td align="center"><%= FormatNumber(olow.FItemList(i).FNaverSellCash,0) %></td> -->
	<td align="center"><%= FormatNumber(olow.FItemList(i).FSellcash,0) %></td>
	<td align="center"><%= olow.FItemList(i).FSamecashCnt %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FLowcash,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FRank2Price,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FRank3Price,0) %></td>
	<td align="center">
	<%
		If olow.FItemList(i).FMyrank = "1000" Then
			response.write "<font color='RED'>������</font>"
		Else
			response.write olow.FItemList(i).FMyrank
		End If
	%>
	</td>
	
	<td align="center"><%= FormatNumber(olow.FItemList(i).FSellcount,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FFavcount,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).Fbuycash,0) %></td>
	<td align="center">
	<%
		If olow.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-olow.FItemList(i).Fbuycash/olow.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	
	<!--
	<td align="center"><%= olow.FItemList(i).FMallmaxrank %></td>
	<td align="center"><%= olow.FItemList(i).FMalllowrank %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FHighcash,0) %></td>
	-->
	<td align="center"><%= LEFT(olow.FItemList(i).FRegdate,10) %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if olow.HasPreScroll then %>
		<a href="javascript:goPage('<%= olow.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + olow.StartScrollPage to olow.FScrollCount + olow.StartScrollPage - 1 %>
    		<% if i>olow.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if olow.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% SET olow = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->