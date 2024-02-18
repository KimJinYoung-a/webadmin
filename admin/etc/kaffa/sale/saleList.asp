<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kaffa/itemsalecls.asp"-->
<%
Dim clsSale, page, i
Dim salestatus, research
Dim selType, sTxt, selDate, iSD, iED

page        = requestCheckvar(request("page"),10)
salestatus  = requestCheckvar(request("salestatus"),10)
research    = requestCheckvar(request("research"),10)
selType     = requestCheckvar(request("selType"),10)
sTxt        = requestCheckvar(request("sTxt"),20)
selDate     = requestCheckvar(request("selDate"),10)
iSD         = requestCheckvar(request("iSD"),10)
iED         = requestCheckvar(request("iED"),10)

If page = "" Then page = 1
if (research="") and (salestatus="") then salestatus="V"
Set clsSale = new CSale
	clsSale.FCurrPage	= page
	clsSale.FPageSize	= 30
	clsSale.FRectSaleStatus = salestatus
	clsSale.FRectSelType = selType
	clsSale.FRectSelText = sTxt
	clsSale.FRectSelDate = selDate
	clsSale.FRectSelStartDt = iSD
	clsSale.FRectSelEndDt = iED
	clsSale.fnGetSaleList
%>
<script language="javascript">
function goPage(p){
    document.frmSearch.page.value=p;
    document.frmSearch.submit();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function jsMod(scode){
	location.href = "saleReg.asp?discountKey="+scode+"&menupos=<%=menupos%>";
}
function jsGoURL(ival){
	location.href = "saleItemReg.asp?discountKey="+ival+"&menupos=<%=menupos%>";
}
function saleRegByTen(){
    var popwin = window.open('saleRegByTen.asp?regyn=N','saleRegByTen','scrollbars=yes,resizable=yes,width=1100,height=700');
    popwin.focus();
}
</script>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frmSearch" method="get"  >
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
		<td align="left">
		<select name="selType">
		<option value="1" <%=CHKIIF(selType="1","selected","")%> >�����ڵ�</option>
		<option value="2" <%=CHKIIF(selType="2","selected","")%> >��ǰ�ڵ�</option>
		<option value="3" <%=CHKIIF(selType="3","selected","")%> >���θ�</option>
		<option value="4" <%=CHKIIF(selType="4","selected","")%> >TEN�����ڵ�</option>
		</select>
		<input type="text" name="sTxt" value="<%=sTxt%>" size="10" maxlength="20">
		&nbsp;�Ⱓ:
	<select name="selDate">
		<option value="S" <%=CHKIIF(selDate="S","selected","")%> >������ ����</option>
		<option value="E" <%=CHKIIF(selDate="E","selected","")%> >������ ����</option>
		</select>
		<input type="text" size="10" name="iSD" value="<%=iSD%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
		~ <input type="text" size="10" name="iED" value="<%=iED%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
		&nbsp;����:

    	<select name="salestatus" class="select" onChange='javascript:document.frmSearch.submit();'>
    	<option value="">����</option>
    	<option value="0" <%=CHKIIF(salestatus="0","selected","")%> >��ϴ��</option>
        <option value="6" <%=CHKIIF(salestatus="6","selected","")%> >���ο���</option>
    	<option value="7" <%=CHKIIF(salestatus="7","selected","")%> >������</option>
    	<option value="9" <%=CHKIIF(salestatus="9","selected","")%> >����</option>
        <option value="V" <%=CHKIIF(salestatus="V","selected","")%> >��������</option>
    	</select>

		</td>
		<td  width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
		</td>
	</tr>
	</form>
</table>
<!---- /�˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="40" valign="bottom">
        <td align="left">
        	<input type="button" value="���ε��" class="button" onclick="javascript:location.href='saleReg.asp?menupos=<%=menupos%>';" >
        	&nbsp;&nbsp;
        	<input type="button" value="TEN�������� ���" class="button" onclick="saleRegByTen();" >
	    </td>
	    <td align="right"></td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="12">�˻���� : <b><%= FormatNumber(clsSale.FTotalCount,0) %></b>&nbsp;&nbsp;������ : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(clsSale.FTotalPage,0) %></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>�����ڵ�</td>
    	<td>TEN�����ڵ�</td>
    	<td>���θ�</td>
    	<td>������</td>
    	<td>���԰�����</td>
    	<td>������</td>
    	<td>������</td>
    	<td>����</td>
    	<td>ó��</td>
    	<td>�����</td>
    </tr>
<% For i = 0 To clsSale.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).FDiscountKey%></a></td>
    	<td><%=CHKIIF(clsSale.FItemList(i).FPromotionType=0,"",clsSale.FItemList(i).FPromotionType)%></td>
    	<td align="left">&nbsp;<a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).FDiscountTitle%></a></td>
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).FDiscountPro%>%</a></td>
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).getRuleStr%></a></td>
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).FStDT%></a></td>
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).FEdDT%></a></td>
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).getSaleStateStr%></a></td>
    	<td>
    		<input type="button" value="��ǰ(<%=clsSale.FItemList(i).FDiscountitem_cnt%>)" class="button" onClick="javascript:jsGoURL('<%=clsSale.FItemList(i).FDiscountKey%>')">
   		</td>
    	<td><a href="javascript:jsMod(<%=clsSale.FItemList(i).FDiscountKey%>)" title="���� ��������"><%=clsSale.FItemList(i).FRegdate%></a></td>
    </tr>
<% Next %>
	<tr height="20">
		<td colspan="11" align="center" bgcolor="#FFFFFF">
		<% If clsSale.HasPreScroll Then %>
			<a href="javascript:goPage('<%= clsSale.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i=0 + clsSale.StartScrollPage To clsSale.FScrollCount + clsSale.StartScrollPage - 1 %>
			<% If i>clsSale.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If clsSale.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
		</td>
	</tr>
</table>
<%
Set clsSale = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->