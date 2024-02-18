<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%
dim itembarcode
dim itemgubun,itemid,itemoption
dim actType, makerid

itembarcode  = requestCheckVar(request("itembarcode"),20)
itemgubun 	 = requestCheckVar(request("itemgubun"),2)
itemid		 = requestCheckVar(request("itemid"),10)
itemoption	 = requestCheckVar(request("itemoption"),4)
actType      = requestCheckVar(request("actType"),10)

if (C_IS_Maker_Upche) then
    makerid = session("ssBctId")
end if

if (Len(itembarcode)=12) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,6))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(6,itemid) + itemoption
elseif (Len(itembarcode)=14) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,8))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(8,itemid) + itemoption
elseif (Len(itembarcode)<>0) and (itemid<>"") then
	itemgubun = "10"
	itemid = itembarcode
	if (itemoption="") then itemoption  = "0000"
elseif (Len(itembarcode)>6) then
    '''���ڵ��ΰ�� �˻��� ��ǰ�ڵ� ������.
    call fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
else
    itemgubun = "10"
    if (itemid="") then itemid = itembarcode
    if (itemoption="") then itemoption  = "0000"

    if (itemid>=1000000) then
        itembarcode = itemgubun + Format00(8,itemid) + itemoption
    else
        itembarcode = itemgubun + Format00(6,itemid) + itemoption
    end if
end if




dim oitem
set oitem = new CItem
oitem.FRectMakerid= makerid
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
end if


dim otodayerritem
set otodayerritem = new CSummaryItemStock
otodayerritem.FRectItemGubun = itemgubun
otodayerritem.FRectItemID =  itemid
otodayerritem.FRectItemOption =  itemoption
if itemid<>"" then
    otodayerritem.GetTodayErrItem
end if

dim difftime
if (osummarystock.FResultcount>0) then
    difftime = ABS(datediff("h",osummarystock.FOneItem.Flastupdate,now()))
end if

dim i
dim IsVaildCode
IsVaildCode = False
if (oitemoption.FResultCount>0) then
    for i=0 to oitemoption.FResultCount-1
        if (oitemoption.FITemList(i).FItemOption=itemoption) then
            IsVaildCode = (oitem.FResultCount>0)
            exit For
        end if
    next
else
    IsVaildCode = (oitem.FResultCount>0) and (itemoption="0000")
end if


dim ErrMsg, sqlStr
dim SelectedOptionStr
dim stockReipgoDate

sqlStr = "select top 1 stockReipgoDate from [db_item].[dbo].tbl_item_option_Stock"
sqlStr = sqlStr & " where itemgubun='" & itemgubun & "'"
sqlStr = sqlStr & " and itemid=" & itemid
sqlStr = sqlStr & " and itemoption='" & itemoption & "'"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    stockReipgoDate = rsget("stockReipgoDate")
end if
rsget.Close

%>
<script language='javascript'>

var delCliked = false;

function ClearVal(comp){
    comp.value='';
    delCliked = true;
}

//���԰��� ����.
function SaveStockReipgoDate(frm){
    var nowDate = "<%= Left(Now(),10) %>";

    if ((frm.stockReipgoDate.value.length<1)&&(!delCliked)){
        alert('���԰� �������� ������ �ּ���.');
        return;
    }



    if (frm.stockReipgoDate.value.length<1){
        if (confirm('���԰� �������� ���� �˴ϴ�. ��� �Ͻðڽ��ϱ�?')){
            frm.mode.value = "stockreipgodate";
            frm.submit();
        }
    }else{
        if (frm.stockReipgoDate.value<nowDate){
            alert('���԰� �������� ���� ���ķ� �����ϼ���.');
            return;
        }

        if (confirm('���԰� �������� ���� �Ͻðڽ��ϱ�?')){
            frm.mode.value = "stockreipgodate";
            frm.submit();
        }
    }
}

//���� ����
function SaveDanjongSoldOut(frm){
    //���������� �����Ǹ��ΰ�츸 ������
	if (frm.isEditValid.value==""){
		alert('���� �Ǹ��� ��츸 ������,����ǰ��, MDǰ���� ���� �� �� �ֽ��ϴ�.');
		//frm.limityn[0].focus();
		return;
	}

    if (confirm('���� ǰ�� ó�� �����Ͻðڽ��ϱ�?')){
        frm.mode.value = "danjong";
        frm.submit();
    }
}

//MDǰ�� ����
function SaveMdSoldOut(frm){
    //���������� �����Ǹ��ΰ�츸 ������
	if (frm.isEditValid.value==""){
		alert('���� �Ǹ��� ��츸 ������,����ǰ��, MDǰ���� ���� �� �� �ֽ��ϴ�.');
		//frm.limityn[0].focus();
		return;
	}

    if (confirm('MD ǰ�� ó�� �����Ͻðڽ��ϱ�?')){
        frm.mode.value = "mssoldout";
        frm.submit();
    }
}



function GetOnLoad(){
	<% if Not IsVaildCode then %>
    	<% if oitemoption.FResultCount>0 then %>
    	alert('��ǰ�ڵ尡 ��Ȯ���� �ʽ��ϴ�. �ɼ� ������ ��˻� �ϼ���.');
    	<% else %>
    	alert('��ǰ�ڵ尡 ��Ȯ���� �ʽ��ϴ�. ��˻� �ϼ���.');
    	<% end if %>
	document.frm.itembarcode.select();
	document.frm.itembarcode.focus();
	<% end if %>
}
window.onload=GetOnLoad;

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="actType" value="<%= actType %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;
				        <strong>
				        <% if (actType="D") then %>
				        ��������
				        <% elseif (actType="R") then %>
				        ���԰����� ����
				        <% end if %>
				        </strong></font>
				    </td>
				    <td align="right">
						��ǰ�ڵ�:
						<% if (C_IS_Maker_Upche) then %>
						<input type=text class="text_ro" ReadOnly name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 UTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
						<% else %>
						<input type=text class="text" name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 UTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
						<% end if %>
						&nbsp;

						<% if oitemoption.FResultCount>0 then %>

						<select class="select" name="itemoption">
						<option value="0000">----
						<% for i=0 to oitemoption.FResultCount-1 %>
						<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
						<% next %>
						</select>
						<% end if %>

        				<input type="button" class="button" value="�˻�" onclick="document.frm.submit();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->
	</form>
</table>

<p>
<% if (oitem.FResultCount<1) then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td align="center">[�˻� ����� �����ϴ�.]</td>
    </tr>
</table>
<% else %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60">��ǰ�ڵ�</td>
      	<td width="300">
      		10<Strong><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></Strong><%= itemoption %>
      	</td>
      	<td width="60"></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td>�Ǹſ���</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FSellyn,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ��</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td>��뿩��</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FIsUsing,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�ǸŰ�</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- ���ο���/�������뿩�� -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     ����
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> ����
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td>��������</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
			<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
			<font color="#CC3333">MDǰ��</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">�Ͻ�ǰ��</font>
			<% else %>
			������
			<% end if %>
		</td>
    </tr>
     <% if oitemoption.FResultCount>1 then %>
	    <!-- �ɼ��� �ִ°�� -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
	    	<% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
	    	<% SelectedOptionStr = "<font color=blue>[" & oitemoption.FITemList(i).FOptionName & "]</font>" %>
	    	<tr bgcolor="#FFFFFF">
	    		<td>�ɼǸ�</td>
		      	<td><%= oitemoption.FITemList(i).FOptionName %> (<%= fnColor(oitemoption.FITemList(i).FOptIsUsing,"yn") %>)</td>
		      	<td>��������</td>
		      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% end if %>
		<% next %>
	<% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td>�ɼǸ�</td>
	      	<td>-</td>
	      	<td>��������</td>
	      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>

</table>
<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frm2 method=post action="itemdanjong_process.asp">
    <input type=hidden name=mode value="">
    <input type=hidden name=itemgubun value="<%= itemgubun %>">
    <input type=hidden name=itemid value="<%= itemid %>">
    <input type=hidden name=itemoption value="<%= itemoption %>">

    <% if (actType="D") and (oitem.FOneItem.FLimityn<>"Y") then %>
    <input type=hidden name=isEditValid value="">
    <% else %>
    <input type=hidden name=isEditValid value="on">
    <% end if %>

<!-- �ǽð� ������Ʈ ��
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td colspan="16" align=right>����������Ʈ : <%= osummarystock.FOneItem.Flastupdate %> </td>
    </tr>
-->
    <tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">���԰�/��ǰ</td>
    	<td width="50">���Ǹ�/��ǰ</td>
		<td width="50">�����/��ǰ</td>
		<td width="50">��Ÿ���/��ǰ</td>
		<td width="50">CS<br>���/��ǰ</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">�ý���<br>�����</td>
		<td width="50">�Ѻҷ�</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">�ý���<br>��ȿ���</td>
		<td width="50">�ѽǻ�<br>����</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">�ǻ�<br>���</td>
		<td width="50">ON��ǰ<br>�غ�</td>
		<td width="50">OFF��ǰ<br>�غ�</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">����ľ�<br>���</td>
		<td width="50">ON����<br>�Ϸ�</td>
		<td width="50">ON�ֹ�<br>����</td>
		<td width="50">OFF�ֹ�<br>����</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ftotipgono %></td>
    	<td rowspan="2"><%= -1*osummarystock.FOneItem.Ftotsellno %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrcsno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Ftotsysstock %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrbaditemno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Favailsysstock %></td>
    	<td rowspan="2" ><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Frealstock %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv5 %></td>
    	<td><%= osummarystock.FOneItem.Foffconfirmno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.GetCheckStockNo %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv4 %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv2 %></td>
    	<td><%= osummarystock.FOneItem.Foffjupno %></td>

    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="2"><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
    	<td colspan="3"><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>

    </tr>


</table>
<p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
<tr>
    <td align="center">
        <% if (actType="D") then %>
            <% if (oitemoption.FResultCount>0) then %>
            (���� ������ ��� �ɼ� ���� ���� �����˴ϴ�)
            <br>
            <% end if %>
            <input type="button" class="button" value="����ǰ������" onclick="SaveDanjongSoldOut(frm2)">
            <% if Not (C_IS_Maker_Upche) then %>
            &nbsp;&nbsp;
            <input type="button" class="button" value="MDǰ������" onclick="SaveMdSoldOut(frm2)">
            <% end if %>
        <% elseif (actType="R") then %>
          <% if ErrMsg<>"" then %>
            <%= ErrMsg %>
          <% else %>
          <%= SelectedOptionStr %> ���԰� ������ : <input type="text" class="text" name="stockReipgoDate" size="10" value="<%= stockReipgoDate %>" readOnly>
          <a href="javascript:calendarOpen(frm2.stockReipgoDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
  	      <a href="javascript:ClearVal(frm2.stockReipgoDate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
          <br><br>
          <input type="button" class="button" value="���԰� ������ ����" onclick="SaveStockReipgoDate(frm2)">
          <% end if %>
        <% end if %>
    </td>
</tr>
</form>
</table>



<% end if %>
<%
set otodayerritem = Nothing
set osummarystock = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->