<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim isEditEnable : isEditEnable = FALSE
if (session("ssBctID")="icommang" or session("ssBctID")="hrkang97" or session("ssBctID")="kjy8517" or session("ssBctID")="coolhas") then
    isEditEnable = TRUE
end if

dim itemid : itemid=requestCheckVar(request("itemid"), 10)
dim ordsch : ordsch=requestCheckVar(request("ordsch"), 10)
dim itemcost : itemcost=requestCheckVar(request("itemcost"), 10)
dim buycash : buycash=requestCheckVar(request("buycash"), 10)
dim exceptbcash : exceptbcash=requestCheckVar(request("exceptbcash"), 10)

dim itemoption : itemoption=requestCheckVar(request("itemoption"), 10)
dim mallsellcash : mallsellcash=requestCheckVar(request("mallsellcash"), 10)
dim orderserial : orderserial=requestCheckVar(request("orderserial"), 16)

dim sitename : sitename=requestCheckVar(request("sitename"), 32)

if (mallsellcash="") then mallsellcash=0
if (itemoption="") then itemoption="0000"
dim sqlStr, arrRows
''�ɼ�/�߰��� ���� �˻�

dim iOptionCNT : iOptionCNT=0
dim iOptAddPrcCNT : iOptAddPrcCNT=0
dim currSellprice : currSellprice=0
dim currOrgPrice  : currOrgPrice=0
dim currBuycash   : currBuycash=0
dim currOrgSuplycash : currOrgSuplycash=0
dim optaddprice     : optaddprice=0
dim optaddbuyprice  : optaddbuyprice=0

dim maybuycash : maybuycash=0

sqlStr = "select count(*) as optcnt,sum(CASE WHEN isNULL(optaddprice,0)<>0 then 1 ELSE 0 end) as optAddPrcCNT from db_item.dbo.tbl_item_option WITH(NOLOCK) where itemid="&itemid&" and isusing='Y'"
if (itemid<>"") and (isNumeric(itemid)) then
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
    if  not rsget.EOF  then
        iOptionCNT = rsget("optcnt")
        iOptAddPrcCNT = rsget("optAddPrcCNT")
    end if
    rsget.close
end if

if (iOptAddPrcCNT<>0) and (itemoption<>"0000") then
    sqlStr = "select isNULL(optaddprice,0) as optaddprice, isNULL(optaddbuyprice,0) as optaddbuyprice from db_item.dbo.tbl_item_option WITH(NOLOCK) where itemid="&itemid&" and itemoption='"&itemoption&"'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
    if  not rsget.EOF  then
        optaddprice = rsget("optaddprice")
        optaddbuyprice = rsget("optaddbuyprice")
    end if
    rsget.close

end if

'if (iOptAddPrcCNT>0) then
    sqlStr = "select sellcash,orgprice, buycash, orgsuplycash from db_item.dbo.tbl_item WITH(NOLOCK)  where itemid="&itemid
    if (itemid<>"") and (isNumeric(itemid)) then
        rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        if  not rsget.EOF  then
            currSellprice = rsget("sellcash")
            currOrgPrice = rsget("orgprice")

            currBuycash = rsget("buycash")
            currOrgSuplycash = rsget("orgsuplycash")
        end if
        rsget.close
    end if

'end if

sqlStr = "select top 20 * from db_log.dbo.tbl_iteminfo_history WITH(NOLOCK) where itemid="&itemid&" order by regdate desc"
if (itemid<>"") and (isNumeric(itemid)) then
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
    if  not rsget.EOF  then
        arrRows = rsget.getRows()
    end if
    rsget.close
end if

dim itemEdtCnt : itemEdtCnt = 0
dim i
if IsArray(arrRows) then
    itemEdtCnt = UBound(arrRows,2) +1
end if

dim arrRows2, ordCnt : ordCnt =0

if (ordsch="on") and (itemid<>"") and (isNumeric(itemid)) and (itemcost<>"") and (isNumeric(itemcost)) and (buycash<>"") and (isNumeric(buycash)) then
    sqlStr = " select m.sitename,convert(varchar(19),d.beasongdate,121) as beasongdate,d.orderserial,d.itemid,d.itemoption,d.idx,d.makerid,d.itemno,d.itemcost,d.cancelyn,d.itemoptionname"
    sqlStr = sqlStr & " ,d.buycash, d.buycashcouponNotApplied, m.cancelyn, d.itemcouponidx, m.ipkumdiv, m.linkorderserial"
    sqlStr = sqlStr & " ,d.reducedprice"
    sqlStr = sqlStr & " ,d.dlvfinishdt, d.jungsanfixdate"
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_detail  d WITH(NOLOCK) "
    sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m WITH(NOLOCK) "
    sqlStr = sqlStr & "	on m.orderserial=d.orderserial"
    sqlStr = sqlStr & " where itemid="&itemid&""
    if (itemcost<>"") then
        sqlStr = sqlStr & " and itemcost="&itemcost&""
    end if
    if (buycash<>"") and (exceptbcash="on") then
        sqlStr = sqlStr & " and buycash<>"&buycash&""
    end if
    if (orderserial<>"") then
        sqlStr = sqlStr & " and m.orderserial='"&orderserial&"'"&VBCRLF
    end if
    if (sitename<>"") then
        sqlStr = sqlStr & " and m.sitename='"&sitename&"'"&VBCRLF
    end if
    sqlStr = sqlStr & " order by d.idx desc"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
    if  not rsget.EOF  then
        arrRows2 = rsget.getRows()
    end if
    rsget.close
end if

if IsArray(arrRows2) then
    ordCnt = UBound(arrRows2,2) +1
end if


dim isValidEdit
dim mayJDate : mayJDate = LEFT(dateadd("d",-3,NOW()),7)&"-01"   ''���� 2�ϱ����� ����.
if (session("ssBctID")="icommang") or (session("ssBctID")="hrkang97") or (session("ssBctID")="coolhas")  then
    mayJDate = LEFT(dateadd("d",-9,NOW()),7)&"-01"
end if

%>
<script>
function bsearch(){
    document.bfrm.itemid.value=document.frm.itemid.value;
    document.bfrm.submit();
}

function research(itemcost,buycash){
    document.frm.ordsch.value="on";
    document.frm.itemcost.value=itemcost;
    document.frm.buycash.value=buycash;
    document.frm.submit();
}

function edtORDDTL(orderserial,itemid,didx,itemcost,buycash){
    var iparam = "orderserial="+orderserial+"&itemid="+itemid+"&didx="+didx+"&itemcost="+itemcost+"&buycash="+buycash;

    var popwin = window.open('extsitejungsan_edit.asp?'+iparam,'edtORDDTL','');
    popwin.focus();
}

function edtORDDTL2(orderserial,itemid,didx,itemcost,buycash){
    var iparam = "orderserial="+orderserial+"&itemid="+itemid+"&didx="+didx+"&itemcost="+itemcost+"&buycash="+buycash+"&onlybuycash=on";

    var popwin = window.open('extsitejungsan_edit.asp?'+iparam,'edtORDDTL','');
    popwin.focus();
}
</script>
<!-- �˻� ���� -->
<table width="90%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td  width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			��ǰ��ȣ : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=14 onkeyPress="if (event.keyCode == 13){ bsearch(); return false;}"><input type="button" class="button_s" value="��˻�" onClick="bsearch()">
			<input type="hidden" name="ordsch" value="">
			<input type="hidden" name="itemcost" value="">
			<input type="hidden" name="buycash" value="">

			<% if (iOptAddPrcCNT<>0) then %>
			    �ɼ��ڵ� : <input type="text" class="text" name="itemoption" value="<%= itemoption %>" size=4 onkeyPress="if (event.keyCode == 13){ document.frm.submit(); return false;}">
			<% end if %>
			&nbsp;&nbsp;&nbsp;&nbsp;
			�����ǸŰ� : <input type="text" class="text" name="mallsellcash" value="<%=mallsellcash%>" size=14 onkeyPress="if (event.keyCode == 13){ document.frm.submit(); return false;}">

			<% if (ordsch="on") and (itemid<>"") and (isNumeric(itemid)) and (itemcost<>"") and (isNumeric(itemcost)) and (buycash<>"") and (isNumeric(buycash)) then %>
			&nbsp;<input type="checkbox" name="exceptbcash" <%=CHKIIF(exceptbcash="on","checked","")%> onClick="research('<%=itemcost%>','<%=buycash%>');"> �ٸ�������
		    <% end if %>

		    �ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size=11 >
		</td>

		<td  width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button_s" value="�˻�" onclick="document.frm.submit()">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="3">
	        &nbsp;&nbsp;&nbsp;&nbsp;

	        <% if (currOrgPrice<>0) then %>
			    &nbsp;���Һ�:<%=FormatNumber(currOrgPrice,0)%>
                /
                <%=FormatNumber(currOrgSuplycash,0)%>

                (<%=100-currOrgSuplycash/currOrgPrice*100%> %)
			<% end if %>
			&nbsp;&nbsp;
			<% if (currSellprice<>0) then %>
			    &nbsp;���ǸŰ�:<%=FormatNumber(currSellprice,0)%>
                /
                <%=FormatNumber(currBuycash,0)%>

                (<%=100-currBuycash/currSellprice*100%> %)
			<% end if %>
	        &nbsp;&nbsp;
			<% if (iOptionCNT<>0) then %>
			    &nbsp;�ɼǼ�:<%=iOptionCNT%>
			<% end if %>
			&nbsp;&nbsp;
			<% if (iOptAddPrcCNT<>0) then %>
			    &nbsp;�ɼ��߰��ݾ׼�:<strong><%=iOptAddPrcCNT%></strong>
			<% end if %>
			&nbsp;&nbsp;
			<% if (optaddprice<>0) then %>
			    &nbsp;�ɼ��߰��ݾ�:<strong><%=optaddprice%>/<%=optaddbuyprice%></strong>

                (<%=100-optaddbuyprice/optaddprice*100%> %)
			<% end if %>
	    </td>
	</tr>
	</form>
	<form name="bfrm" method="get" >
	<input type="hidden" name="itemid" value="">
	</form>
</table>
<p>
*��ǰ ���� ���� (��ǰ������ 2017/10/19 ����)
<p>

<table width="90%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="100">������</td>
    <td width="80">��ǰ�ڵ�</td>
    <td width="80">�ɼ�</td>
    <td width="90">�ǸŰ�</td>
    <td width="90">���԰�</td>
    <td width="70">�Ǹſ���</td>
    <td width="70">��������</td>
    <td width="70">���Ա���</td>
    <td width="70">��۱���</td>
    <td width="123">��ǰ��</td>
    <td width="70">�ɼǼ�</td>
    <td width="70">��ǰ����</td>
    <td width="70">���</td>
    <td width="100">����</td>
</tr>
<% for i=0 to itemEdtCnt-1 %>
<tr align="center" bgcolor="#FFFFFF" height="25">
   <td><%= arrRows(0,i) %></td>
   <td><%= arrRows(1,i) %></td>
   <td><%= arrRows(2,i) %></td>
   <td align="right"><%= FormatNumber(arrRows(3,i),0) %></td>
   <td align="right"><%= FormatNumber(arrRows(4,i),0) %></td>
   <td><%= arrRows(5,i) %></td>
   <td><%= arrRows(6,i) %></td>
   <td><%= arrRows(8,i) %></td>
   <td><%= arrRows(9,i) %></td>
   <td><%= arrRows(10,i) %></td>
   <td><%= arrRows(11,i) %></td>
   <td><%= arrRows(12,i) %></td>
   <td>
        <% if (mallsellcash<>0) then %>
            <% if (mallsellcash-arrRows(3,i)=optaddprice) then %>
                <% maybuycash=arrRows(4,i)+optaddbuyprice %>
                <a onClick="research('<%=mallsellcash%>','<%=maybuycash%>')" style="cursor:pointer"><%=mallsellcash%></a>
            <% end if %>
        <% end if %>
   </td>
   <td><input type="button" value="����" onClick="research('<%=arrRows(3,i)%>','<%=arrRows(4,i)%>')">
</tr>
<% next %>
</table>

<p>
*�ֹ� ���� (<%=ordCnt%>)��
<p>
<table width="90%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="100">�Ǹ�ó</td>
    <td width="80">�����</td>
    <td width="80">�ֹ���ȣ</td>
    <td width="90">��ǰ�ڵ�</td>
    <td width="90">�ɼ��ڵ�</td>
    <td width="70">IDX</td>
    <td width="70">�귣��ID</td>
    <td width="70">����</td>
    <td width="70">�ǸŰ�</td>
    <td width="70">���ǸŰ�</td>
    <td width="70">��ҿ���</td>
    <td width="123">�ɼǸ�</td>
    <td width="100">���԰�</td>
    <td width="100">�����̸��԰�</td>
    <td width="60">��ǰ����</td>
    <td width="60">����</td>
    <td width="80">������</td>
    <% if (isEditEnable) then %>
    <td width="100">����</td>
    <% end if %>
</tr>
<% for i=0 to ordCnt-1 %>
<%
isValidEdit = FALSE
if NOT (isNULL(arrRows2(1,i)) or isNULL(arrRows2(19,i)))  then
    'isValidEdit = (LEFT(arrRows2(1,i),10)>=mayJDate)
    isValidEdit = FALSE

    if NOT (isNULL(arrRows2(19,i))) then
        isValidEdit = (LEFT(arrRows2(19,i),10)>=mayJDate)
    end if
else
    isValidEdit = TRUE

    if NOT isNULL(arrRows2(1,i)) then  ''2020-01-01 �������� jugsanFixdate �� ����.
        if (LEFT(arrRows2(1,i),10)<"2020-01-01") then
            isValidEdit = FALSE
        end if
    end if
end if




if (NOT isValidEdit) then isValidEdit="disabled"
%>
<tr align="center" bgcolor="#FFFFFF" height="25">
   <td><%= arrRows2(0,i) %></td>
   <td><%= arrRows2(1,i) %><% if (arrRows2(15,i)="1") then %>�ֹ�����<% end if %></td>
   <td><%= arrRows2(2,i) %>
    <% if NOT isNULL(arrRows2(16,i)) then %>
    <br><strong><%= arrRows2(16,i) %></strong>
    <% end if %>
   </td>
   <td><%= arrRows2(3,i) %></td>
   <td><%= arrRows2(4,i) %></td>
   <td><%= arrRows2(5,i) %></td>
   <td><%= arrRows2(6,i) %></td>
   <td><%= arrRows2(7,i) %></td>
   <td align="right"><%= FormatNumber(arrRows2(8,i),0) %></td>
   <td align="right"><%= FormatNumber(arrRows2(17,i),0) %></td>
   <% if (arrRows2(15,i)="1") then %>
   <td bgcolor='#AAAA77'><%= arrRows2(13,i) %>,<%= arrRows2(9,i) %></td>
   <% else %>
   <td <%= CHKIIF((arrRows2(13,i)="Y" or arrRows2(9,i)="Y"),"bgcolor='#AA77AA'","")%>><%= arrRows2(13,i) %>,<%= arrRows2(9,i) %></td>
   <% end if %>
   <td><%= arrRows2(10,i) %></td>
   <td align="right" <%= CHKIIF(CLNG(buycash)<>(arrRows2(11,i)),"bgcolor='#77AAAA'","")%>>
    <% if (arrRows2(11,i)<>CLNG(arrRows2(11,i))) then %>
    <%= FormatNumber(arrRows2(11,i),2) %>
    <% else %>
    <%= FormatNumber(arrRows2(11,i),0) %>
    <% end if %>
   </td>
   <td align="right" <%= CHKIIF(CLNG(buycash)<>(arrRows2(12,i)),"bgcolor='#77AAAA'","")%>>
    <% if (arrRows2(12,i)<>CLNG(arrRows2(12,i))) then %>
    <%= FormatNumber(arrRows2(12,i),2) %>
    <% else %>
    <%= FormatNumber(arrRows2(12,i),0) %>
    <% end if %>
   </td>
   <td><%=arrRows2(14,i)%></td>
   <td>
   <% if arrRows2(17,i)<>0 then %>
   <%=100-CLNG(arrRows2(11,i)/arrRows2(17,i)*100*100)/100%>
   <% end if %>
   </td>
   <td><%= arrRows2(19,i) %></td>
   <% if (isEditEnable) then %>
   <td>
        <% if (CLNG(buycash)<>(arrRows2(11,i))) then '(arrRows2(0,i)<>"10x10")and %>

            <% if (CLNG(arrRows2(11,i))<>CLNG(arrRows2(12,i))) then %>
                <input type="button" value="����(OnlyB)" onClick="edtORDDTL2('<%= arrRows2(2,i) %>','<%= arrRows2(3,i) %>','<%= arrRows2(5,i) %>','<%=itemcost%>','<%=buycash%>')" <%=isValidEdit%>>
            <% end if %>
        <input type="button" value="����" onClick="edtORDDTL('<%= arrRows2(2,i) %>','<%= arrRows2(3,i) %>','<%= arrRows2(5,i) %>','<%=itemcost%>','<%=buycash%>')" <%=isValidEdit%>>
        <% end if %>
   </td>
   <% end if %>
</tr>
<% next %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
