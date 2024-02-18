<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim oChart, vArr1, i, j, k
Dim vSDate, vEDate, onlynv, itemid, grpdate
Dim page : page = requestCheckvar(request("page"),10)
Dim itemoption : itemoption = requestCheckvar(request("itemoption"),4)
vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
onlynv = requestCheckvar(request("onlynv"),10)
grpdate = requestCheckvar(request("grpdate"),10)
itemid = requestCheckvar(request("itemid"),10)


if (page="") then page=1

If vSDate = "" Then
	vSDate = LEFT(dateadd("d",-15,Date()),10)
End If

If vEDate = "" Then
	vEDate = LEFT(dateadd("d",-1,Date()),10)
End If

dim vMakerid
dim oitem
set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo

	if (oitem.FResultCount>0) then
		vMakerid = oitem.FOneItem.FMakerid
	end if
end if


dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

Dim isOptionExists : isOptionExists = (oitemoption.FResultCount>0)


SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectitemid = itemid
	oChart.FRectOnlyNvShop = onlynv
    oChart.FRectGrpDate = grpdate

    if (itemid<>"") then
	    vArr1 = oChart.fngetOneItemCpnSellTrend_Datamart()
    end if

Dim iSellStDate 
if (oitem.FResultCount>0) then
	iSellStDate=oitem.FOneItem.FSellStdate
	if isNULL(iSellStDate) then 
		iSellStDate=""
	else
		iSellStDate=LEFT(iSellStDate,10)
	end if
end if

dim imgURL, iSellyn, iLimityn, iLimitNo, isellStr, iLimitStr
dim icpnStr, imaycouponVal, imaycouponBuyVal
dim isActivCpn
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>


<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type="text/javascript">
$(function() {
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
		//	CAL_End.args.min = date;
		//	CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	
});

function goSearch(){
	if($("#sdate").val() == ""){
		alert("�������� �Է��ϼ���");	
		return false;
	}
	if($("#edate").val()== ""){
		alert("�������� �Է��ϼ���");	
		return false;
	}
	document.frm1.submit();
}


function pop_couponView(v){
    var pwin = window.open('/admin/shopmaster/itemcouponlist.asp?menupos=786&research=on&page=1&iSerachType=1&sSearchTxt='+v,'popOutMallEtcLink','width=1200,height=700,scrollbars=yes,resizable=yes');
    pwin.focus();
}

function showoption(comp){
	var ioptdiv = document.getElementById("idoptlist");
	if (comp.value=="�ɼ�ǥ��"){
		ioptdiv.style.display = "block";
		comp.value="�ɼǼ���";
	}else{
		ioptdiv.style.display = "none";
		comp.value="�ɼ�ǥ��";
	}
}
</script>


<body>
<form name="frm1" method="get" >
<input type="hidden" name="page">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	   

	��¥(�ֹ���) : 
	<input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
	<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    ~
    <input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
	<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
	</span>		
	
	&nbsp;&nbsp;
    ��ǰ�ڵ� : 
    <input type="text" name="itemid" value="<%=itemid%>" size="7" maxlength="10">
	&nbsp;&nbsp;
    <input type="checkbox" name="onlynv" <%=CHKIIF(onlynv="on","checked","")%> >rdsite NvShop ���⸸ ����
    &nbsp;&nbsp;
    <input type="checkbox" name="grpdate" <%=CHKIIF(grpdate="on","checked","")%> >�ⰣGrouping
    </td>
    <td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>

<p>
<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=5 width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
      	<td width="35%">
      		10 <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
      	</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ǸŽ�����</td>
      	<td><%=iSellStDate%></td>
		<td align="right">
		<input type="button" value="��ǰ���� web"  onClick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=ItemID%>','_viewitem','');">

		<input type="button" value="���ݺ���LOG"  onClick="window.open('/admin/etc/extsitejungsan_check.asp?itemid=<%=ItemID%>','_itemlog','');">

		<input type="button" value="���޼���LOG" onClick="window.open('/admin/etc/outmall/index.asp?research=on&menupos=1742&makerid=<%=vMakerid%>','_outmallsellyn','');">
		</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">�Ǹſ���</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
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
      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">�Ͻ�ǰ��</font>
			<% else %>
			������
			<% end if %>
		</td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">�ɼ�</td>
        <td><%=CHKIIF(isOptionExists,"��"&oitemoption.FResultCount,"-")%>
		<% if (isOptionExists) then %>
		&nbsp;&nbsp;<input type="button" value="�ɼ�ǥ��" onClick="showoption(this);">
		<% end if %>
		</td>
        <td bgcolor="<%= adminColor("tabletop") %>">��������</td>
        <td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
        <td>���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
    </tr>
</table>
<div id="idoptlist" name="idoptlist" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
    <% if oitemoption.FResultCount>1 then %>
	    <% for i=0 to oitemoption.FResultCount -1 %>
            <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
                <% if (i=0) then %><td rowspan="<%=oitemoption.FResultCount%>" width="110" ></td><% end if %>
		      	<td bgcolor="<%= adminColor("tabletop") %>" width="60"><font color="#AAAAAA">�ɼǸ� :</font></td>
		      	<td width="35%"><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
		      	<td bgcolor="<%= adminColor("tabletop") %>" width="60"><font color="#AAAAAA">�������� : </font></td>
		      	<td ><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
		      	<td width="120">���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>
				<tr bgcolor="<%=CHKIIF(oitemoption.FITemList(i).Fitemoption=itemoption,"#EEEEEE","#FFFFFF")%>">
					<% if (i=0) then %><td rowspan="<%=oitemoption.FResultCount%>" width="110" ></td><% end if %>
					<td width="60">�ɼǸ�</td>
					<td width="35%"><%= oitemoption.FITemList(i).FOptionName %></td>
					<td width="60">��������</td>
					<td ><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
					<td width="120">���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
				</tr>
		    <% end if %>
	    <% next %>
    <% end if %>
</table>
</div>
<% end if %>
<%
SET oitem = Nothing
SET oitemoption = Nothing
%>
<p />

* �ش�Ⱓ ���� ��� ���������, �ֹ��� ����, �ڻ�����ش��.
<p>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<% If isArray(vArr1) Then %>
    
	<tr bgcolor="#DDDDDD" align="center">
		<td width="<%=CHKIIF(grpdate<>"","200","100")%>">�ֹ���</td>
        <td width="50">�����ڵ�</td>
		<td width="50">���ֹ���</td>
		<td width="50">�Ǹż���</td>
		<td width="60">�����Ѿ���</td>
		<td width="70">���<br>�ǸŰ�</td>
		<td width="70">���<br>�����Ѿ�</td>
		<td width="70">���<br>����������</td>
		<td width="70">���<br>��������</td>

        <td width="50">��������</td>
		<td>����������</td>
        <td>����������</td>
        <td>��������</td>
        <td>��������</td>
        <td>������</td>
        <td>����ø��԰�</td>
	</tr>
	<% For i = 0 To UBound(vArr1,2) %>
	<%
    isActivCpn = false
    if NOT isNULL(vArr1(17,i)) then
        if (vArr1(15,i)=7) and NOT isNULL(vArr1(17,i)) and (vArr1(10,i)>now()) then isActivCpn = true
    end if
    %>
	<tr  bgcolor="#FFFFFF" align="center">
		<td><%=vArr1(0,i)%></td>
        <td>
        <% if (isActivCpn) then %>
        <a href="#" onClick="pop_couponView('<%=vArr1(1,i)%>');return false;"><strong><%=vArr1(1,i)%></strong></a>
        <% else %>
        <a href="#" onClick="pop_couponView('<%=vArr1(1,i)%>');return false;"><%=vArr1(1,i)%></a>
        </td>
        <% end if %>
        

		<td align="center"><%=FormatNumber(vArr1(3,i),0)%></td>
		<td align="center"><%=FormatNumber(vArr1(4,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(5,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(6,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(7,i),0)%></td>
		<td align="center">
			<% if vArr1(6,i)<>0 then %>
			<font color=green><%=100-CLNG(vArr1(7,i)/vArr1(6,i)*100)%></font> %
			<% end if %>
		</td>
		<td align="right"><%=FormatNumber(vArr1(8,i),0)%></td>

		<td align="center"><%=CHKIIF(vArr1(11,i)="V","NV ","�Ϲ� ")%></td>
		<td align="center"><%=vArr1(9,i)%></td>
        <td align="center"><%=vArr1(10,i)%></td>
		<td align="center">
            <%=CHKIIF(vArr1(15,i)=9,"����",CHKIIF(vArr1(10,i)<now(),"����",""))%>
            <% if (isActivCpn) then %><font color="red">Ȱ��</font><% end if %>
        </td>
        <td>
            <%=FormatNumber(vArr1(12,i),0)%>
        <% if (vArr1(13,i)="1") then  %> 
            % ����
        <% else %>
            �� ����
        <% end if %>
        </td>
        <td><%=vArr1(14,i)%></td>
        <td align="right">
        <% if not isNULL(vArr1(18,i)) then %>
        <%=FormatNumber(vArr1(18,i),0)%>
        <% end if %>
        </td>

	</tr>
	<% next %>
	
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td>
			�˻������ �����ϴ�.
		</td>
	</tr>
	<% end if %>
	</table>

<%
SET oChart = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
