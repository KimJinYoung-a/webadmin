<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim oChart, vArr1, i, j, k
Dim vSDate, vEDate, vChannel, itopn, makerid, mwdiv, exceptsoldout, onlynv, ordercntOver, nocpn
Dim page : page = requestCheckvar(request("page"),10)
vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
itopn = requestCheckvar(request("itopn"),10)
makerid = requestCheckvar(request("makerid"),32)
mwdiv = requestCheckvar(request("mwdiv"),10)
exceptsoldout = requestCheckvar(request("exceptsoldout"),10)
onlynv = requestCheckvar(request("onlynv"),10)
ordercntOver = requestCheckvar(request("ordercntOver"),10)
nocpn = requestCheckvar(request("nocpn"),10)

if (itopn="") then itopn=100
if (page="") then page=1

If vSDate = "" Then
	vSDate = LEFT(dateadd("d",-15,Date()),10)
End If

If vEDate = "" Then
	vEDate = LEFT(dateadd("d",-1,Date()),10)
End If

if (ordercntOver="") then ordercntOver=1
if not isNumeric(ordercntOver) then ordercntOver=1

Dim iszozimtype : iszozimtype = 1



SET oChart = new CChart
	oChart.FPageSize = itopn
	oChart.FCurrPage = page
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectMakerid = makerid
	oChart.FRectMwdiv = mwdiv
	oChart.FRectExceptSoldout = exceptsoldout
	oChart.FRectOnlyNvShop = onlynv
	oChart.FRectordercntOver = ordercntOver
	oChart.FRectnocpn = nocpn

	vArr1 = oChart.fngetItemCpnBestSell_Datamart()



dim imgURL, iSellyn, iLimityn, iLimitNo, isellStr, iLimitStr
dim icpnStr, imaycouponVal, imaycouponBuyVal
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

function popCpnSellTrend(itemid) {
	var popItemSellGraph = window.open("popOneitemCpnTrend.asp?itemid="+itemid+"&startdate=<%=vSDate%>&enddate=<%=vEDate%>&onlynv=<%=onlynv%>","popOneitemCpnTrend","width=1400, height=1000,resizable=yes, scrollbars=yes");
	popItemSellGraph.focus();
}

function popItemTrend(itemid){
	var popwin = window.open("/admin/maechul/itemTrend.asp?itemid="+itemid,"popItemTrend","width=1400, height=1000,resizable=yes, scrollbars=yes");
	popwin.focus();
}

//�귣�� �߼�
function popBrandSellGraph(makerid,startdate,enddate) {
	var popBrandSellGraph = window.open("/admin/dataanalysis/chart/sellbybrand.asp?pvalue="+makerid+"&startdate="+startdate+"&enddate="+enddate,"popBrandSellGraph","width=1700, height=800,resizable=yes, scrollbars=yes");
	popBrandSellGraph.focus();
}

function pop_couponView(v){
    var pwin = window.open('/admin/shopmaster/itemcouponlist.asp?menupos=786&research=on&page=1&iSerachType=1&sSearchTxt='+v,'popOutMallEtcLink','width=1200,height=700,scrollbars=yes,resizable=yes');
    pwin.focus();
}

function goPage(page) {
	var frm = document.frm1;
	frm.page.value = page;
	frm.submit();
}
</script>


<body>
<form name="frm1" method="get" >
<input type="hidden" name="page">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	
    

	��¥(�ֹ���) : 
	<input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
	<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    ~
    <input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
	<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
	</span>		
	
	
	&nbsp;&nbsp;
   
    
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>
<tr bgcolor="#F4F4F4">
	<td>
	�Ǽ�
	<select name="itopn">
		<option value="100" <%=CHKIIF(itopn="100","selected","") %> >100</option>
		<option value="200" <%=CHKIIF(itopn="200","selected","") %> >200</option>
		<option value="300" <%=CHKIIF(itopn="300","selected","") %> >300</option>
    </select>
	&nbsp;
	�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
	&nbsp;
		���Ա��� : <% Call drawSelectBoxMWU("mwdiv",mwdiv) %>
		&nbsp;
		<input type="checkbox" name="exceptsoldout" <%=CHKIIF(exceptsoldout="on","checked","")%>>ǰ����ǰ ����
		&nbsp;
		<input type="checkbox" name="onlynv" <%=CHKIIF(onlynv="on","checked","")%>>rdsite NvShop ���⸸ ����

	&nbsp;
	�ֹ��Ǽ� <input type="text" name="ordercntOver" value="<%=ordercntOver%>" size="2" maxlength="6"> �� �̻�
	
	&nbsp;
	�ֹ��Ǽ� <input type="checkbox" name="nocpn" <%=CHKIIF(nocpn="on","checked","")%>>�����������»�ǰ��


	</td>
</tr>

</table>
</form>
<br />
* �ش�Ⱓ ���� ��� ���������, �ֹ��� ����, �ڻ�����ش��.
<p>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<% If isArray(vArr1) Then %>
	<tr bgcolor="#FFFFFF" >
		<td colspan="21">
			�˻���� : <b><%= FormatNumber(oChart.FTotalCount,0) %></b>
            &nbsp;
            ������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oChart.FTotalPage,0) %></b>
		</td>
	</tr>
	<tr bgcolor="#DDDDDD" align="center">
		<td width="60">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td width="50">���ֹ���</td>
		<td width="50">�Ǹż���</td>
		<td width="60">�����Ѿ���</td>
		<td width="70">���<br>�ǸŰ�</td>
		<td width="70">���<br>�����Ѿ�</td>
		<td width="70">���<br>����������</td>
		<td width="70">���<br>��������</td>

		<td width="90">�귣��</td>
		<td>��ǰ��</td>
		<td>����������</td>
		<td width="80">(��)�ǸŰ�</td>
		<td width="80">(��)���԰�</td>
		<td width="40">����</td>
		<td>����<br>����</td>
		<td>�Ǹ�<br>����</td>
		<td>����<br>����</td>
		<td>������������</td>
		<td>�����Ǹ�<br>����</td>
		<td>��ǰ<br>�߼�</td>
	</tr>
	<% For i = 0 To UBound(vArr1,2) %>
	<%
	imgURL = vArr1(16,i)
	if ((Not IsNULL(imgURL)) and (imgURL<>"")) then 
		imgURL = "<img src='"&webImgUrl & "/image/small/" + GetImageSubFolderByItemid(vArr1(0,i)) + "/"  + vArr1(16,i)&"'>"
	else
		imgURL = ""
	end if

	isellStr =""
	iLimitStr=""
	icpnStr =""

	iSellyn = vArr1(10,i)
	iLimityn = vArr1(11,i)
	iLimitNo = vArr1(12,i)-vArr1(13,i)
	if (iLimitNo<1) then iLimitNo=0
		

	if (iSellyn<>"Y") then isellStr="<strong><font color='#FF0000'>ǰ��</font></strong>"
	if (iSellyn="S") then isellStr="<strong><font color='#CC3333'>�Ͻ�ǰ��</font></strong>"
	if (iLimityn="Y") then iLimitStr="<font color='#3333CC'>����<br>("&iLimitNo&")</font>"

	imaycouponVal    = vArr1(14,i)
	imaycouponBuyVal = vArr1(15,i)

	if (NOT isNULL(vArr1(21,i))) then
		icpnStr = CHKIIF(vArr1(17,i)="V","NV ","�Ϲ� ")
		icpnStr = icpnStr&"<a href=""#"" onClick=""pop_couponView('"&vArr1(21,i)&"');return false;"">"&vArr1(21,i)&"</a>"
		icpnStr = icpnStr&"<br><font color=green>"&vArr1(18,i)&"</font>"

		if (vArr1(19,i)="1") then  ''%
			imaycouponVal = imaycouponVal-CLNG(imaycouponVal*vArr1(20,i)/100)
		elseif (vArr1(19,i)="2") then  ''�ݾ�
			imaycouponVal = imaycouponVal-vArr1(20,i)
		end if

		if (vArr1(22,i)<>0) then imaycouponBuyVal=vArr1(22,i)
	end if
	%>
	<tr  bgcolor="#FFFFFF" align="center">
		<td><%=vArr1(0,i)%></td>
		<td><%=imgURL%></td>
		<td align="center"><%=FormatNumber(vArr1(1,i),0)%></td>
		<td align="center"><%=FormatNumber(vArr1(2,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(3,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(4,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(5,i),0)%></td>
		<td align="center">
			<% if vArr1(4,i)<>0 then %>
			<font color=green><%=100-CLNG(vArr1(5,i)/vArr1(4,i)*100)%></font> %
			<% end if %>
		</td>
		<td align="right"><%=FormatNumber(vArr1(6,i),0)%></td>
		<td align="center"><%=vArr1(7,i)%></td>
		<td align="left"><%=vArr1(8,i)%></td>
		<td align="center">
			<% if NOT isNULL(vArr1(25,i)) then %>
			<%=vArr1(25,i)%><br>
			<%=vArr1(26,i)%><br>
			<% if vArr1(27,i)<vArr1(28,i) then %>
				<strong><%=FormatNumber(vArr1(27,i),0)%></strong><br>
				<%=FormatNumber(vArr1(28,i),0)%><br>
			<% else %>
				<%=FormatNumber(vArr1(27,i),0)%><br>
				<%=FormatNumber(vArr1(28,i),0)%><br>
			<% end if %>
			<%=LEFT(vArr1(29,i),13)%>
			<% end if %>
		</td>
		<td align="right">
			<% if vArr1(23,i)>vArr1(14,i) then %>
			<%=FormatNumber(vArr1(23,i),0)%><br>
			<font color="red"><%=FormatNumber(vArr1(14,i),0)%></font>
			<% else %>
			<%=FormatNumber(vArr1(14,i),0)%>
			<% end if %>
			
			<% if (imaycouponVal<>vArr1(14,i)) then %>
			<br><font color=green><%=FormatNumber(imaycouponVal,0)%></font>
			<% end if %>
		</td>
		<td align="right">
			<% if vArr1(23,i)>vArr1(14,i) then %>
			<%=FormatNumber(vArr1(24,i),0)%><br>
			<font color="red"><%=FormatNumber(vArr1(15,i),0)%></font>
			<% else %>
			<%=FormatNumber(vArr1(15,i),0)%>
			<% end if %>

			<% if (imaycouponVal<>vArr1(14,i)) then %>
			<br><font color=green><%=FormatNumber(imaycouponBuyVal,0)%></font>
			<% end if %>
		</td>
		<td>
			<% if vArr1(23,i)>vArr1(14,i) then %>
				<% if (vArr1(23,i)<>0) then %>
				<%=100-CLNG(vArr1(24,i)/vArr1(23,i)*100) %><br>
				<% end if %>	
				<% if (vArr1(14,i)<>0) then %>
				<font color="red"><%=100-CLNG(vArr1(15,i)/vArr1(14,i)*100) %></font>
				<% end if %>
			<% else %>
				<% if (vArr1(14,i)<>0) then %>
				<%=100-CLNG(vArr1(15,i)/vArr1(14,i)*100) %>
				<% end if %>
			<% end if%>
			<% if (imaycouponVal<>vArr1(14,i)) then %>
			<br>
			<% if (imaycouponVal<>0) then %>
			<font color=green><%=100-CLNG(imaycouponBuyVal/imaycouponVal*100) %></font>
			<% end if %>
			<% end if %>

		</td>
		<td><%=vArr1(9,i)%></td>
		<td><%=isellStr%></td>
		<td><%=iLimitStr%></td>
		<td><%=icpnStr%></td>
		<td><img src="/images/icon_search.jpg" onClick="popCpnSellTrend('<%=vArr1(0,i)%>');" style="cursor:pointer"></td>
		<td>
			<img src="/images/icon_search.jpg" onClick="popItemTrend('<%=vArr1(0,i)%>');" style="cursor:pointer">
		</td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="21" align="center" bgcolor="#FFFFFF">
	        <% if oChart.HasPreScroll then %>
			<a href="javascript:goPage('<%= oChart.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oChart.StartScrollPage to oChart.FScrollCount + oChart.StartScrollPage - 1 %>
	    		<% if i>oChart.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oChart.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
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
