<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ������� �귣�庰����
' History : 2013.01.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/statistic/statisticCls_datamart.asp" -->
<%
dim page,shopid ,yyyymmdd1,yyymmdd2 ,offgubun ,oldlist ,fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,reload
dim i, sum1, sum2, sum3 ,makerid ,datefg , parameter ,CurrencyUnit, CurrencyChar, ExchangeRate 
dim dategubun, vPurchaseType, BanPum, ordertype, FmNum, vOffCateCode, vOffMDUserID
dim totIorgsellprice, totcnt, totrealsellprice, totsuplyprice, totprofit, inc3pl, commcd
	dategubun = request("dategubun")	
	shopid = request("shopid")
	page = request("page")
	if page="" then page=1
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	offgubun = request("offgubun")
	oldlist = request("oldlist")
	makerid = request("makerid")
	datefg = request("datefg")
	vOffCateCode = request("offcatecode")
	vOffMDUserID = request("offmduserid")
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	reload = request("reload")
	BanPum     = request("BanPum")
	ordertype = request("ordertype")
    inc3pl = request("inc3pl")
    commcd = requestCheckVar(request("commcd"),10)
    
if ordertype = "" then ordertype = "totalprice"
if datefg = "" then datefg = "maechul"
if dategubun = "" then dategubun = "G"	
if reload <> "on" and offgubun = "" then offgubun = "95"
	
if (yyyy1="") then
	'fromDate = DateSerial(Cstr( Year(now())), Cstr(Month(now())), Cstr(day(now()))-7 )
	fromDate = DateSerial(Cstr( Year(now())), Cstr(Month(now())), Cstr(day(now())) )
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

'/����
if (C_IS_SHOP) then
	
	'//�������϶�
	if C_IS_OWN_SHOP then
		
		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if		
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

if shopid<>"" then offgubun=""

dim ooffsell
set ooffsell = new cStaticdatamart_list
	ooffsell.FRectShopID = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOffgubun = offgubun
	ooffsell.FRectOldData = oldlist
	ooffsell.frectmakerid = makerid
	ooffsell.frectdatefg = datefg
	ooffsell.frectdategubun = dategubun
	ooffsell.frectoffcatecode = vOffCateCode
	ooffsell.frectoffmduserid = vOffMDUserID
	ooffsell.FCurrPage = page
	ooffsell.Fpagesize=5000
	ooffsell.FRectBrandPurchaseType = vPurchaseType
	ooffsell.FRectBanPum = BanPum
	ooffsell.FRectOrdertype = ordertype
	ooffsell.FRectInc3pl = inc3pl	
	''ooffsell.FRectJungSanGubun = commcd
	ooffsell.FRectCommCd = commcd
	ooffsell.GetBrandSellSumList_datamart

Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

parameter = "menupos="& menupos &"&datefg="& datefg &"&shopid="& shopid &"&offgubun="& offgubun &"&oldlist="& oldlist &"&purchasetype="& vPurchaseType &"&offcatecode="& vOffCateCode &"&offmduserid="& vOffMDUserID &"&BanPum="& BanPum & "&inc3pl=" & inc3pl

sum1 =0
sum2 =0
sum3 =0
totIorgsellprice = 0
totcnt = 0
totrealsellprice = 0
totsuplyprice = 0
totprofit = 0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
	
function pop_category(makerid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var pop_category = window.open('/common/offshop/maechul/statistic/statistic_category_datamart.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&makerid='+makerid+'&<%=parameter%>','pop_category','width=1024,height=768,scrollbars=yes,resizable=yes');
    pop_category.focus();
}

function pop_stock(makerid){
	var pop_stock = window.open('/admin/offshop/jaegolist.asp?makerid='+makerid+'&<%=parameter%>','pop_stock','width=1024,height=768,scrollbars=yes,resizable=yes');
    pop_stock.focus();
}

function pop_detail(makerid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var pop_detail = window.open('/admin/offshop/todayselldetail_datamart.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&makerid='+makerid+'&<%=parameter%>','pop_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
    pop_detail.focus();
}
 
function frmsubmit(){   
	//��¥ ��
	var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
	var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);
    
    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));
                
	if (diffDay > 1095 && frm.oldlist.checked == false){
		alert('3�� ���� �����ʹ� 3������������ȸ �� üũ�ϼž� �մϴ�');
		return;
	}
	 $("#btnSearch").prop("disabled", true);
	frm.submit();
}

function pop_shop(makerid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var pop_shop = window.open('/admin/offshop/brandshopdetail_datamart.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&makerid='+makerid+'&<%=parameter%>','pop_shop','width=1024,height=768,scrollbars=yes,resizable=yes');
    pop_shop.focus();
}

</script>
	
<!-- ǥ ��ܹ� ����-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ :		
				<% drawmaechul_datefg "datefg" ,datefg ,""%>						
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������������ȸ
				&nbsp;&nbsp;&nbsp;	
				<%
				'����/������
				if (C_IS_SHOP) then
				%>	
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='frmsubmit();'" %>
					<% end if %>
				<% else %>
					* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='frmsubmit();'" %>
				<% end if %>
					
				<p>
				* ��ǰ���� :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* ī�װ� : <% SelectBoxBrandCategory "offcatecode", vOffCateCode %>
				&nbsp;&nbsp;
				* �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
				&nbsp;&nbsp;
				* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* �������δ��MD : <% drawSelectBoxCoWorker_OnOff "offmduserid", vOffMDUserID, "off" %>
				&nbsp;&nbsp;
				<% if (C_IS_Maker_Upche) then %>
					* �귣�� : <%= makerid %><br>
					<input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>	
				<% end if %>
				<p>
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>		
	             <%if shopid <> "" then%>
	            &nbsp;&nbsp;* ���걸��
	            
	            <% call drawSelectBoxOFFJungsanCommCD("commcd",commcd) %>
    	            <% if (FALSE) then '' ��� ����.%>
        	            <select name="sJGb" class="select">
        	            	<option value="">��ü</option>
        	            	<%sbOptJungSanGubun sJungSangubun%>
        	            </select>
        	        <% end if %>
	            <%end if%>		
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" name="btnSearch" id="btnSearch" class="button_s" value="�˻�" onclick="frmsubmit();">
	</td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		�� �˻� �Ⱓ�� ��� �������ϴ�. �˻� ��ư�� ������, �ƹ� ������ ����δٰ�, �ٽ� �˻���ư�� Ŭ������ ������.
    </td>
    <td align="right">
		<input type="radio" name="dategubun" value="G" <% if dategubun="G" then response.write " checked" %> onclick="frmsubmit();">�Ⱓ�����
		<input type="radio" name="dategubun" value="M" <% if dategubun="M" then response.write " checked" %> onclick="frmsubmit();">�������
		/ ����:
		<% drawordertype "ordertype" ,ordertype ," onchange='frmsubmit();'" ,"B"  %>			
    </td>        
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ooffsell.FResultCount %></b> �� �ִ� 5000�� ���� �˻��˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<% if dategubun = "M" then %>
		<td>��¥</td>
	<% end if %>
	
	<td>�귣��ID</td> 
	<td>���걸��</td> 
	<td>��������</td>
	<td>��ǰ����</td>
	<% if (NOT C_InspectorUser) then %>
	<td>�Һ��ڰ�[��ǰ]</td>
    <% end if %>
	<td>�����</td>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td>�����Ѿ�[��ǰ]</td>
		<td><b>�������</b></td>
		<td>������</td>
	<% end if %>
	
	<td>���</td>
</tr>
<%
if ooffsell.FresultCount > 0 then
	
for i=0 to ooffsell.FresultCount-1

totIorgsellprice = totIorgsellprice + ooffsell.FItemList(i).fIorgsellprice
totcnt = totcnt + ooffsell.FItemList(i).FCount
totrealsellprice = totrealsellprice + ooffsell.FItemList(i).FSum
totsuplyprice = totsuplyprice + ooffsell.FItemList(i).fsuplyprice
totprofit = totprofit + ooffsell.FItemList(i).fprofit

sum1 = sum1 + ooffsell.FItemList(i).FSum

if ooffsell.FItemList(i).FJcomm_cd="B012" then
	sum2 = sum2 + ooffsell.FItemList(i).FSum
else
	sum3 = sum3 + ooffsell.FItemList(i).FSum
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<% if dategubun = "M" then %>
		<td>
			<%= ooffsell.FItemList(i).fIXyyyymmdd %>
		</td>
	<% end if %>	
		
	<% if ooffsell.FItemList(i).FJcomm_cd="B012" then %>
		<td><b><font color="#3333CC"><a href="javascript:PopBrandInfoEdit('<%= ooffsell.FItemList(i).FMakerid %>')"><%= ooffsell.FItemList(i).FMakerid %></a></font></b></td>
	<% else %>
		<td><a href="javascript:PopBrandInfoEdit('<%= ooffsell.FItemList(i).FMakerid %>')"><%= ooffsell.FItemList(i).FMakerid %></a></td>
	<% end if %>
	<td><% if ooffsell.FItemList(i).FJcomm_cd="B012" then %><b><font color="#3333CC"><%=getmwdiv_beasongdivname(ooffsell.FItemList(i).FJcomm_cd)%></font></b>
		<% else%>
		<%=getmwdiv_beasongdivname(ooffsell.FItemList(i).FJcomm_cd)%>
		<% end if%>
	</td> 
	<td><%=ooffsell.FItemList(i).fpurchasetypename%></td>
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
    <% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fIorgsellprice,0) %></td>
    <% end if %>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).FSum,0) %></td>
	
	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) %></td>
		<td align="right"><b><%= FormatNumber(ooffsell.FItemList(i).fprofit,0) %></b></td>
		<td align="right">
			<% if ooffsell.FItemList(i).fsuplyprice > 0 and ooffsell.FItemList(i).FSum > 0 then %>
				<%= FormatNumber(100-ooffsell.FItemList(i).fsuplyprice/ooffsell.FItemList(i).FSum*100,0) %>%
			<% else %>
				0%
			<% end if %>
		</td>
	<% end if %>

	<td width=250>
		<% if dategubun = "G" then %>
			<input type="button" onclick="pop_shop('<%= ooffsell.FItemList(i).FMakerid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>');" value="���庰" class="button">
			<input type="button" onclick="pop_detail('<%= ooffsell.FItemList(i).FMakerid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>');" value="��ǰ��" class="button">
			<input type="button" onclick="pop_category('<%= ooffsell.FItemList(i).FMakerid %>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>');" value="ī�װ���" class="button">
		<% elseif dategubun = "M" then %>
			<input type="button" onclick="pop_shop('<%= ooffsell.FItemList(i).FMakerid %>','<%= left(ooffsell.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ooffsell.FItemList(i).fIXyyyymmdd,6,2) %>','01','<%= left(ooffsell.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ooffsell.FItemList(i).fIXyyyymmdd,6,2) %>','31');" value="���庰" class="button">
			<input type="button" onclick="pop_detail('<%= ooffsell.FItemList(i).FMakerid %>','<%= left(ooffsell.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ooffsell.FItemList(i).fIXyyyymmdd,6,2) %>','01','<%= left(ooffsell.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ooffsell.FItemList(i).fIXyyyymmdd,6,2) %>','31');" value="��ǰ��" class="button">
			<input type="button" onclick="pop_category('<%= ooffsell.FItemList(i).FMakerid %>','<%= left(ooffsell.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ooffsell.FItemList(i).fIXyyyymmdd,6,2) %>','01','<%= left(ooffsell.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ooffsell.FItemList(i).fIXyyyymmdd,6,2) %>','31');" value="ī�װ���" class="button">
		<% end if %>

		<% if not(C_IS_SHOP) then %>
			<!--<input type="button" onclick="pop_stock('<%'= ooffsell.FItemList(i).FMakerid %>');" value="���" class="button">-->
		<% end if %>
	</td>
</tr>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<% if dategubun = "M" then %>
		<td colspan=2>�Ѱ�</td>
	<% else %>
		<td>�հ�</td>
	<% end if %>
	<td></td> 
	<td></td> 
	<td><%= FormatNumber(totcnt,0) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(totIorgsellprice,0) %></td>
    <% end if %>
	<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>	
		<td align="right"><%= FormatNumber(totsuplyprice,0) %></td>
		<td align="right"><b><%= FormatNumber(totprofit,0) %></b></td>
		<td></td>
	<% end if %>
	
	<td align="right">
		<b><font color="#3333CC">��üƯ�� : </font></b><%= FormatNumber(sum2,0) %>
		<br>�Ϲ� : <%= FormatNumber(sum3,0) %>
		<br>Total : <%= FormatNumber(sum1,0) %>
	</td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="16">�˻� ����� �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->