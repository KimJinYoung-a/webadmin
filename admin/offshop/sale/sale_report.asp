<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ���
' History : 2012.10.23 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/salereport_cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
Dim clsSale, shopid, page ,i, yyyy1, mm1, dd1, yyyy2, mm2, dd2, fromDate, toDate
Dim iStartPage, iEndPage, iTotalPage, ix, iPerCnt, strParm, inc3pl
Dim iSerachType, sSearchTxt, sBrand, datefg, isStatus
	iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
	datefg     		= requestCheckVar(Request("datefg"),1)		'�˻��� ����
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	isStatus		= requestCheckVar(Request("salestatus"),4)	'���� ����
	shopid		= requestCheckVar(Request("shopid"),32)		'����
	page		= requestCheckVar(Request("page"),10)		'����
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "S"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-90)
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

if page = "" then page = 1
	 
'�˻��κ��� ��ȣ�� �޾ƾߵȴٸ� ���ڸ� ���� 
if iSerachType="1" or iSerachType="2" then 		
	sSearchTxt = getNumeric(sSearchTxt)
end if
	
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

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if	

dim oReport
set oReport = new Csalereport_list
	oReport.FPageSize = 500
	oReport.FCurrPage = page
	oReport.FSearchType = iSerachType    
	oReport.FSearchTxt  = sSearchTxt     
	oReport.FBrand		= sBrand     	
	oReport.frectdatefg   = datefg
	oReport.frectevt_startdate		= fromDate     	
	oReport.frectevt_enddate		= toDate     			
	oReport.FSStatus	= isStatus
 	oReport.frectshopid = 	shopid
	oReport.FRectInc3pl = inc3pl
	
	'/��� ������
	oReport.getsale_sum

Dim arrsalemargin, arrsalestatus , arrsaleshopmargin
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arrsalemargin = fnSetCommonCodeArr_off("salemargin",False)
	arrsaleshopmargin = fnSetCommonCodeArr_off("shopsalemargin",False)
	arrsalestatus= fnSetCommonCodeArr_off("salestatus",False)	
%>

<script language="javascript">

	function submitfrm(page){
		frmSearch.page.value=page;
		frmSearch.submit();
	}

	//�󼼺���
	function pop_detail(SType,sale_code,shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){	
		 var pop_detail = window.open('/admin/offshop/sale/sale_report_detail.asp?SType='+SType+'&sale_code='+sale_code+'&shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&inc3pl=<%=inc3pl%>&menupos=<%=menupos%>','pop_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
		 pop_detail.focus();
	}
	
</script>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
<form name="frmSearch" method="get"  action="" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ:
		<select name="datefg">
			<option value="S" <%if Cstr(datefg) = "S" THEN %>selected<%END IF%>>������ ����</option>
			<option value="E" <%if Cstr(datefg) = "E" THEN %>selected<%END IF%>>������ ����</option>
		</select>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3" ,"","" %>
			<% end if %>
		<% else %>
			<% if not(C_IS_Maker_Upche) then %> 
				* ���� : <% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3" ,"","" %>
			<% else %>
				* ���� : <% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3" ,"","" %>
			<% end if %>
		<% end if %>
		<p>
		<select name="selType">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>�����ڵ�</option>
			<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>���θ�</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		&nbsp;&nbsp;	
		* ����:<% sbGetOptCommonCodeArr_off "salestatus", isStatus, True, False,"onChange='javascript:document.frmSearch.submit();'"%>
        &nbsp;&nbsp;
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>		
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="submitfrm('');">
	</td>
</tr>
	
</form>
</table>
<!---- /�˻� ---->

<Br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		<font color="red">[�ʵ�]</font> �� ��� �����ʹ� �Ϸ翡 �ѹ� ������ ������Ʈ �˴ϴ�.
    </td>
    <td align="right">

    </td>        
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">	
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		�˻���� : <b><%=oReport.FResultCount%></b>�� �� 500�� ���� �˻�����
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����<br>�ڵ�</td>
	<td>���θ�<br>�������</td>
	<td>���Ը���<br>�����޸���</td>
	<td>������<br>������</td>		 	
	<td>����<br>�ϼ�</td>
	<td>����</td>
	<td>������</td>
	<td>����<br>����Ʈ</td>
	<td>���<br>��ǰ��</td>
	<!--<td>�Ǹž�</td>-->
	<td>�����<br>�����</td>	
	<td>�����</td>
	<!--<td>�ֹ�<br>�Ǽ�</td>-->
	<td>�Ǹ�<br>����</td>	
	<td>���</td>
</tr>
<%
dim totsellprice ,totrealsellprice ,totselljumuncnt ,totsellCnt ,datelen
	totsellprice = 0
	totrealsellprice = 0
	totselljumuncnt = 0
	totsellCnt = 0
	
if oReport.FresultCount>0 then
	
For i = 0 To oReport.FResultCount - 1

totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
totselljumuncnt = totselljumuncnt + oReport.FItemList(i).ftotselljumuncnt
totsellCnt = totsellCnt + oReport.FItemList(i).ftotsellCnt

'/�̺�Ʈ�ϼ�
datelen = ""
datelen = datediff("d",oReport.FItemList(i).fsale_startdate,oReport.FItemList(i).fsale_enddate)
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td width=60>
		<%= oReport.FItemList(i).fsale_code %>
	</td>	
	<td>
		<%= oReport.FItemList(i).fsale_name %><br><%= oReport.FItemList(i).fshopname %>
	</td>
	<td>
		<%=fnGetCommCodeArrDesc_off(arrsalemargin,oReport.FItemList(i).fsale_margin )%>
		<br><%=fnGetCommCodeArrDesc_off(arrsaleshopmargin,oReport.FItemList(i).fsale_shopmargin )%>
	</td>    
	<td width=80>
		<%= oReport.FItemList(i).fsale_startdate %><br><%= oReport.FItemList(i).fsale_enddate %>
	</td>
		
	<td width=50><%= datelen %></td>
	<td width=80>
		<%
		'/����
		IF oReport.FItemList(i).fsale_status = 6 THEN
		%>
			<font color="blue"><%=fnGetCommCodeArrDesc_off(arrsalestatus,oReport.FItemList(i).fsale_status)%></font>
		<%
		'/����
		elseIF oReport.FItemList(i).fsale_status = 8 THEN
		%>
			<font color="gray"><%=fnGetCommCodeArrDesc_off(arrsalestatus,oReport.FItemList(i).fsale_status)%></font>
		<%
		'/���¿�û , �����û
		elseIF oReport.FItemList(i).fsale_status = 7 or oReport.FItemList(i).fsale_status = 9 THEN
		%>
			<font color="red"><%=fnGetCommCodeArrDesc_off(arrsalestatus,oReport.FItemList(i).fsale_status)%></font>		
		<% else %>
			<%=fnGetCommCodeArrDesc_off(arrsalestatus,oReport.FItemList(i).fsale_status)%>
		<% end if %>
	</td>
	<td width=50>
		<%= oReport.FItemList(i).fsale_rate %> %
	</td>
	<td width=50>
		<%= oReport.FItemList(i).fpoint_rate %> %
	</td>
	<td width=50 align="right">
		<%= FormatNumber(oReport.FItemList(i).fsaleitem_cnt,0) %>
	</td>
	<!--<td width=80 align="right">
		<%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %>
	</td>-->
	<td width=80 align="right">
		<% if datelen <> "" and datelen <> 0 then %>
			<%= FormatNumber(oReport.FItemList(i).ftotrealsellprice / datelen,0) %>
		<% else %>
			0	
		<% end if %>
	</td>		
	<td width=80 align="right" bgcolor="#E6B9B8">
		<%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %>
	</td>	
	<!--<td width=80 align="right"><%'= FormatNumber(oReport.FItemList(i).ftotselljumuncnt,0) %></td>-->
	<td width=50 align="right">
		<%= FormatNumber(oReport.FItemList(i).ftotsellCnt,0) %>
	</td>	
	<td width=170>
		<input type="button" onclick="pop_detail('D','<%= oReport.FItemList(i).fsale_code %>','<%= oReport.FItemList(i).fshopid %>','<%= left(oReport.FItemList(i).fsale_startdate,4) %>','<%= mid(oReport.FItemList(i).fsale_startdate,6,2) %>','<%= right(oReport.FItemList(i).fsale_startdate,2) %>','<%= left(oReport.FItemList(i).fsale_enddate,4) %>','<%= mid(oReport.FItemList(i).fsale_enddate,6,2) %>','<%= right(oReport.FItemList(i).fsale_enddate,2) %>');" value="��¥��" class="button">
		<input type="button" onclick="pop_detail('I','<%= oReport.FItemList(i).fsale_code %>','<%= oReport.FItemList(i).fshopid %>','<%= left(oReport.FItemList(i).fsale_startdate,4) %>','<%= mid(oReport.FItemList(i).fsale_startdate,6,2) %>','<%= right(oReport.FItemList(i).fsale_startdate,2) %>','<%= left(oReport.FItemList(i).fsale_enddate,4) %>','<%= mid(oReport.FItemList(i).fsale_enddate,6,2) %>','<%= right(oReport.FItemList(i).fsale_enddate,2) %>');" value="��ǰ��" class="button">
		<input type="button" onclick="pop_detail('B','<%= oReport.FItemList(i).fsale_code %>','<%= oReport.FItemList(i).fshopid %>','<%= left(oReport.FItemList(i).fsale_startdate,4) %>','<%= mid(oReport.FItemList(i).fsale_startdate,6,2) %>','<%= right(oReport.FItemList(i).fsale_startdate,2) %>','<%= left(oReport.FItemList(i).fsale_enddate,4) %>','<%= mid(oReport.FItemList(i).fsale_enddate,6,2) %>','<%= right(oReport.FItemList(i).fsale_enddate,2) %>');" value="�귣�庰" class="button">
	</td>
</tr>
<%	Next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=9 align="left">
		����ո���� : <%=FormatNumber(totrealsellprice/oReport.FResultCount,0) %>��
	</td>
	<!--<td align="right"><%'= FormatNumber(totsellprice,0) %></td>-->
	<td></td>
	<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>	
	<!--<td align="right"><%'= FormatNumber(totselljumuncnt,0) %></td>-->
	<td align="right"><%= FormatNumber(totsellCnt,0) %></td>
	<td></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="15">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

<%
set oReport = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->