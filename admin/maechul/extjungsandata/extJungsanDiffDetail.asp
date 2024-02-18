<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� ���� ����α� ������ ��
' Hieditor : 2018.04.23 ������ ���� 
'###########################################################
Server.ScriptTimeOut = 180
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/extjungsan/extjungsandiffcls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sellsite,rdost
dim searchfield,searchtext 
dim chulgoDate_yyyy1,chulgoDate_yyyy2,chulgoDate_mm1,chulgoDate_mm2,chulgoDate_dd1,chulgoDate_dd2
dim confirmDate_yyyy1,confirmDate_yyyy2,confirmDate_mm1,confirmDate_mm2,confirmDate_dd1,confirmDate_dd2
dim chulgoDate_fromDate, chulgoDate_toDate,confirmDate_fromDate, confirmDate_toDate, tmpDate
dim iCurrpage, iTotCnt, iTotPage,iPageSize,iPerCnt
dim arrList, intLoop
dim chkErr
dim extMeachul, logMeachul
  iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
  sellsite = requestCheckVar(Request("sellsite"),10)  
  rdost = requestCheckVar(Request("rdost"),1)
  chkErr = requestCheckVar(Request("chkErr"),1)
  searchfield = requestCheckVar(Request("searchfield"),30)
  searchtext = requestCheckVar(Request("searchtext"),120)
  iPageSize = requestCheckVar(Request("ips"),10)
	chulgoDate_yyyy1   = 	requestCheckvar(request("chulgoDate_yyyy1"),4)
	chulgoDate_mm1     = requestCheckvar(request("chulgoDate_mm1"),2)
	chulgoDate_dd1     = requestCheckvar(request("chulgoDate_dd1"),2)
	chulgoDate_yyyy2   = requestCheckvar(request("chulgoDate_yyyy2"),4)
	chulgoDate_mm2     = requestCheckvar(request("chulgoDate_mm2"),2)
	chulgoDate_dd2     = requestCheckvar(request("chulgoDate_dd2"),2)
	confirmDate_yyyy1   = 	requestCheckvar(request("confirmDate_yyyy1"),4)
	confirmDate_mm1     = requestCheckvar(request("confirmDate_mm1"),2)
	confirmDate_dd1     = requestCheckvar(request("confirmDate_dd1"),2)
	confirmDate_yyyy2   = requestCheckvar(request("confirmDate_yyyy2"),4)
	confirmDate_mm2     = requestCheckvar(request("confirmDate_mm2"),2)
	confirmDate_dd2     = requestCheckvar(request("confirmDate_dd2"),2)
if sellsite ="" then sellsite ="ssg"
if rdost	="" then rdost="1"
IF iCurrpage = "" THEN		iCurrpage = 1
if iPageSize ="" then iPageSize = 20
		iPerCnt = 10		'�������� ������ ����
	
if (chulgoDate_yyyy1="") then
		if rdost="1" then
		chulgoDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 2), 1)
		else
		chulgoDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
		end if
	chulgoDate_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) ), 1)
 
	chulgoDate_yyyy1 = Cstr(Year(chulgoDate_fromDate))
	chulgoDate_mm1 = Cstr(Month(chulgoDate_fromDate))
	chulgoDate_dd1 = Cstr(day(chulgoDate_fromDate))

	tmpDate = DateAdd("d", -1, chulgoDate_toDate)
	chulgoDate_yyyy2 = Cstr(Year(tmpDate))
	chulgoDate_mm2 = Cstr(Month(tmpDate))
	chulgoDate_dd2 = Cstr(day(tmpDate))
else
	chulgoDate_fromDate = DateSerial(chulgoDate_yyyy1, chulgoDate_mm1, chulgoDate_dd1)
	chulgoDate_toDate = DateSerial(chulgoDate_yyyy2, chulgoDate_mm2, chulgoDate_dd2+1)
end if


if (confirmDate_yyyy1="") then 
	confirmDate_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1) 
	confirmDate_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) ), 1)
 
	confirmDate_yyyy1 = Cstr(Year(confirmDate_fromDate))
	confirmDate_mm1 = Cstr(Month(confirmDate_fromDate))
	confirmDate_dd1 = Cstr(day(confirmDate_fromDate))

	tmpDate = DateAdd("d", -1, confirmDate_toDate)
	confirmDate_yyyy2 = Cstr(Year(tmpDate))
	confirmDate_mm2 = Cstr(Month(tmpDate))
	confirmDate_dd2 = Cstr(day(tmpDate))
else
	confirmDate_fromDate = DateSerial(confirmDate_yyyy1, confirmDate_mm1, confirmDate_dd1)
	confirmDate_toDate = DateSerial(confirmDate_yyyy2, confirmDate_mm2, confirmDate_dd2+1)
end if

dim cEJDiff
 set cEJDiff = new CextJungsanDiff
 cEJDiff.FCPage = iCurrpage
 cEJDiff.FPSize = iPageSize
 cEJDiff.FCGFDate = chulgoDate_fromDate
 cEJDiff.FCGTDate = chulgoDate_toDate
 cEJDiff.FCFFDate = confirmDate_fromDate
 cEJDiff.FCFTDate = confirmDate_toDate
 cEJDiff.FSellsite = sellsite
 cEJDiff.FRectST = rdost
 cEJDiff.FRectErr = chkErr
 if rdost ="1" then
 arrList =cEJDiff.fnGetextJsDiffList
else
 arrList =cEJDiff.fnGetlogJsDiffList
end if
 iTotCnt = cEJDiff.FTotCnt
 extMeachul = cEJDiff.FextMeachul
 logMeachul = cEJDiff.FlogMeachul
 set cEJDiff = nothing
 
 if extMeachul ="" or isNull(extMeachul) then extMeachul =0
 if logMeachul ="" or isNull(logMeachul) then logMeachul =0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	function jsSearh(){
		$("#btnSubmit").prop("disabled", true); 
		document.frm.submit(); 
		}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="research" value="on">
		<tr  bgcolor="#FFFFFF" >
			<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
			<td align="left"> 
				����ó : <% fnGetOptOutMall sellsite %>
				 
				&nbsp;&nbsp;
				* �˻����� :
				<select class="select" name="searchfield">
					<option value=""></option>
					<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option>
				</select>
				<input type="text" class="text" name="searchtext" value="<%= searchtext %>">				
			</td>
			<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="jsSearh();">
			</td>
		</tr>
		<tr   bgcolor="#FFFFFF" >
			<td align="left"> 
				<input type="radio" name="rdost" class="radio" value="1" <%if rdost ="1" then%>checked<%end if%> > ���޸����� ���� 
				<input type="radio" name="rdost" class="radio" value="2" <%if rdost ="2" then%>checked<%end if%>> ����αױ���
				&nbsp;&nbsp;
 				���ޱ���Ȯ����:
			  <% DrawDateBoxdynamic confirmDate_yyyy1, "confirmDate_yyyy1", confirmDate_yyyy2, "confirmDate_yyyy2", confirmDate_mm1, "confirmDate_mm1", confirmDate_mm2, "confirmDate_mm2", confirmDate_dd1, "confirmDate_dd1", confirmDate_dd2, "confirmDate_dd2" %>
			  &nbsp;
				������� :
				<% DrawDateBoxdynamic chulgoDate_yyyy1, "chulgoDate_yyyy1", chulgoDate_yyyy2, "chulgoDate_yyyy2", chulgoDate_mm1, "chulgoDate_mm1", chulgoDate_mm2, "chulgoDate_mm2", chulgoDate_dd1, "chulgoDate_dd1", chulgoDate_dd2, "chulgoDate_dd2" %>
				&nbsp;&nbsp;
			<input type="checkbox" name="chkErr" value="Y" <%if chkErr ="Y" then%> checked<%end if%>>�̸�Ī
			</td>
		</tr>
		<tr  bgcolor="#FFFFFF" >
			<td> 
			
				��ǥ�� :
				<select class="select" name="ips">
					<option value="20" <% if (iPageSize = "20") then %>selected<% end if %> >20</option>
					<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100</option>
					<option value="1000" <% if (iPageSize = "1000") then %>selected<% end if %> >1000</option>
					<option value="3000" <% if (iPageSize = "3000") then %>selected<% end if %> >3000</option>
				</select> 
			</td>
		</tr>
	</form>
</table>
<!-- �˻� �� -->
<p style="padding-top:10px;"></p>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		�˻���� : <b><%=iTotCnt%></b>
		&nbsp;
		������ : <b> /  </b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2">���޸�</td>
	<td colspan="5">�����ֹ�</td>
	<td  colspan="5">����α�</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td>�ֹ���ȣ</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ��ڵ�</td>
	<td>����ݾ�</td>
	<td></td>
	
	<td>�ֹ���ȣ</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ��ڵ�</td>
	<td>����ݾ�</td>
</tr>
<%if isArray(arrList) then%>
	<tr bgcolor="<%=adminColor("pink")%>">
		<td></td>
		<td colspan="3"></td>
		<td align="right"><%=formatnumber(extMeachul,0)%></td>
		<td></td>
		<td colspan="3"></td>
		<td align="right"><%=formatnumber(logMeachul,0)%></td>
		<td></td>
	</tr>
	<%
		for intLoop = 0 To uBound(arrList,2)
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%=arrList(0,intLoop)%></td>
		
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=arrList(3,intLoop)%></td>
		<td align="right"><%=arrList(7,intLoop)%></td>
		<td><%=arrList(4,intLoop)%></td>
		
		<td><%=arrList(9,intLoop)%></td>
		<td><%=arrList(10,intLoop)%></td>
		<td><%=arrList(11,intLoop)%></td>
		<td align="right"><%=arrList(13,intLoop)%></td>
		<td></td>
	</tr>
<%next %>
<%end if%>
</table>
<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->