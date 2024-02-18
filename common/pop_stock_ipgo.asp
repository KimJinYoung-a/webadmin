<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 입출고
' History : 2017.04.10 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/AcountItemIpChulCls.asp"-->
<%

dim ipchulflag,designer,itemgubun,itemid,itemoption
ipchulflag  = RequestCheckVar(request("ipchulflag"),9)
designer    = RequestCheckVar(request("designer"),32)
itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
itemoption  = RequestCheckVar(request("itemoption"),4)

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

fromdate = RequestCheckVar(request("fromdate"),10)
todate = RequestCheckVar(request("todate"),10)

if fromdate<>"" then
	yyyy1 = Left(fromdate,4)
	mm1 = Mid(fromdate,6,2)
	dd1 = Mid(fromdate,9,2)
else
	yyyy1 = RequestCheckVar(request("yyyy1"),4)
	mm1 = RequestCheckVar(request("mm1"),2)
	dd1 = RequestCheckVar(request("dd1"),2)
end if

if todate<>"" then
	yyyy2 = Left(todate,4)
	mm2 = Mid(todate,6,2)
	dd2 = Mid(todate,9,2)
else
	yyyy2 = RequestCheckVar(request("yyyy2"),4)
	mm2 = RequestCheckVar(request("mm2"),2)
	dd2 = RequestCheckVar(request("dd2"),2)
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oacctipchul
set oacctipchul = new CAcountItemIpChul
oacctipchul.FRectStartday = fromDate
oacctipchul.FRectEndday   = toDate
'oacctipchul.FRectGubun   = ipchulflag	'/잘못 만든듯. 구분자 때문에 아무것도 안나옴. 주석처리	'/2017.03.03 한용민
oacctipchul.FRectDesigner = designer
oacctipchul.FRectItemGubun = itemgubun
oacctipchul.FRectItemID = itemid
oacctipchul.FRectItemOption = itemoption
oacctipchul.FRectDeletInclude = "on"

oacctipchul.getIpChulListByItem

dim i, totitemno

totitemno=0
%>

<script type='text/javascript'>

function EditIpCulNSheet(code,makerid){
	var popwin = window.open('/admin/inspect/popadminipchuledit.asp?code=' + code + '&makerid=' + makerid,'popadminipchuledit','width=900,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<form name="frm" method="get" action="">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	상품코드 :
	        	<select class="select" name="itemgubun">
	        		<option value="10" <%= chkIIF(itemgubun="10","selected","") %> >10</option>
	        		<option value="55" <%= chkIIF(itemgubun="55","selected","") %> >55</option>
					<option value="70" <%= chkIIF(itemgubun="70","selected","") %> >70</option>
	        		<option value="80" <%= chkIIF(itemgubun="80","selected","") %> >80</option>
	        		<option value="90" <%= chkIIF(itemgubun="90","selected","") %> >90</option>
	        	</select>
	        	<input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">&nbsp;&nbsp;
	        	<input type="text" class="text_ro" name="itemoption" value="<%= itemoption %>" size=4 maxlength=4 readonly>

	        	브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
	        	<!--구분 : <select name="ipchulflag">
	        				<option value=''  >선택</option>
	        				<option value='I' <% 'if ipchulflag="I" then response.write "selected" %> >입고</option>
	        				<option value='S' <% 'if ipchulflag="S" then response.write "selected" %> >출고</option>
	        				<option value='E' <% 'if ipchulflag="E" then response.write "selected" %> >기타출고</option>
	        		  </select>-->
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	        </td>
	        <td align="right" bgcolor="F4F4F4">(최대 1,000건)</td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">입출코드</td>
      <td width="60">구분</td>
      <td width="60">입출고일</td>
      <td width="100">입출고처ID</td>
      <td width="25">구분</td>
      <td width="50">상품코드</td>
      <td>아이템명</td>
      <td>옵션</td>
      <td width="50">소비자가</td>
      <td width="50">매입가</td>
      <td width="50">공급가</td>
      <td width="30">수량</td>
      <td width="30">삭제여부</td>
      <% if C_ADMIN_AUTH then %>
      <td>관리자</td>
      <% end if %>
    </tr>
    <% for i=0 to oacctipchul.FResultCount-1 %>
    <%
    totitemno = totitemno + oacctipchul.FItemList(i).FItemNo
    %>
    <tr align="center" bgcolor="<%= CHKIIF(oacctipchul.FItemList(i).isDeleted,"#EEEEEE","#FFFFFF") %>">
      <% if Left(oacctipchul.FItemList(i).FIpchulCode,2)="ST" then %>
      <td><a href="/admin/newstorage/ipgodetail.asp?idx=<%= oacctipchul.FItemList(i).FIpChulidx %>&menupos=539" target="_blank"><font color="<%= oacctipchul.FItemList(i).GetIpchulColor %>"><%= oacctipchul.FItemList(i).FIpchulCode %></font></a></td>
      <% else %>
      <td><a href="/admin/newstorage/chulgodetail.asp?idx=<%= oacctipchul.FItemList(i).FIpChulidx %>&menupos=540" target="_blank"><font color="<%= oacctipchul.FItemList(i).GetIpchulColor %>"><%= oacctipchul.FItemList(i).FIpchulCode %></font></a></td>
      <% end if %>
      <td><font color="<%= oacctipchul.FItemList(i).GetDivCodeColor %>"><%= oacctipchul.FItemList(i).GetDivCodeName %></font></td>
      <td><%= oacctipchul.FItemList(i).Fexecutedt %></td>
      <td><%= oacctipchul.FItemList(i).FSocID %></td>
      <td><%= oacctipchul.FItemList(i).FItemgubun %></td>
      <td><%= oacctipchul.FItemList(i).FItemID %></td>
      <td><%= oacctipchul.FItemList(i).FItemName %></td>
      <td><%= oacctipchul.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FSellCash,0) %></td>
      <% if oacctipchul.FItemList(i).Fipchulflag="I" then %>
      <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FsuplyCash,0) %></td>
      <td align="right">&nbsp;</td>
      <% else %>
      <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FbuyCash,0) %></td>
      <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FsuplyCash,0) %></td>
      <% end if %>
      <td align="center"><%= oacctipchul.FItemList(i).FItemNo %></td>
      <td align="center">
        <% if  (oacctipchul.FItemList(i).isDeleted) then %>
        <font color="#FF3333" alt="<%= oacctipchul.FItemList(i).FmasterDeldt %>">삭제됨</font>
        <% end if %>
      </td>
      <% if C_ADMIN_AUTH then %>
      <td><a href="javascript:EditIpCulNSheet('<%= oacctipchul.FItemList(i).FIpchulCode %>','<%= oacctipchul.FItemList(i).FiMakerid %>');"><font color=red>ED</font></a></td>
      <% end if %>
    </tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	  <td align=center>Total</td>
	  <td colspan=10></td>
	  <td align=center><%= FormatNumber(totitemno,0) %></td>
	  <td align=center></td>
	  <% if C_ADMIN_AUTH then %>
      <td></td>
      <% end if %>
	</tr>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oacctipchul = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
