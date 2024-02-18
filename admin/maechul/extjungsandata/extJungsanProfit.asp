<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsanProfitCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate, toDate
dim research : research = requestCheckvar(request("research"),10)
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
dim makerid  : makerid = requestCheckvar(request("makerid"),32)
dim rtnExcept : rtnExcept = requestCheckvar(request("rtnExcept"),10)
dim minusgain : minusgain = requestCheckvar(request("minusgain"),10)
dim ctp : ctp = requestCheckvar(request("ctp"),10)

if (ctp="") then ctp="C"

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)
	toDate = LEFT(DateAdd("d", -1,now()),10)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	yyyy2 = Cstr(Year(toDate))
	mm2 = Cstr(Month(toDate))
	dd2 = Cstr(day(toDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2)
end if


Dim oExtJungsanProfit, i
set oExtJungsanProfit = new CExtJungsanProfit
	oExtJungsanProfit.FPageSize = 100
	oExtJungsanProfit.FCurrPage = 1

	oExtJungsanProfit.FRectStartdate = fromDate
	oExtJungsanProfit.FRectEndDate = toDate

	''oExtJungsanProfit.FRectGroupGubun = grpgubun
	oExtJungsanProfit.FRectSellSite = sellsite
    oExtJungsanProfit.FRectReturnExcept = CHKIIF(rtnExcept="on","1","")
    oExtJungsanProfit.FRectMinusGainOnly = CHKIIF(minusgain="on","1","")
    
    if (ctp="D") then
        oExtJungsanProfit.GetExtUpcheDlvProfit
    else
        oExtJungsanProfit.GetExtJungsanProfit
    end if

Dim TTLextitemno
Dim TTLextTenMeachulPrice, TTLextTenJungsanPrice, TTLtenbuycash, TTLjungsangain
Dim TTLU_extTenJungsanPrice, TTLU_buycash, TTLU_jungsangain
Dim TTLW_extTenJungsanPrice, TTLW_buycash, TTLW_jungsangain
Dim TTLM_extTenJungsanPrice, TTLM_buycash, TTLM_jungsangain
Dim TTLN_extTenMeachulPrice, TTLN_buycash, TTLN_jungsangain

dim linkParam : linkParam = "menupos="&menupos&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&rtnExcept="&rtnExcept&"&minusgain="&minusgain&"&ctp="&ctp


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function rePage(isellsite){
    document.frm.sellsite.value = isellsite;
    document.frm.submit();
}
</script>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���޸�:
		<select class="select" name="sellsite" onChange="jsGrpAuto(this.value);">
			<option value="">- ��ü -</option>
			<option value="interpark" <% if (sellsite = "interpark") then %>selected<% end if %> >������ũ</option>
			<option value="lotteimall" <% if (sellsite = "lotteimall") then %>selected<% end if %> >�Ե����̸�</option>
			<option value="lotteCom" <% if (sellsite = "lotteCom") then %>selected<% end if %> >�Ե�����</option>
			<option value="11st1010" <% if (sellsite = "11st1010") then %>selected<% end if %> >11����</option>
			<option value="auction1010" <% if (sellsite = "auction1010") then %>selected<% end if %> >����</option>
			<option value="gmarket1010" <% if (sellsite = "gmarket1010") then %>selected<% end if %> >������(NEW)</option>
			<!-- option value="lotteComM" <% if (sellsite = "lotteComM") then %>selected<% end if %> >�Ե�����(������)</option -->
			<option value="gseshop" <% if (sellsite = "gseshop") then %>selected<% end if %> >GS��</option>
			<!-- option value="dnshop" <% if (sellsite = "dnshop") then %>selected<% end if %> >��ؼ�</option -->
			<option value="cjmall" <% if (sellsite = "cjmall") then %>selected<% end if %> >CJ��</option>
			<!-- option value="wizwid" <% if (sellsite = "wizwid") then %>selected<% end if %> >��������</option -->
			<!-- option value="gabangpop" <% if (sellsite = "gabangpop") then %>selected<% end if %> >�м���(������)</option -->
			<!-- option value="wconcept" <% if (sellsite = "wconcept") then %>selected<% end if %> >����������</option -->
			<!-- option value="privia" <% if (sellsite = "privia") then %>selected<% end if %> >�����������</option -->
			<!-- option value="player" <% if (sellsite = "player") then %>selected<% end if %> >�÷��̾�</option -->
			<option value="homeplus" <% if (sellsite = "homeplus") then %>selected<% end if %> >Ȩ�÷���</option>
			<option value="ssg" <% if (sellsite = "ssg") then %>selected<% end if %> >SSG</option>
			<option value="ssg6006" <% if (sellsite = "ssg6006") then %>selected<% end if %> >SSG-�̸�Ʈ</option>
			<option value="ssg6007" <% if (sellsite = "ssg6007") then %>selected<% end if %> >SSG-ssg</option>
			<option value="nvstorefarm" <% if (sellsite = "nvstorefarm") then %>selected<% end if %> >�������</option>
			<option value="ezwel" <% if (sellsite = "ezwel") then %>selected<% end if %> >���������</option>
			<option value="kakaogift" <% if (sellsite = "kakaogift") then %>selected<% end if %> >īī������Ʈ</option>
			<option value="coupang" <% if (sellsite = "coupang") then %>selected<% end if %> >����</option>
			<option value="halfclub" <% if (sellsite = "halfclub") then %>selected<% end if %> >����Ŭ��</option>
			<option value="hmall" <% if (sellsite = "hmall") then %>selected<% end if %> >Hmall</option>
			
		</select>
		
		&nbsp;
		* ��������:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        &nbsp;
		* ����:
        <select name="ctp">
            <option value="C" <%=CHKIIF(ctp="C","selected","")%> >��ǰ(�����������)
            <option value="D" <%=CHKIIF(ctp="D","selected","")%> >��ü���ǹ�ۺ�(�ֹ��Է±���)
        </select>
		
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr>
    <td bgcolor="#FFFFFF">
        &nbsp;&nbsp;<input type="checkbox" name="rtnExcept" <%=CHKIIF(rtnExcept="on","checked","")%> >��ǰ��������
        &nbsp;&nbsp;<input type="checkbox" name="minusgain" <%=CHKIIF(minusgain="on","checked","")%> >���̳ʽ����͸�����
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p>
<!-- �׼� ���� -->
<form name="frmAct">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	
	</td>
	<td align="right">
	   
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<% if (ctp="C") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100" rowspan="2">���޸�</td>
    <td width="100" rowspan="2">�귣��ID</td>
	
    <td width="70" rowspan="2">����</td>
    <td width="80" rowspan="2">�����</td>
    <td width="80" rowspan="2">�����</td>
    <td width="80" rowspan="2">���Ծ�</td>
    <td width="80" rowspan="2">����</td>

	<td colspan="3">��ü</td>
	<td colspan="3">��Ź</td>
	<td colspan="3">����</td>
	<td colspan="3">�̸���</td>
    <td width="50" rowspan="2">��</td>
	
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	
    <td width="60">�����</td>
    <td width="60">���Ծ�</td>
    <td width="60">����</td>

    <td width="60">�����</td>
    <td width="60">���Ծ�</td>
    <td width="60">����</td>

    <td width="60">�����</td>
    <td width="60">���Ծ�</td>
    <td width="60">����</td>

    <td width="60">�����</td>
    <td width="60">���Ծ�</td>
    <td width="60">����</td>
	
</tr>
<% for i=0 to oExtJungsanProfit.FresultCount -1 %>
<%
    TTLextitemno = TTLextitemno + oExtJungsanProfit.FItemList(i).Fextitemno
    TTLextTenMeachulPrice = TTLextTenMeachulPrice + oExtJungsanProfit.FItemList(i).FextTenMeachulPrice
    TTLextTenJungsanPrice = TTLextTenJungsanPrice + oExtJungsanProfit.FItemList(i).FextTenJungsanPrice
    TTLtenbuycash 				= TTLtenbuycash + oExtJungsanProfit.FItemList(i).Ftenbuycash
    TTLjungsangain 				= TTLjungsangain + oExtJungsanProfit.FItemList(i).Fjungsangain
    TTLU_extTenJungsanPrice = TTLU_extTenJungsanPrice + oExtJungsanProfit.FItemList(i).FU_extTenJungsanPrice
    TTLU_buycash = TTLU_buycash + oExtJungsanProfit.FItemList(i).FU_buycash
    TTLU_jungsangain = TTLU_jungsangain + oExtJungsanProfit.FItemList(i).FU_jungsangain
    TTLW_extTenJungsanPrice = TTLW_extTenJungsanPrice + oExtJungsanProfit.FItemList(i).FW_extTenJungsanPrice
    TTLW_buycash = TTLW_buycash + oExtJungsanProfit.FItemList(i).FW_buycash
    TTLW_jungsangain = TTLW_jungsangain + oExtJungsanProfit.FItemList(i).FW_jungsangain
    TTLM_extTenJungsanPrice = TTLM_extTenJungsanPrice + oExtJungsanProfit.FItemList(i).FM_extTenJungsanPrice
    TTLM_buycash = TTLM_buycash + oExtJungsanProfit.FItemList(i).FM_buycash
    TTLM_jungsangain = TTLM_jungsangain + oExtJungsanProfit.FItemList(i).FM_jungsangain
    TTLextTenMeachulPrice = TTLextTenMeachulPrice + oExtJungsanProfit.FItemList(i).FN_extTenJungsanPrice
    TTLN_buycash = TTLN_buycash + oExtJungsanProfit.FItemList(i).FN_buycash
    TTLN_jungsangain = TTLN_jungsangain + oExtJungsanProfit.FItemList(i).FN_jungsangain
%>
<tr bgcolor="FFFFFF" align="right" >
	<td align="left" ><%=oExtJungsanProfit.FItemList(i).FSellsite %></td>
    <td align="left" ><%=oExtJungsanProfit.FItemList(i).FMakerid %></td>
    
    <td align="center"><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fextitemno,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FextTenMeachulPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FextTenJungsanPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Ftenbuycash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fjungsangain,0) %></td>

    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FU_extTenJungsanPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FU_buycash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FU_jungsangain,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FW_extTenJungsanPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FW_buycash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FW_jungsangain,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FM_extTenJungsanPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FM_buycash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FM_jungsangain,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FN_extTenJungsanPrice,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FN_buycash,0) %></td>
    <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FN_jungsangain,0) %></td>
    <td align="center">
    <% if oExtJungsanProfit.FItemList(i).FMakerid<>"" then %>
        <a target="extJungsanProfitDetail" href="extJungsanProfitDetail.asp?<%=linkParam%>&sellsite=<%=oExtJungsanProfit.FItemList(i).FSellsite %>&makerid=<%=oExtJungsanProfit.FItemList(i).FMakerid %>">����</a>
    <% else %>
		<a href="?<%=linkParam%>&sellsite=<%=oExtJungsanProfit.FItemList(i).FSellsite %>">����</a>
    <% end if %>
    </td>
</tr>
<% next %>
<tr bgcolor="FFFFFF" align="right">
    <td></td>
    <td></td>
    <td align="center"><%=FormatNumber(TTLextitemno,0)%></td>
    <td><%=FormatNumber(TTLextTenMeachulPrice,0)%></td>
    <td><%=FormatNumber(TTLextTenJungsanPrice,0)%></td>
    <td><%=FormatNumber(TTLtenbuycash,0)%></td>
    <td><%=FormatNumber(TTLjungsangain,0)%></td>
    <td><%=FormatNumber(TTLU_extTenJungsanPrice,0)%></td>
    <td><%=FormatNumber(TTLU_buycash,0)%></td>
    <td><%=FormatNumber(TTLU_jungsangain,0)%></td>
    <td><%=FormatNumber(TTLW_extTenJungsanPrice,0)%></td>
    <td><%=FormatNumber(TTLW_buycash,0)%></td>
    <td><%=FormatNumber(TTLW_jungsangain,0)%></td>
    <td><%=FormatNumber(TTLM_extTenJungsanPrice,0)%></td>
    <td><%=FormatNumber(TTLM_buycash,0)%></td>
    <td><%=FormatNumber(TTLM_jungsangain,0)%></td>
    <td><%=FormatNumber(TTLN_extTenMeachulPrice,0)%></td>
    <td><%=FormatNumber(TTLN_buycash,0)%></td>
    <td><%=FormatNumber(TTLN_jungsangain,0)%></td>
    <td></td>
</tr>
</table>
<% else %>
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="100" >���޸�</td>
        <td width="100" >�귣��ID</td>
        
        <td width="70" >���ֹ���</td>
        <td width="80" >�����</td>
        <td width="80" >���Ծ�</td>
        <td width="80" >����</td>
        <td width="100" >��ۺ����</td>
        <td width="100" >��ۺ�</td>
        <td></td>
        <td width="50" >��</td>
    </tr>
    <% for i=0 to oExtJungsanProfit.FresultCount -1 %>
    <%
        TTLextitemno = TTLextitemno + oExtJungsanProfit.FItemList(i).Fextitemno
        TTLextTenMeachulPrice = TTLextTenMeachulPrice + oExtJungsanProfit.FItemList(i).FextTenMeachulPrice
        'TTLextTenJungsanPrice = TTLextTenJungsanPrice + oExtJungsanProfit.FItemList(i).FextTenJungsanPrice
        TTLtenbuycash 				= TTLtenbuycash + oExtJungsanProfit.FItemList(i).Ftenbuycash
        TTLjungsangain 				= TTLjungsangain + oExtJungsanProfit.FItemList(i).Fjungsangain
    %>
    <tr bgcolor="FFFFFF" align="right" >
        <td align="left" ><%=oExtJungsanProfit.FItemList(i).FSellsite %></td>
        <td align="left" ><%=oExtJungsanProfit.FItemList(i).FMakerid %></td>
        
        <td align="center"><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fextitemno,0) %></td>
        <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).FextTenMeachulPrice,0) %></td>
        <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Ftenbuycash,0) %></td>
        <td><%=FormatNumber(oExtJungsanProfit.FItemList(i).Fjungsangain,0) %></td>
        <td align="center">
            <% if oExtJungsanProfit.FItemList(i).FdefaultDeliveryType="7" then %>
            ��ü���ҹ��
            <% elseif oExtJungsanProfit.FItemList(i).FdefaultDeliveryType="9" then %>
            <%=FormatNumber(oExtJungsanProfit.FItemList(i).FdefaultFreeBeasongLimit,0)%>
            <% end if %>
        </td>
        <td align="center">
            <% if oExtJungsanProfit.FItemList(i).FdefaultDeliveryType="9" or oExtJungsanProfit.FItemList(i).FdefaultDeliveryType="7" then %>
            <%=FormatNumber(oExtJungsanProfit.FItemList(i).FdefaultDeliverPay,0)%>
            <% end if %>
        </td>
        <td></td>
        <td align="center">
        <% if oExtJungsanProfit.FItemList(i).FMakerid<>"" then %>
            <a target="extJungsanProfitDetail" href="extJungsanProfitDetail.asp?<%=linkParam%>&sellsite=<%=oExtJungsanProfit.FItemList(i).FSellsite %>&makerid=<%=oExtJungsanProfit.FItemList(i).FMakerid %>">����</a>
        <% else %>
            <a href="?<%=linkParam%>&sellsite=<%=oExtJungsanProfit.FItemList(i).FSellsite %>">����</a>
        <% end if %>
        </td>
    </tr>
    <% next %>
    <tr bgcolor="FFFFFF" align="right">
        <td></td>
        <td></td>
        <td align="center"><%=FormatNumber(TTLextitemno,0)%></td>
        <td><%=FormatNumber(TTLextTenMeachulPrice,0)%></td>
        <td><%=FormatNumber(TTLtenbuycash,0)%></td>
        <td><%=FormatNumber(TTLjungsangain,0)%></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    </table>
<% end if %>
<%
set oExtJungsanProfit = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
