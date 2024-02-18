<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy, mm, dd, jungsanfixdate
dim fromDate ,toDate, tmpDate
dim sellsite, jungsantype, searchfield, searchtext

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

yyyy1   = requestCheckvar(request("yyyy1"),4)
mm1     = requestCheckvar(request("mm1"),2)
dd1     = requestCheckvar(request("dd1"),2)
yyyy2   = requestCheckvar(request("yyyy2"),4)
mm2     = requestCheckvar(request("mm2"),2)
dd2     = requestCheckvar(request("dd2"),2)
jungsanfixdate = requestCheckvar(request("jungsanfixdate"),2)

sellsite		= requestCheckvar(request("sellsite"),32)
jungsantype		= requestCheckvar(request("jungsantype"),32)
searchfield 	= requestCheckvar(request("searchfield"),32)
searchtext 		= Replace(Replace(requestCheckvar(request("searchtext"),32), "'", ""), Chr(34), "")

dim extjdate : extjdate = requestCheckvar(request("extjdate"),8) ''YYYYMMDD
dim mimap : mimap = requestCheckvar(request("mimap"),10)
dim mimapminus : mimapminus = requestCheckvar(request("mimapminus"),10)

dim vatyn : vatyn = requestCheckvar(request("vatyn"),1)
dim retonly : retonly = requestCheckvar(request("retonly"),10)
dim errexists : errexists = requestCheckvar(request("errexists"),10)
dim existsBigo : existsBigo = requestCheckvar(request("existsBigo"),10)
dim dotview : dotview = requestCheckvar(request("dotview"),10)
dim FormatDotNo : FormatDotNo = 0
dim exceptcost0 : exceptcost0 = requestCheckvar(request("exceptcost0"),10)
dim xjungsanchk : xjungsanchk = requestCheckvar(request("xjungsanchk"),20)

if (dotview<>"") then FormatDotNo = 2

if (extjdate="") then
    extjdate = replace(LEFT(dateAdd("d",-1,now()),10),"-","")
end if

if (page="") then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, DateAdd("m",1,toDate))
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.action = "extJungsanDataList.asp";
    document.frm.submit();
}

function jsSubmit(){
    document.frm.page.value = "1";
    document.frm.action = "extJungsanDataList.asp";
    document.frm.submit();
}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popAPIJungsanData() {
    var window_width = 600;
    var window_height = 500;

    var popwin2 = window.open("/admin/maechul/extjungsandata/popApiJungsanData.asp","popAPIJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin2.focus();
}

function popJungsanFixData(v){
	if (v == ''){
		alert('�˻������� ���޸��� �˻� �ϼ���');
		return;
	}
	if(v=='ssg6006' || v=='ssg6007'){
		v = 'ssg';
	}
	if(confirm("2022-01-01 ���� �����ϸ� ����Ȯ���Ϸ� �Էµ˴ϴ�.\n\n-����Ȯ���� ���Է°� ����\n-��Ұ� ����\n-���Ϸ�Ǹ�\n\n�����Ͻðڽ��ϱ�?")){
        var frm = document.editFixFrm;
        frm.sellsite.value=v;
        frm.submit();
	}
}

function popExtSiteJungsanByAPI(sitename,comp){
    var yyyymmdd = comp.form.extjdate.value;

    <% If application("Svr_Info")="Dev" Then %>
        var idomain = "http://testwapi.10x10.co.kr";
    <% else %>
        var idomain = "<%=apiURL%>";
    <% end if %>

    if (sitename=="lottecom"){
        var popwin = window.open(idomain+"/outmall/proc/LotteCom_JungsanRsvProc.asp?yyyymmdd="+yyyymmdd,"LotteCom_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
    }

    if (sitename=="interpark"){
        var popwin = window.open(idomain+"/outmall/proc/Interpark_JungsanRsvProc.asp?yyyymmdd="+yyyymmdd,"Interpark_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
    }

    if (sitename=="ssg"){
		<% if (sellsite="ssg6006") or (sellsite="ssg6007") then %>
		var popwin = window.open(idomain+"/outmall/ssg/xSiteJungsan_ssg_Process.asp?yyyymmdd="+yyyymmdd+"&sellsite=<%=sellsite%>","ssg_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
		<% else %>
        var popwin = window.open(idomain+"/outmall/ssg/xSiteJungsan_ssg_Process.asp?yyyymmdd="+yyyymmdd,"ssg_JungsanRsvProc","width=400,height=200,crollbars=yes,resizable=yes,status=yes");
		<% end if %>
    }

    popwin.focus();
}

function jsExcelDown(){
	if (frm.sellsite.value.length<1){
		alert('����Ʈ�� �����ϼ���.')
		return;
	}

	alert('�ִ� 2���� ���� �����ؿ�.')
    document.frm.action = "extJungsanDataList_excel.asp";
    document.frm.submit();
}

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sellsite%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

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
		&nbsp;
		���޸�:
		<select class="select" name="sellsite">
			<option></option>
			<option value="interpark" <% if (sellsite = "interpark") then %>selected<% end if %> >������ũ</option>
			<option value="lotteimall" <% if (sellsite = "lotteimall") then %>selected<% end if %> >�Ե����̸�</option>
			<!-- <option value="lotteCom" <% if (sellsite = "lotteCom") then %>selected<% end if %> >�Ե�����</option> -->
			<option value="11st1010" <% if (sellsite = "11st1010") then %>selected<% end if %> >11����</option>
			<option value="auction1010" <% if (sellsite = "auction1010") then %>selected<% end if %> >����</option>
			<option value="gmarket1010" <% if (sellsite = "gmarket1010") then %>selected<% end if %> >������(NEW)</option>
			<!--option value="lotteComM" <% if (sellsite = "lotteComM") then %>selected<% end if %> >�Ե�����(������)</option -->
			<option value="gseshop" <% if (sellsite = "gseshop") then %>selected<% end if %> >GS��</option>
			<!--option value="dnshop" <% if (sellsite = "dnshop") then %>selected<% end if %> >��ؼ�</option -->
			<option value="cjmall" <% if (sellsite = "cjmall") then %>selected<% end if %> >CJ��</option>
			<!--option value="wizwid" <% if (sellsite = "wizwid") then %>selected<% end if %> >��������</option -->
			<!--option value="gabangpop" <% if (sellsite = "gabangpop") then %>selected<% end if %> >�м���(������)</option -->
			<option value="wconcept1010" <% if (sellsite = "wconcept1010") then %>selected<% end if %> >W����</option>
			<option value="withnature1010" <% if (sellsite = "withnature1010") then %>selected<% end if %> >�ڿ��̶�</option>
			<option value="GS25" <% if (sellsite = "GS25") then %>selected<% end if %> >GS25ī�޷α�</option>
			<!--option value="privia" <% if (sellsite = "privia") then %>selected<% end if %> >�����������</option -->
			<!--option value="player" <% if (sellsite = "player") then %>selected<% end if %> >�÷��̾�</option -->
			<!-- <option value="homeplus" <% if (sellsite = "homeplus") then %>selected<% end if %> >Ȩ�÷���</option> -->
			<option value="ssg" <% if (sellsite = "ssg") then %>selected<% end if %> >SSG</option>
			<option value="ssg6006" <% if (sellsite = "ssg6006") then %>selected<% end if %> >SSG-�̸�Ʈ</option>
			<option value="ssg6007" <% if (sellsite = "ssg6007") then %>selected<% end if %> >SSG-ssg</option>
			<option value="shintvshopping" <% if (sellsite = "shintvshopping") then %>selected<% end if %> >�ż���TV����</option>
			<option value="skstoa" <% if (sellsite = "skstoa") then %>selected<% end if %> >SKSTOA</option>
			<!-- <option value="wetoo1300k" <% if (sellsite = "wetoo1300k") then %>selected<% end if %> >1300k</option> -->
			<option value="nvstorefarm" <% if (sellsite = "nvstorefarm") then %>selected<% end if %> >�������</option>
			<!-- <option value="nvstorefarmclass" <% if (sellsite = "nvstorefarmclass") then %>selected<% end if %> >�������-Ŭ����</option> -->
			<option value="nvstoremoonbangu" <% if (sellsite = "nvstoremoonbangu") then %>selected<% end if %> >������� ���汸</option>
			<option value="Mylittlewhoopee" <% if (sellsite = "Mylittlewhoopee") then %>selected<% end if %> >������� Ĺ�ص�</option>
			<option value="nvstoregift" <% if (sellsite = "nvstoregift") then %>selected<% end if %> >������� �����ϱ�</option>
			<option value="ezwel" <% if (sellsite = "ezwel") then %>selected<% end if %> >���������</option>
			<option value="kakaogift" <% if (sellsite = "kakaogift") then %>selected<% end if %> >īī������Ʈ</option>
			<option value="kakaostore" <% if (sellsite = "kakaostore") then %>selected<% end if %> >īī���彺���</option>
			<option value="boribori1010" <% if (sellsite = "boribori1010") then %>selected<% end if %> >��������</option>
			<option value="coupang" <% if (sellsite = "coupang") then %>selected<% end if %> >����</option>
			<!-- <option value="halfclub" <% if (sellsite = "halfclub") then %>selected<% end if %> >����Ŭ��</option> -->
			<option value="hmall1010" <% if (sellsite = "hmall1010") then %>selected<% end if %> >Hmall</option>
			<option value="WMP" <% if (sellsite = "WMP") then %>selected<% end if %> >WMP</option>
			<!-- <option value="wmpfashion" <% if (sellsite = "wmpfashion") then %>selected<% end if %> >WMPW�м�</option> -->
			<option value="LFmall" <% if (sellsite = "LFmall") then %>selected<% end if %> >LFmall</option>
			<option value="lotteon" <% if (sellsite = "lotteon") then %>selected<% end if %> >�Ե�On</option>
			<option value="yes24" <% if (sellsite = "yes24") then %>selected<% end if %> >yes24</option>
			<option value="alphamall" <% if (sellsite = "alphamall") then %>selected<% end if %> >���ĸ�</option>
			<option value="ohou1010" <% if (sellsite = "ohou1010") then %>selected<% end if %> >��������</option>
			<!-- <option value="wadsmartstore" <% if (sellsite = "wadsmartstore") then %>selected<% end if %> >�͵彺��Ʈ�����</option> -->
			<option value="casamia_good_com" <% if (sellsite = "casamia_good_com") then %>selected<% end if %> >���̾�</option>
			<option value="cookatmall" <% if (sellsite = "cookatmall") then %>selected<% end if %> >��Ĺ</option>
			<option value="aboutpet" <% if (sellsite = "aboutpet") then %>selected<% end if %> >��ٿ���</option>
			<option value="goodshop1010" <% if (sellsite = "goodshop1010") then %>selected<% end if %> >�¼�</option>
		</select>
		&nbsp;
		���޸�������:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		����Ȯ����:
		<select name="jungsanfixdate" class="select">
			<option value="">��ü</option>
			<option value="A" <%= Chkiif(jungsanfixdate="A","selected","")  %>>���</option>
			<option value="B" <%= Chkiif(jungsanfixdate="B","selected","")  %>>������</option>
			<option value="C" <%= Chkiif(jungsanfixdate="C","selected","")  %>>���Է�</option>
		</select>
		&nbsp;
		������:
		<select class="select" name="jungsantype">
			<option></option>
			<option value="C" <% if (jungsantype = "C") then %>selected<% end if %> >��ǰ��(�Һ��ڸ���)</option>
			<option value="D" <% if (jungsantype = "D") then %>selected<% end if %> >��ۺ�</option>
			<option value="E" <% if (jungsantype = "E") then %>selected<% end if %> >��Ÿ����</option>
		</select>
		&nbsp;
		��������:
		<select class="select" name="vatyn">
			<option value="" <% if (vatyn = "") then %>selected<% end if %> ></option>
			<option value="Y" <% if (vatyn = "Y") then %>selected<% end if %> >����</option>
			<option value="N" <% if (vatyn = "N") then %>selected<% end if %> >�鼼</option>
		</select>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="jsSubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="extOrderserial" <% if (searchfield = "extOrderserial") then %>selected<% end if %> >�����ֹ���ȣ</option>
			<option value="extOrgOrderserial" <% if (searchfield = "extOrgOrderserial") then %>selected<% end if %> >���޿��ֹ���ȣ</option>
			<option value="OrgOrderserial" <% if (searchfield = "OrgOrderserial") then %>selected<% end if %> >����(TEN)�ֹ���ȣ</option>
			<option value="matchitemid" <% if (searchfield = "matchitemid") then %>selected<% end if %> >����(TEN)��ǰ��ȣ</option>
		</select>
		<input type="text" class="text" name="searchtext" size="30" value="<%= searchtext %>">
		&nbsp;
		<label><input type="checkbox" name="mimap" <%=CHKIIF(mimap="on","checked","")%> >�̸���(�ֹ�)</label>
		&nbsp;
		<label><input type="checkbox" name="mimapminus" <%=CHKIIF(mimapminus="on","checked","")%> >(���̳ʽ�)�̸��γ�����</label>
		&nbsp;
		<label><input type="checkbox" name="retonly" <%=CHKIIF(retonly="on","checked","")%> >��ǰ������</label>
		&nbsp;
		<label><input type="checkbox" name="errexists" <%=CHKIIF(errexists="on","checked","")%> >����������</label>
		&nbsp;
		<label><input type="checkbox" name="dotview" <%=CHKIIF(dotview="on","checked","")%> >�Ҽ���2�ڸ�ǥ��</label>
		&nbsp;
		<label><input type="checkbox" name="exceptcost0" <%=CHKIIF(exceptcost0="on","checked","")%> >�ǸŰ�0����</label>
		&nbsp;|&nbsp;
		<label><input type="checkbox" name="existsBigo" <%=CHKIIF(existsBigo="on","checked","")%> >���Y</label>
		&nbsp;|&nbsp;
		���� :
		<select name="xjungsanchk">
			<option value="">����
			<option value="summinus" <%=CHKIIF(xjungsanchk="summinus","selected","") %> >�հ踶�̳ʽ�
		</select>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" class="a">
<tr>
	<td>
* ����ݾ� = �Ǹűݾ�-���޺δ�����-���ٺδ�����<br>
* ����ݾ� = ����ݾ�-������<br><br>
* ���޸� ��ǰ ��ۺ�<br>
&nbsp; - 1. ���޸����� �����ǰ ���� ���� �ݿ��ȵ� : ���� ���� ����<br>
&nbsp; - 2. ���ֹ��� ���޸� �������� ��� : ��ü ����ŭ ��ǰ��ۺ� ������(�ߺ� ���� ���)<br>
&nbsp; - 3. ��ǰ����� �����ϰ� ���޸�-�ٹ����� ����������� ���̷� ���� ����ġ

* �����ֹ���ȣ �߰���Ī �ʿ� (GS�� 622364193 ����)
	</td>
	<td align="right">
		<input type="button" value="Excel Down" style="height:50px;" onClick="jsExcelDown();">
	</td>
</tr>
</table>
<%
Dim oCExtJungsanSum
Dim totROWno, totSumItemno, totSumitemcost, totSumOwnCouponPrice, totSumTenCouponPrice, totSumMeachulPrice, totSumCommPrice, totSumJungsanPrice_ETC, totSumJungsanPrice, totMiMapTTLCnt, totBigoSum, totBigoBeasongSum, totReducedPrice
Dim totBigoSum1, totBigoSum2
set oCExtJungsanSum = new CExtJungsan
	oCExtJungsanSum.FRectStartdate = fromDate
	oCExtJungsanSum.FRectEndDate = toDate
	oCExtJungsanSum.FRectJungsanfixdate = jungsanfixdate

	oCExtJungsanSum.FRectSellSite = sellsite
	oCExtJungsanSum.FRectJungsanType = jungsantype

	oCExtJungsanSum.FRectSearchField = searchfield
	oCExtJungsanSum.FRectSearchText = searchtext

	oCExtJungsanSum.FRectMimap = mimap
	oCExtJungsanSum.FRectMimapMinus = mimapminus

	oCExtJungsanSum.FRectVatYn = vatyn
	oCExtJungsanSum.FRectReturnOnly = retonly
	oCExtJungsanSum.FRectErrexists = errexists
	oCExtJungsanSum.FRectExistsBigo = existsBigo
	oCExtJungsanSum.FRectExceptItemCostZero = exceptcost0
   	oCExtJungsanSum.GetExtJungsanSum
%>
<p style="padding-top:10px;">
<table width="1500" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"  >
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�����</td>
		<td>�����</td>
		<td>�Ǽ�</td>
		<td>����</td>
		<td>�Ǹűݾ�</td>
    	<td>���޺δ�<br>����</td>
    	<td>���ٺδ�<br>����</td>
		<td>����ݾ�</td>
		<td>������</td>
		<td>��Ÿ����</td>
		<td>����ݾ�</td>
		<td>�̸��ΰǼ�</td>
		<td>��޾�</td>
		<td>���<br/>(�������ݾ� - ����������)</td>
		<td>���<br/>(�������ݾ� - ��޾�)</td>
	</tr>
	<%
	for i=0 to oCExtJungsanSum.FresultCount -1
		totROWno   				= totROWno + oCExtJungsanSum.FItemList(i).FROWno
		totSumItemno 			= totSumItemno + oCExtJungsanSum.FItemList(i).FSumItemno
		totSumitemcost 			= totSumitemcost + oCExtJungsanSum.FItemList(i).FSumitemcost
		totSumOwnCouponPrice	= totSumOwnCouponPrice + oCExtJungsanSum.FItemList(i).FSumOwnCouponPrice
		totSumTenCouponPrice	= totSumTenCouponPrice + oCExtJungsanSum.FItemList(i).FSumTenCouponPrice
		totSumMeachulPrice		= totSumMeachulPrice + oCExtJungsanSum.FItemList(i).FSumMeachulPrice
		totSumCommPrice			= totSumCommPrice + oCExtJungsanSum.FItemList(i).FSumCommPrice
		totSumJungsanPrice_ETC	= totSumJungsanPrice_ETC + oCExtJungsanSum.FItemList(i).FSumJungsanPrice_ETC
		totSumJungsanPrice		= totSumJungsanPrice + oCExtJungsanSum.FItemList(i).FSumJungsanPrice
		totMiMapTTLCnt			= totMiMapTTLCnt + oCExtJungsanSum.FItemList(i).FMiMapTTLCnt
		totReducedPrice			= totReducedPrice + oCExtJungsanSum.FItemList(i).FSumtenReducedPrice
		totBigoSum1				= totBigoSum1 + (oCExtJungsanSum.FItemList(i).FSumMeachulPrice - oCExtJungsanSum.FItemList(i).FSumReducedPrice)
		totBigoSum2				= totBigoSum2 + (oCExtJungsanSum.FItemList(i).FSumMeachulPrice - oCExtJungsanSum.FItemList(i).FSumtenReducedPrice)
'		totBigoSum				= totBigoSum + oCExtJungsanSum.FItemList(i).FBigoSum
'		totBigoBeasongSum		= totBigoBeasongSum + oCExtJungsanSum.FItemList(i).FBigoBeasongSum
	%>
	<tr bgcolor="FFFFFF" align="right">
		<td align="center" ><%=oCExtJungsanSum.FItemList(i).FextMeachulMonth%></td>
		<td align="center" ><%= CHKIIF(oCExtJungsanSum.FItemList(i).FjungsanfixMonth = "", "�̸�Ī", oCExtJungsanSum.FItemList(i).FjungsanfixMonth) %></td>
		<td align="center" ><%=FormatNumber(oCExtJungsanSum.FItemList(i).FROWno, FormatDotNo)%></td>
		<td align="center" ><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumItemno, FormatDotNo)%></td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumitemcost, FormatDotNo)%></td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumOwnCouponPrice, FormatDotNo)%></td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumTenCouponPrice, FormatDotNo)%></td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumMeachulPrice, FormatDotNo)%>
		<% if (oCExtJungsanSum.FItemList(i).GetDiffMeachulPrice<>0) then %>
			<br>(<font color="red"><%=formatNumber(oCExtJungsanSum.FItemList(i).GetDiffMeachulPrice,FormatDotNo)%></font>)
		<% end if %>
		</td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumCommPrice, FormatDotNo)%></td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumJungsanPrice_ETC,FormatDotNo)%>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumJungsanPrice,FormatDotNo)%>
		<% if (oCExtJungsanSum.FItemList(i).GetDiffJungsanPrice<>0) then %>
			<br>(<font color="red"><%=formatNumber(oCExtJungsanSum.FItemList(i).GetDiffJungsanPrice,FormatDotNo)%></font>)
		<% end if %>
		</td>
		<td align="center" ><%=FormatNumber(oCExtJungsanSum.FItemList(i).FMiMapTTLCnt, FormatDotNo)%></td>
		<td><%=FormatNumber(oCExtJungsanSum.FItemList(i).FSumtenReducedPrice, FormatDotNo)%></td>
		<td>
			<% 'response.write FormatNumber(oCExtJungsanSum.FItemList(i).FBigoSum, FormatDotNo) %>
			<% response.write FormatNumber(oCExtJungsanSum.FItemList(i).FSumMeachulPrice - oCExtJungsanSum.FItemList(i).FSumReducedPrice, FormatDotNo) %>
		</td>
		<td>
			<% 'response.write FormatNumber(oCExtJungsanSum.FItemList(i).FBigoBeasongSum, FormatDotNo) %>
			<% response.write FormatNumber(oCExtJungsanSum.FItemList(i).FSumMeachulPrice - oCExtJungsanSum.FItemList(i).FSumtenReducedPrice, FormatDotNo) %>
		</td>
	</tr>
	<% Next %>
	<tr bgcolor="FFFFFF" align="right">
		<td colspan="2" align="center">�հ�</td>
		<td bgcolor="#E6B9B8" align="center"><%= FormatNumber(totROWno, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8" align="center"><%= FormatNumber(totSumItemno, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumitemcost, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumOwnCouponPrice, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumTenCouponPrice, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumMeachulPrice, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumCommPrice, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumJungsanPrice_ETC, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totSumJungsanPrice, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8" align="center"><%= FormatNumber(totMiMapTTLCnt, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(totReducedPrice, FormatDotNo) %></td>
		<td bgcolor="#E6B9B8">
			<% ' response.write FormatNumber(totBigoSum, FormatDotNo) %>
			<% response.write FormatNumber(totBigoSum1, FormatDotNo) %>
		</td>
		<td bgcolor="#E6B9B8">
			<% 'response.write FormatNumber(totBigoBeasongSum, FormatDotNo) %>
			<% response.write FormatNumber(totBigoSum2, FormatDotNo) %>
		</td>
	</tr>
</table>
<br><br>
<p style="padding-top:20px;">
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
<form name="frmAct">
<tr>
	<td align="left">
		<input type="button" class="button" value="����ϱ�" onClick="popExtSiteJungsanData();">
	</td>
	<td align="right">
		<input type="button" class="button" value="����Ȯ�����Է�" onClick="popJungsanFixData('<%=sellsite%>');">
		&nbsp;
		<input type="button" class="button" value="API�����˾�" onClick="popAPIJungsanData();">
		<!--
	    <input type="text" name="extjdate" value="<%=extjdate%>" size="10" maxlength="8">
	    &nbsp;
	    <input type="button" class="button" value="SSG API ���" onClick="popExtSiteJungsanByAPI('ssg',this);">

	    <input type="button" class="button" value="������ũ API ���" onClick="popExtSiteJungsanByAPI('interpark',this);">

	    <input type="button" class="button" value="lotteCom API ���" onClick="popExtSiteJungsanByAPI('lottecom',this);">
	    &nbsp;
	    <input type="button" class="button" value="lotteiMall API ���" onClick="popExtSiteJungsanByAPI('lotteimall',this);">
	    -->
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<%
set oCExtJungsanSum = nothing

Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 25
	oCExtJungsan.FCurrPage = page

	oCExtJungsan.FRectStartdate = fromDate
	oCExtJungsan.FRectEndDate = toDate
	oCExtJungsan.FRectJungsanfixdate = jungsanfixdate

	oCExtJungsan.FRectSellSite = sellsite
	oCExtJungsan.FRectJungsanType = jungsantype

	oCExtJungsan.FRectSearchField = searchfield
	oCExtJungsan.FRectSearchText = searchtext

	oCExtJungsan.FRectMimap = mimap
	oCExtJungsan.FRectMimapMinus = mimapminus

	oCExtJungsan.FRectVatYn = vatyn
	oCExtJungsan.FRectReturnOnly = retonly
	oCExtJungsan.FRectErrexists = errexists
	oCExtJungsan.FRectExistsBigo = existsBigo
	oCExtJungsan.FRectExceptItemCostZero = exceptcost0

	if (xjungsanchk<>"") then
		oCExtJungsan.FRectDiffType = xjungsanchk
		oCExtJungsan.GetExtJungsanCheckTargetList
	else
    	oCExtJungsan.GetExtJungsan
	end if

	'oCExtJungsan.FRectGroupGubun = "sellsite"
  'oCExtJungsan.GetExtJungsanStatistic
%>

<p  >
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= oCExtJungsan.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCExtJungsan.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">���޸�</td>
	<td width="80">��������</td>
	<td width="150">����<br>�ֹ���ȣ</td>
	<td width="60">����<br>�ֹ�����</td>
	<td width="150">����<br>���ֹ���ȣ</td>
	<td width="40">����</td>

	<td width="60">�ǸŰ�</td>
	<td width="60">���޺δ�<br>����</td>
	<td width="60">���ٺδ�<br>����</td>
	<td width="60">������</td>
	<td width="70"><b>����ݾ�</b></td>
	<td width="60">������</td>
	<td width="70">����ݾ�</td>
	<td width="70">��������</td>
	<td width="80">���ֹ���ȣ</td>
	<td width="100">��ǰ�ڵ�</td>
	<td width="60">�ɼ��ڵ�</td>
	<td width="60">��޾�</td>
	<td width="70">�����</td>
	<td width="70">�����</td>
	<td width="70">������</td>
	<td>���</td>
</tr>

<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCExtJungsan.FItemList(i).GetSellSiteName %></td>
	<td><%= oCExtJungsan.FItemList(i).FextMeachulDate %></td>
	<td><a href="#" onClick="popByExtorderserial('<%= oCExtJungsan.FItemList(i).FextOrderserial %>');return false;"><%= oCExtJungsan.FItemList(i).FextOrderserial %></a></td>
	<td><%= oCExtJungsan.FItemList(i).FextOrderserSeq %></td>
	<td><a href="#" onClick="popByExtorderserial('<%= oCExtJungsan.FItemList(i).FextOrgOrderserial %>');return false;"><%= oCExtJungsan.FItemList(i).FextOrgOrderserial %></a></td>
	<td><%= oCExtJungsan.FItemList(i).FextItemNo %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextItemCost, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextOwnCouponPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenCouponPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextReducedPrice, FormatDotNo) %></td>
	<td align="right"><b><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenMeachulPrice, FormatDotNo) %></b>
	<% if (oCExtJungsan.FItemList(i).GetDiffMeachulPrice<>0) then %>
		<br>(<font color="red"><%=formatNumber(oCExtJungsan.FItemList(i).GetDiffMeachulPrice,FormatDotNo)%></font>)
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextCommPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenJungsanPrice, FormatDotNo) %>
	<% if (oCExtJungsan.FItemList(i).GetDiffJungsanPrice<>0) then %>
		<br>(<font color="red"><%=formatNumber(oCExtJungsan.FItemList(i).GetDiffJungsanPrice,FormatDotNo)%></font>)
	<% end if %>
	</td>
	<td>
		<%=oCExtJungsan.FItemList(i).GetSusumargin%>
	</td>
	<td><%= oCExtJungsan.FItemList(i).FOrgOrderserial %></td>
	<td><%= oCExtJungsan.FItemList(i).Fitemid %></td>
	<td><%= oCExtJungsan.FItemList(i).Fitemoption %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FReducedprice, FormatDotNo) %></td>
	<td><%= oCExtJungsan.FItemList(i).Fbeasongdate %></td>
	<td><%= oCExtJungsan.FItemList(i).Fdlvfinishdt %></td>
	<td><%= oCExtJungsan.FItemList(i).Fjungsanfixdate %></td>
	<td>
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FMinusOrderserial) then %>
			<%= oCExtJungsan.FItemList(i).FMinusOrderserial %>
		<% end if %>

		<% if (oCExtJungsan.FItemList(i).GetDiffReducedPrice <> 0) then %>
			<% if NOT isNULL(oCExtJungsan.FItemList(i).FMinusOrderserial) then %><br><% end if %>
		<%= oCExtJungsan.FItemList(i).GetDiffReducedPrice %>
		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="25" align="center">
		<% if oCExtJungsan.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsan.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsan.StartScrollPage to oCExtJungsan.FScrollCount + oCExtJungsan.StartScrollPage - 1 %>
			<% if i>oCExtJungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>
<form name="editFixFrm" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="jungsanfixdateUpd">
<input type="hidden" name="sellsite" value="">
</form>
<%
set oCExtJungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
