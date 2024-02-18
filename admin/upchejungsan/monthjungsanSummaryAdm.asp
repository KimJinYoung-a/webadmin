<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, yyyy1,mm1, jgubun, targetGbn, groupid, page, finishflag, taxtype, itemvatyn, comm_cd
page    = requestCheckvar(request("page"),10)
makerid = requestCheckvar(request("makerid"),32)
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)
jgubun  = requestCheckvar(request("jgubun"),10)
targetGbn= requestCheckvar(request("targetGbn"),10)
groupid  = requestCheckvar(request("groupid"),10)
finishflag = requestCheckvar(request("finishflag"),10)
taxtype   = requestCheckvar(request("taxtype"),10)
itemvatyn = requestCheckvar(request("itemvatyn"),10)
comm_cd = requestCheckvar(request("comm_cd"),16)

if (page="") then page=1

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

if (targetGbn="") then
    targetGbn = "ON"
end if

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FPageSize = 30
ojungsanTax.FCurrPage = page
ojungsanTax.FRectMakerid = makerid
ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTax.FRectJGubun = jgubun
ojungsanTax.FRectTargetGbn = targetGbn
ojungsanTax.FRectGroupid = groupid
ojungsanTax.FRectFinishFlag = finishflag
ojungsanTax.FRectTaxType = taxtype
ojungsanTax.FRectItemVatYn = itemvatyn
ojungsanTax.FRectcomm_cd = comm_cd
ojungsanTax.getMonthUpcheJungsanSummaryAdm


dim i
%>
<script type='text/javascript'>

function NextPage(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function PopDetail(iidx,tg,makerid){
    var uri = 'jungsandetailsumONAdm.asp?id=' + iidx;
    if (tg=="OF") uri = 'jungsandetailsumOFAdm.asp?idx=' + iidx;
	var popwin = window.open(uri+'&makerid='+makerid,'PopDetail','width=1300, height=900, scrollbars=1, resizable=yes');
	popwin.focus();
}

function PopTaxRegPrdCommission(makerid, yyyy1, mm1, onoffGubun, jidx) {
	var popwin = window.open("popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=640 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopTaxPrintReDirect(itax_no){
	var popwinsub = window.open("red_taxprint.asp?tax_no=" + itax_no ,"taxview","width=800,height=700,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

function PopConfirm(mnupos,iidx){
	//var popwin = window.open('jungsanmaster.asp?id=' + iidx + '&menupos=' + mnupos,'popshowdetail','width=900, height=540, scrollbars=1');
	//popwin.focus();
}

function PopTaxReg(v){
	//var popwin = window.open("poptaxreg.asp?id=" + v,"poptaxreg","width=640 height=700 scrollbars=yes resizable=yes");
	//popwin.focus();
}

function PopTaxRegOff(v){
	//var popwin = window.open("poptaxregoff.asp?idx=" + v,"poptaxregoff1","width=640 height=680 scrollbars=yes resizable=yes");
	//popwin.focus();
}

function XLDown(){

    var paramURL = 'monthjungsanAdmAllXL.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&makerid=<%=makerid%>&jgubun=<%=jgubun%>&targetGbn=<%=targetGbn%>&groupid=<%=groupid%>&finishflag=<%=finishflag%>&taxtype=<%=taxtype%>';

    var popwin = window.open(paramURL,'monthjungsanAdmAllXL','width=100,height=100,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if (ojungsanTax.FresultCount>0) then %>
//alert('2014�� 1�� ������� ������ ����п� ���ؼ���\n\n�ٹ����ٿ��� ��꼭�� ���� �Ͽ���\n\n�̼��� ���� ���� ���� �������� ���� �ֽñ� �ٶ��ϴ�.');
<% end if %>
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���� ��� ��� :&nbsp;<% DrawYMBox yyyy1,mm1 %>
		&nbsp;
		* �귣��ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;
        ��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        * �����ı��� :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
        &nbsp;
		* ��꼭��������
		<select name="taxtype" >
		<option value="">��ü
		<option value="01" <%= CHKIIF(taxtype="01","selected","") %> >����
		<option value="02" <%= CHKIIF(taxtype="02","selected","") %> >�鼼
		<option value="03" <%= CHKIIF(taxtype="03","selected","") %> >����
		</select>
        &nbsp;
        * ����ó ���� :
        <select name="targetGbn" >
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >OF
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>
		&nbsp;
		* ����
		<select name="finishflag" >
		<option value="">��ü
		<option value="0" <%= CHKIIF(finishflag="0","selected","") %> >������
		<option value="1" <%= CHKIIF(finishflag="1","selected","") %> >��üȮ�δ��
		<option value="2" <%= CHKIIF(finishflag="2","selected","") %> >��üȮ�οϷ�
		<option value="3" <%= CHKIIF(finishflag="3","selected","") %> >����Ȯ��
		<option value="7" <%= CHKIIF(finishflag="7","selected","") %> >�ԱݿϷ�
		</select>
		&nbsp;
		* ��ǰ��������
		<select name="itemvatyn" >
		<option value="">��ü
		<option value="Y" <%= CHKIIF(itemvatyn="Y","selected","") %> >����
		<option value="N" <%= CHKIIF(itemvatyn="N","selected","") %> >�鼼
		</select>
		&nbsp;
		* ����Ӽ����� : <% DrawJungsanGubun "comm_cd",comm_cd,"Z003","" %>
    </td>
</tr>
</form>
</table>
<p>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="100" >�����ı���</td>
    <td width="100" >����ó�ڵ�</td>
    <td width="100" >����ó</td>
    <td width="80" >��ǰ���� ����</td>
    <td width="90" >����Ӽ�����</td>
    <td width="90" >�ǸŰ�</td>
    <td width="90" >���Ǹűݾ�<br>(���»� �����)</td>
    <td width="80" >������</td>
    <td width="80" >��������<br>������</td>
  	<td width="80">�����</td>
</tr>

<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTax.FItemList(i).getJGubunName%></td>
    <td><%=ojungsanTax.FItemList(i).Fsitename%></td>
    <td ><%=ojungsanTax.FItemList(i).FsitenameName%></td>
    <td><%=ojungsanTax.FItemList(i).getItemVatTypeName%></td>
    <td ><%=ojungsanTax.FItemList(i).Fcomm_name%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FsellcashSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FreducedpriceSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FitemcommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPgcommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FsuplycashSum,0)%></td>


</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
</tr>
</table>

<%
set ojungsanTax = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
