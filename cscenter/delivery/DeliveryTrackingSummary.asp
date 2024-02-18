<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���Ӹ�
' Hieditor : 2019.05.20 eastone ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim page, i, j, k
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, basedate, fromdate, todate
dim songjangdiv, makerid, isupbea
dim grpbeasongdate, grpsongjangdiv, grpbrand
dim MijipHaExists, MiBeasongExists, etcdivinc, errchktype, errchksub
dim chulgodt2, songjangdiv2, makerid2, isupbea2, mibeatype2, errchksub2
dim research

page     = requestCheckVar(request("page"),10)
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1		= requestCheckVar(request("mm1"),2)
dd1		= requestCheckVar(request("dd1"),2)
yyyy2	= requestCheckVar(request("yyyy2"),4)
mm2		= requestCheckVar(request("mm2"),2)
dd2		= requestCheckVar(request("dd2"),2)
songjangdiv		= requestCheckVar(request("songjangdiv"),10)
research		= requestCheckVar(request("research"),3)
makerid			= requestCheckVar(request("makerid"),32)
isupbea         = requestCheckVar(request("isupbea"),10)
grpbeasongdate  = requestCheckVar(request("grpbeasongdate"),10)
grpsongjangdiv  = requestCheckVar(request("grpsongjangdiv"),10)
grpbrand        = requestCheckVar(request("grpbrand"),10)
MijipHaExists   = requestCheckVar(request("MijipHaExists"),10)
MiBeasongExists = requestCheckVar(request("MiBeasongExists"),10)
etcdivinc       = requestCheckVar(request("etcdivinc"),10)

chulgodt2  = requestCheckVar(request("chulgodt2"),24)
songjangdiv2   = requestCheckVar(request("songjangdiv2"),10)
makerid2   = requestCheckVar(request("makerid2"),32)
isupbea2   = requestCheckVar(request("isupbea2"),10)
mibeatype2 = requestCheckVar(request("mibeatype2"),10)
errchksub2 = requestCheckVar(request("errchksub2"),10)

errchktype  = requestCheckVar(request("errchktype"),10)
errchksub   = requestCheckVar(request("errchksub"),10)

''rw chulgodt2&"|"&songjangdiv2&"|"&makerid2&"|"&isupbea2&"|"&mibeatype2
''delayDelivOnly	= requestCheckVar(request("delayDelivOnly"),3)
'checkCnt		= requestCheckVar(request("checkCnt"),32)
'orderserial		= requestCheckVar(request("orderserial"),16)
'songjangno      = requestCheckVar(request("songjangno"),20)

If page = "" Then page = 1
If research = "" Then
	''delayDelivOnly = "Y"
	''checkCnt = "5"
end if

if (yyyy1="") then
	basedate = Left(CStr(DateAdd("d", -7, now())),7)+"-01"
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -1, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2)),10)

if (grpbeasongdate="") then grpbeasongdate="1"
if (grpsongjangdiv="") then grpsongjangdiv="0"
if (grpbrand="") then grpbrand="0"
if (MijipHaExists="") then MijipHaExists="0"
if (MiBeasongExists="") then MiBeasongExists="0"
if (etcdivinc="") then etcdivinc="0"


dim oDeliveryTrackSum
SET oDeliveryTrackSum = New CDeliveryTrack
'oDeliveryTrackSum.FCurrPage				= 1
'oDeliveryTrackSum.FPageSize				= 100
oDeliveryTrackSum.FRectStartDate		= fromdate
oDeliveryTrackSum.FRectEndDate			= todate

oDeliveryTrackSum.FRectGrpBeasongdate   = grpbeasongdate
oDeliveryTrackSum.FRectGrpSongjangDiv   = grpsongjangdiv
oDeliveryTrackSum.FRectGrpBrand         = grpbrand
oDeliveryTrackSum.FRectSongjangDiv      = songjangdiv
oDeliveryTrackSum.FRectMakerid          = makerid
oDeliveryTrackSum.FRectisUpchebeasong   = isupbea
oDeliveryTrackSum.FRectMijipHaExists    = MijipHaExists
oDeliveryTrackSum.FRectMiBeasongExists  = MiBeasongExists
oDeliveryTrackSum.FRectEtcdivinc        = etcdivinc
if (errchktype="999") then
    oDeliveryTrackSum.FRectErrChkType       = CHKIIF((errchksub=""),errchktype,errchksub)
end if

oDeliveryTrackSum.getDeliveryTrackSummary()


dim oDeliveryTrackSum2
SET oDeliveryTrackSum2 = New CDeliveryTrack

if (chulgodt2<>"") then
    ''chulgodt2, songjangdiv2, makerid2, isupbea2, mibeatype2
    if InStr(chulgodt2,"~")>0 then
        oDeliveryTrackSum2.FRectStartDate		= SplitValue(chulgodt2,"~",0) 
        oDeliveryTrackSum2.FRectEndDate			= SplitValue(chulgodt2,"~",1) 
    else
        oDeliveryTrackSum2.FRectStartDate		= chulgodt2
        oDeliveryTrackSum2.FRectEndDate			= chulgodt2
    end if
    oDeliveryTrackSum2.FRectGrpBeasongdate   = grpbeasongdate
    oDeliveryTrackSum2.FRectGrpSongjangDiv   = "1"
    oDeliveryTrackSum2.FRectGrpBrand         = "1"
    oDeliveryTrackSum2.FRectSongjangDiv      = songjangdiv2
    oDeliveryTrackSum2.FRectMakerid          = makerid2
    oDeliveryTrackSum2.FRectisUpchebeasong       = isupbea2

    if (mibeatype2="1") then
        oDeliveryTrackSum2.FRectMijipHaExists    = "1"
        oDeliveryTrackSum2.FRectMiBeasongExists  = "0"
    elseif (mibeatype2="2") then
        oDeliveryTrackSum2.FRectMijipHaExists    = "0"
        oDeliveryTrackSum2.FRectMiBeasongExists  = "1"
    elseif (mibeatype2="999") then
        oDeliveryTrackSum2.FRectMijipHaExists    = "0"
        oDeliveryTrackSum2.FRectMiBeasongExists  = "0"
        oDeliveryTrackSum2.FRectErrChkType = CHKIIF((errchksub2=""),"999",errchksub2)  
    end if
    oDeliveryTrackSum2.FRectEtcdivinc        = etcdivinc

    oDeliveryTrackSum2.getDeliveryTrackSummary()

end if

Dim ttlchulgono, jiphafinCNT, dlvfinCNT, DminusCnt, Dplus0Cnt, Dplus1Cnt, Dplus2Cnt, Dplus3UpCnt
Dim MijiphaCnt, MidlvfinCnt, mijipHaPro, MiBeasongPro, errchkcnt

Dim ttlchulgono2, jiphafinCNT2, dlvfinCNT2, DminusCnt2, Dplus0Cnt2, Dplus1Cnt2, Dplus2Cnt2, Dplus3UpCnt2
Dim MijiphaCnt2, MidlvfinCnt2, mijipHaPro2, MiBeasongPro2, errchkcnt2

%>
<script language='javascript'>
function twoDepthSearch(chulgodt2,songjangdiv2,makerid2,isupbea2,mibeatype2,errchksub2){
    document.frm.chulgodt2.value=chulgodt2;
    document.frm.songjangdiv2.value=songjangdiv2;
    document.frm.makerid2.value=makerid2;
    document.frm.isupbea2.value=isupbea2;
    document.frm.mibeatype2.value=mibeatype2;
    document.frm.errchksub2.value=errchksub2;

    frm.submit();
}

function threeDepthSearch(chulgodtrng,songjangdiv2,makerid2,isupbea2,mibeatype2,errchksub2){
    var uri = "DeliveryTrackingSummaryDetail.asp?chulgodtrng="+chulgodtrng+"&songjangdiv="+songjangdiv2+"&makerid="+makerid2+"&isupbea="+isupbea2+"&mibeatype="+mibeatype2+"&etcdivinc=<%=etcdivinc%>"+"&errchksub="+errchksub2;
    var popwin = window.open(uri,'DeliveryTrackingSummaryDetail','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
}

function jsSubmit(frm) {
    document.frm.chulgodt2.value="";
    document.frm.songjangdiv2.value="";
    document.frm.makerid2.value="";
    document.frm.isupbea2.value="";
    document.frm.mibeatype2.value="";
    document.frm.errchksub2.value="";

	frm.submit();
}

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

var ptblrow;
function chgrowcolor(obj){
	obj.parentElement.style.background = "#FCE6E0";
    if ((ptblrow)&&(ptblrow.parentElement)){
        ptblrow.parentElement.style.background = "#FFFFFF";
    }
    ptblrow=obj;
}


</script>
<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<input type="hidden" name="chulgodt2" value="<%=chulgodt2%>">
<input type="hidden" name="songjangdiv2" value="<%=songjangdiv2%>">
<input type="hidden" name="makerid2" value="<%=makerid2%>">
<input type="hidden" name="isupbea2" value="<%=isupbea2%>">
<input type="hidden" name="mibeatype2" value="<%=mibeatype2%>">
<input type="hidden" name="errchksub2" value="<%=errchksub2%>">


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�����Է���(�����) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        &nbsp;
        ���豸�� :
        <select name="isupbea" >
            <option value="">��ü
            <option value="N" <%=CHKIIF(isupbea="N","selected","")%> >�ٹ�
            <option value="Y" <%=CHKIIF(isupbea="Y","selected","")%> >����
        </select>
		&nbsp;
		�ù�� :
        <% Call drawTrackDeliverBox("songjangdiv",songjangdiv,"Y") %>
		
		&nbsp;
		�귣��ID : <input type="text" class="text" name="makerid" value="<%= makerid %>">
		&nbsp;


        <% if (FALSE) then %>
		�����ȣ : <input type="text" class="text" name="songjangno" value="<%= songjangno %>">
        &nbsp;
		�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>">
        
        ��ȸCNT :
		<select class="select" name="checkCnt">
			<option></option>
			<option value="1" <%= CHKIIF(checkCnt="1", "selected", "") %> >1ȸ�̻�</option>
			<option value="2" <%= CHKIIF(checkCnt="2", "selected", "") %> >2ȸ�̻�</option>
			<option value="3" <%= CHKIIF(checkCnt="3", "selected", "") %> >3ȸ�̻�</option>
			<option value="4" <%= CHKIIF(checkCnt="4", "selected", "") %> >4ȸ�̻�</option>
			<option value="5" <%= CHKIIF(checkCnt="5", "selected", "") %> >5ȸ</option>
		</select>
        <% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        ����Ϻ��׷���
        <input type="radio" name="grpbeasongdate" value="0" <%=CHKIIF(grpbeasongdate="0","checked","")%> >�հ�
        <input type="radio" name="grpbeasongdate" value="1" <%=CHKIIF(grpbeasongdate="1","checked","")%> >����Ϻ�
        &nbsp;|&nbsp;
        �ù�纰�׷���
        <input type="radio" name="grpsongjangdiv" value="0" <%=CHKIIF(grpsongjangdiv="0","checked","")%> >�հ�
        <input type="radio" name="grpsongjangdiv" value="1" <%=CHKIIF(grpsongjangdiv="1","checked","")%> >�ù�纰
        &nbsp;|&nbsp;
        �귣�庰�׷���
        <input type="radio" name="grpbrand" value="0" <%=CHKIIF(grpbrand="0","checked","")%> >�հ�
        <input type="radio" name="grpbrand" value="1" <%=CHKIIF(grpbrand="1","checked","")%> >�귣�庰

        &nbsp;|&nbsp;
        ��Ÿ�ù�����
        <input type="radio" name="etcdivinc" value="0" <%=CHKIIF(etcdivinc="0","checked","")%> >��ü
        <input type="radio" name="etcdivinc" value="1" <%=CHKIIF(etcdivinc="1","checked","")%> >��Ÿ/�� ����
        <input type="radio" name="etcdivinc" value="2" <%=CHKIIF(etcdivinc="2","checked","")%> >��Ÿ/�� �� �˻�
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        
        <input type="checkbox" name="MijipHaExists" value="1" <%= CHKIIF(MijipHaExists="1", "checked", "") %> > ����������� �˻�
        &nbsp; or &nbsp;
        <input type="checkbox" name="MiBeasongExists" value="1" <%= CHKIIF(MiBeasongExists="1", "checked", "") %> > �̹������� �˻�
        &nbsp; | &nbsp;
        <input type="checkbox" name="errchktype" value="999" <%= CHKIIF(errchktype="999", "checked", "") %> > ERR�����
        <select name="errchksub">
            <option value="">������ü
            <option value="1" <%=CHKIIF(errchksub="1","selected","")%> >�������ʿ�
            <option value="2" <%=CHKIIF(errchksub="2","selected","")%> >�����ϴٸ�
            <option value="3" <%=CHKIIF(errchksub="3","selected","")%> >�����<������
            <option value="4" <%=CHKIIF(errchksub="4","selected","")%> >Digitüũ(Ÿ�ù�翹��)
            <option value="5" <%=CHKIIF(errchksub="5","selected","")%> >Digitüũ(����,�ڵ�)
            <option value="9" <%=CHKIIF(errchksub="9","selected","")%> >�����ȣ���̿���
        </select>
    </td>
</tr>
</table>
</form>

<br>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(oDeliveryTrackSum.FTotalCount,0) %></b> (�ִ� 2,000��)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="150">�����</td>
	<td width="120">�ù��</td>
	<td width="120">�귣��ID</td>
	<td width="70">���豸��</td>
	<td width="70">�����Ǽ�</td>
	<td width="70">�����ϰǼ�</td>
    <td width="70">��ۿϷ�Ǽ�</td>
    <td width="70">D-N�Ϸ�</td>
    <td width="70">D+0�Ϸ�</td>
    <td width="70">D+1�Ϸ�</td>
    <td width="70">D+2�Ϸ�</td>
    <td width="70">D+3�Ϸ�</td>
    <td width="70">������</td>
    <td width="70">�̹��</td>
    <td width="70">ERR</td>
    <td width="70">������</td>
    <td width="70">�����</td>
    <td width="140">���</td>
</tr>
<% for i = 0 to (oDeliveryTrackSum.FResultCount - 1) %>
<%
ttlchulgono = ttlchulgono + oDeliveryTrackSum.FItemList(i).Fttlchulgono
jiphafinCNT = jiphafinCNT + oDeliveryTrackSum.FItemList(i).FjiphafinCNT
dlvfinCNT   = dlvfinCNT   + oDeliveryTrackSum.FItemList(i).FdlvfinCNT
DminusCnt   = DminusCnt   + oDeliveryTrackSum.FItemList(i).FDminusCnt
Dplus0Cnt   = Dplus0Cnt   + oDeliveryTrackSum.FItemList(i).FDplus0Cnt
Dplus1Cnt   = Dplus1Cnt   + oDeliveryTrackSum.FItemList(i).FDplus1Cnt
Dplus2Cnt   = Dplus2Cnt   + oDeliveryTrackSum.FItemList(i).FDplus2Cnt
Dplus3UpCnt = Dplus3UpCnt   + oDeliveryTrackSum.FItemList(i).FDplus3UpCnt
MijiphaCnt  = MijiphaCnt   + oDeliveryTrackSum.FItemList(i).FMijiphaCnt
MidlvfinCnt = MidlvfinCnt   + oDeliveryTrackSum.FItemList(i).FMidlvfinCnt
errchkcnt   = errchkcnt   + oDeliveryTrackSum.FItemList(i).Ferrchkcnt
if (ttlchulgono<>0) then
     MijipHaPro = FIX((1-MijiphaCnt*1.0/ttlchulgono)*100) 
     MiBeasongPro = FIX((1-MidlvfinCnt*1.0/ttlchulgono)*100) 
end if 

%>
<tr bgcolor="<%=CHKIIF(chulgodt2=oDeliveryTrackSum.FItemList(i).Fbeasongdate AND songjangdiv2=CStr(oDeliveryTrackSum.FItemList(i).Fsongjangdiv) AND makerid2=oDeliveryTrackSum.FItemList(i).Fmakerid AND isupbea2=oDeliveryTrackSum.FItemList(i).Fisupchebeasong,"EEEEEE","FFFFFF")%>" align="right">
    <td align="center"><%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %></td>
    <td align="center"><%=oDeliveryTrackSum.FItemList(i).getSongjangDivName %></td>
    <td align="center"><%=oDeliveryTrackSum.FItemList(i).Fmakerid %></td>
    <td align="center"><%=oDeliveryTrackSum.FItemList(i).getUpbeaGubunName %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).Fttlchulgono,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FjiphafinCNT,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FdlvfinCNT,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FDminusCnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FDplus0Cnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FDplus1Cnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FDplus2Cnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FDplus3UpCnt,0) %></td>
    <td><a href="#" onClick="twoDepthSearch('<%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum.FItemList(i).Fisupchebeasong %>','1','');return false;" ><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FMijiphaCnt,0) %></a></td>
    <td><a href="#" onClick="twoDepthSearch('<%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum.FItemList(i).Fisupchebeasong %>','2','');return false;" ><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FMidlvfinCnt,0) %></a></td>
    <td><a href="#" onClick="twoDepthSearch('<%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum.FItemList(i).Fisupchebeasong %>','999','<%=errchksub%>');return false;" ><%=FormatNumber(oDeliveryTrackSum.FItemList(i).FErrChkCnt,0) %></a></td>
    <td align="center"><%=oDeliveryTrackSum.FItemList(i).getMijipHaPro %></td>
    <td align="center"><%=oDeliveryTrackSum.FItemList(i).getMiBeasongPro %></td>
    <td align="center">
        <a href="#" onClick="chgrowcolor(this);threeDepthSearch('<%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum.FItemList(i).Fisupchebeasong %>','1','');return false;" >[������]</a>
        <a href="#" onClick="chgrowcolor(this);threeDepthSearch('<%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum.FItemList(i).Fisupchebeasong %>','2','');return false;" >[�̹��]</a>
        <a href="#" onClick="chgrowcolor(this);threeDepthSearch('<%=oDeliveryTrackSum.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum.FItemList(i).Fisupchebeasong %>','999','<%=errchksub%>');return false;" >[ERR]</a>
    </td>
</tr>
<% next %>
<tr align="right" bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">�հ�</td>
    <td></td>
    <td></td>
    <td></td>
    <td><%=FormatNumber(ttlchulgono,0)%></td>
    <td><%=FormatNumber(jiphafinCNT,0)%></td>
    <td><%=FormatNumber(dlvfinCNT,0)%></td>
    <td><%=FormatNumber(DminusCnt,0)%></td>
    <td><%=FormatNumber(Dplus0Cnt,0)%></td>
    <td><%=FormatNumber(Dplus1Cnt,0)%></td>
    <td><%=FormatNumber(Dplus2Cnt,0)%></td>
    <td><%=FormatNumber(Dplus3UpCnt,0)%></td>
    <td><a href="#" onClick="twoDepthSearch('<%=fromdate %>~<%= todate %>','<%=songjangdiv %>','<%=makerid %>','<%=isupbea %>','1','');return false;" ><%=FormatNumber(MijiphaCnt,0)%></a></td>
    <td><a href="#" onClick="twoDepthSearch('<%=fromdate %>~<%= todate %>','<%=songjangdiv %>','<%=makerid %>','<%=isupbea %>','2','');return false;" ><%=FormatNumber(MidlvfinCnt,0)%></a></td>
    <td><a href="#" onClick="twoDepthSearch('<%=fromdate %>~<%= todate %>','<%=songjangdiv %>','<%=makerid %>','<%=isupbea %>','999','<%=errchksub%>');return false;" ><%=FormatNumber(errchkcnt,0)%></a></td>
    <td><%=mijipHaPro%> %</td>
    <td><%=MiBeasongPro %> %</td>
    <td align="center">
        <a href="#" onClick="threeDepthSearch('<%=fromdate %>~<%= todate %>','<%=songjangdiv %>','<%=makerid %>','<%=isupbea %>','1','');return false;" >[������]</a>
        <a href="#" onClick="threeDepthSearch('<%=fromdate %>~<%= todate %>','<%=songjangdiv %>','<%=makerid %>','<%=isupbea %>','2','');return false;" >[�̹��]</a>
        <a href="#" onClick="threeDepthSearch('<%=fromdate %>~<%= todate %>','<%=songjangdiv %>','<%=makerid %>','<%=isupbea %>','999','<%=errchksub%>');return false;" >[ERR]</a>
    </td>
</tr>
</table>

<br>
<p />
<% if (oDeliveryTrackSum2.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(oDeliveryTrackSum2.FTotalCount,0) %></b> (�ִ� 2,000��)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="150">�����</td>
	<td width="120">�ù��</td>
	<td width="120">�귣��ID</td>
	<td width="70">���豸��</td>
	<td width="70">�����Ǽ�</td>
	<td width="70">�����ϰǼ�</td>
    <td width="70">��ۿϷ�Ǽ�</td>
    <td width="70">D-N�Ϸ�</td>
    <td width="70">D+0�Ϸ�</td>
    <td width="70">D+1�Ϸ�</td>
    <td width="70">D+2�Ϸ�</td>
    <td width="70">D+3�Ϸ�</td>
    <td width="70">������</td>
    <td width="70">�̹��</td>
    <td width="70">ERR</td>
    <td width="70">������</td>
    <td width="70">�����</td>
    <td width="140">���</td>
</tr>
<% for i = 0 to (oDeliveryTrackSum2.FResultCount - 1) %>
<%
ttlchulgono2 = ttlchulgono2 + oDeliveryTrackSum2.FItemList(i).Fttlchulgono
jiphafinCNT2 = jiphafinCNT2 + oDeliveryTrackSum2.FItemList(i).FjiphafinCNT
dlvfinCNT2   = dlvfinCNT2   + oDeliveryTrackSum2.FItemList(i).FdlvfinCNT
DminusCnt2   = DminusCnt2  + oDeliveryTrackSum2.FItemList(i).FDminusCnt
Dplus0Cnt2   = Dplus0Cnt2   + oDeliveryTrackSum2.FItemList(i).FDplus0Cnt
Dplus1Cnt2   = Dplus1Cnt2   + oDeliveryTrackSum2.FItemList(i).FDplus1Cnt
Dplus2Cnt2   = Dplus2Cnt2   + oDeliveryTrackSum2.FItemList(i).FDplus2Cnt
Dplus3UpCnt2 = Dplus3UpCnt2   + oDeliveryTrackSum2.FItemList(i).FDplus3UpCnt
MijiphaCnt2  = MijiphaCnt2   + oDeliveryTrackSum2.FItemList(i).FMijiphaCnt
MidlvfinCnt2 = MidlvfinCnt2   + oDeliveryTrackSum2.FItemList(i).FMidlvfinCnt
errchkcnt2   = errchkcnt2   + oDeliveryTrackSum2.FItemList(i).Ferrchkcnt

if (ttlchulgono2<>0) then
     MijipHaPro2   = FIX((1-MijiphaCnt2*1.0/ttlchulgono2)*100) 
     MiBeasongPro2 = FIX((1-MidlvfinCnt2*1.0/ttlchulgono2)*100) 
end if 

%>
<tr bgcolor="FFFFFF" align="right">
    <td align="center"><%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %></td>
    <td align="center"><%=oDeliveryTrackSum2.FItemList(i).getSongjangDivName %></td>
    <td align="center"><%=oDeliveryTrackSum2.FItemList(i).Fmakerid %></td>
    <td align="center"><%=oDeliveryTrackSum2.FItemList(i).getUpbeaGubunName %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).Fttlchulgono,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FjiphafinCNT,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FdlvfinCNT,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FDminusCnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FDplus0Cnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FDplus1Cnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FDplus2Cnt,0) %></td>
    <td><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FDplus3UpCnt,0) %></td>
    <td><a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum2.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum2.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum2.FItemList(i).Fisupchebeasong %>','1','');return false;" ><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FMijiphaCnt,0) %></a></td>
    <td><a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum2.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum2.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum2.FItemList(i).Fisupchebeasong %>','2','');return false;" ><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FMidlvfinCnt,0) %></a></td>
    <td><a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum2.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum2.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum2.FItemList(i).Fisupchebeasong %>','999','<%=errchksub2%>');return false;" ><%=FormatNumber(oDeliveryTrackSum2.FItemList(i).FErrChkCnt,0) %></a></td>
    <td align="center"><%=oDeliveryTrackSum2.FItemList(i).getMijipHaPro %></td>
    <td align="center"><%=oDeliveryTrackSum2.FItemList(i).getMiBeasongPro %></td>
    <td align="center">
        <a href="#" onClick="chgrowcolor(this);threeDepthSearch('<%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum2.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum2.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum2.FItemList(i).Fisupchebeasong %>','1','');return false;" >[������]</a>
        <a href="#" onClick="chgrowcolor(this);threeDepthSearch('<%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum2.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum2.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum2.FItemList(i).Fisupchebeasong %>','2','');return false;" >[�̹��]</a>
        <a href="#" onClick="chgrowcolor(this);threeDepthSearch('<%=oDeliveryTrackSum2.FItemList(i).Fbeasongdate %>','<%=oDeliveryTrackSum2.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackSum2.FItemList(i).Fmakerid %>','<%=oDeliveryTrackSum2.FItemList(i).Fisupchebeasong %>','999','<%=errchksub2%>');return false;" >[ERR]</a>
    </td>
</tr>
<% next %>
<tr align="right" bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">�հ�</td>
    <td></td>
    <td></td>
    <td></td>
    <td><%=FormatNumber(ttlchulgono2,0)%></td>
    <td><%=FormatNumber(jiphafinCNT2,0)%></td>
    <td><%=FormatNumber(dlvfinCNT2,0)%></td>
    <td><%=FormatNumber(DminusCnt2,0)%></td>
    <td><%=FormatNumber(Dplus0Cnt2,0)%></td>
    <td><%=FormatNumber(Dplus1Cnt2,0)%></td>
    <td><%=FormatNumber(Dplus2Cnt2,0)%></td>
    <td><%=FormatNumber(Dplus3UpCnt2,0)%></td>
    <td><a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FRectStartDate %>~<%=oDeliveryTrackSum2.FRectEndDate %>','<%=songjangdiv2 %>','<%=makerid2 %>','<%=isupbea2 %>','1','');return false;" ><%=FormatNumber(MijiphaCnt2,0)%></a></td>
    <td><a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FRectStartDate %>~<%=oDeliveryTrackSum2.FRectEndDate %>','<%=songjangdiv2 %>','<%=makerid2 %>','<%=isupbea2 %>','2','');return false;" ><%=FormatNumber(MidlvfinCnt2,0)%></a></td>
    <td><a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FRectStartDate %>~<%=oDeliveryTrackSum2.FRectEndDate %>','<%=songjangdiv2 %>','<%=makerid2 %>','<%=isupbea2 %>','999','<%=errchksub2%>');return false;" ><%=FormatNumber(errchkcnt2,0)%></a></td>
    <td><%=mijipHaPro2%> %</td>
    <td><%=MiBeasongPro2 %> %</td>
    <td align="center">
        <a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FRectStartDate %>~<%=oDeliveryTrackSum2.FRectEndDate %>','<%=songjangdiv2 %>','<%=makerid2 %>','<%=isupbea2 %>','1','');return false;" >[������]</a>
        <a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FRectStartDate %>~<%=oDeliveryTrackSum2.FRectEndDate %>','<%=songjangdiv2 %>','<%=makerid2 %>','<%=isupbea2 %>','2','');return false;" >[�̹��]</a>
        <a href="#" onClick="threeDepthSearch('<%=oDeliveryTrackSum2.FRectStartDate %>~<%=oDeliveryTrackSum2.FRectEndDate %>','<%=songjangdiv2 %>','<%=makerid2 %>','<%=isupbea2 %>','999','<%=errchksub2%>');return false;" >[ERR]</a>
    </td>
</tr>
</table>

<p />
<% end if %>


<%
SET oDeliveryTrackSum = Nothing
SET oDeliveryTrackSum2 = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
