<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���Ӹ� ��
' Hieditor : 2019.05.22 eastone ����
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
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, basedate, fromdate, todate, chulgodtrng
dim songjangdiv, makerid, isupbea, mibeatype, etcdivinc
dim MijipHaExists, MiBeasongExists, errchktype, errchksub
dim songjangno, orderserial

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

MijipHaExists   = requestCheckVar(request("MijipHaExists"),10)
MiBeasongExists = requestCheckVar(request("MiBeasongExists"),10)
etcdivinc       = requestCheckVar(request("etcdivinc"),10)
errchktype      = requestCheckVar(request("errchktype"),10)
errchksub       = requestCheckVar(request("errchksub"),10)

songjangno      = requestCheckVar(request("songjangno"),32)
orderserial     = requestCheckVar(request("orderserial"),11)

mibeatype = requestCheckVar(request("mibeatype"),10)
chulgodtrng = requestCheckVar(request("chulgodtrng"),21) '' 2019-05-01~2019-05-11

If page = "" Then page = 1
If research = "" Then
	''delayDelivOnly = "Y"
	''checkCnt = "5"
end if

if (yyyy1="") then
    if (chulgodtrng<>"") then
        if (InStr(chulgodtrng,"~")>0) then
            basedate    = SplitValue(chulgodtrng,"~",0)
        else
            basedate    = chulgodtrng
        end if
    else
        basedate = Left(CStr(DateAdd("d", -10, now())),10)
    end if

	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

    if (chulgodtrng<>"") then
        if (InStr(chulgodtrng,"~")>0) then
            basedate    = SplitValue(chulgodtrng,"~",1)
        else
            basedate    = chulgodtrng
        end if
    else
	    basedate = Left(CStr(DateAdd("d", -1, now())),10)
    end if

	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2)),10)


if (MijipHaExists="") then MijipHaExists="0"
if (MiBeasongExists="") then MiBeasongExists="0"
if (etcdivinc="") then etcdivinc="0"

if (isupbea="N") and (makerid="10x10logistics") then makerid=""

if (mibeatype="1") then
    MijipHaExists ="1"
    MiBeasongExists ="0"
elseif (mibeatype="2") then
    MijipHaExists ="0"
    MiBeasongExists ="1"
elseif (mibeatype="999") then
    MijipHaExists ="0"
    MiBeasongExists ="0"
    errchktype = "999"
end if

dim oDeliveryTrackList
SET oDeliveryTrackList = New CDeliveryTrack
oDeliveryTrackList.FCurrPage				= page
oDeliveryTrackList.FPageSize				= 100
oDeliveryTrackList.FRectStartDate		= fromdate
oDeliveryTrackList.FRectEndDate			= todate

oDeliveryTrackList.FRectSongjangDiv      = songjangdiv
oDeliveryTrackList.FRectMakerid          = makerid
oDeliveryTrackList.FRectisUpchebeasong   = isupbea
oDeliveryTrackList.FRectMijipHaExists    = MijipHaExists
oDeliveryTrackList.FRectMiBeasongExists  = MiBeasongExists
oDeliveryTrackList.FRectEtcdivinc        = etcdivinc

oDeliveryTrackList.FRectOrderserial      = orderserial
oDeliveryTrackList.FRectSongjangNo       = songjangno

if (errchktype="999") then
    oDeliveryTrackList.FRectErrChkType       = CHKIIF(errchksub="","999",errchksub)
end if

' if (oDeliveryTrackList.FRectErrChkType<>"") then
'     oDeliveryTrackList.getDeliveryTrackSummaryDetailRealTime
' else
    oDeliveryTrackList.getDeliveryTrackSummaryDetail()
'end if

%>
<script language="javascript">
function jsSubmit(frm) {
	frm.submit();
}

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv,imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv+"&makerid="+imakerid;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function poporderdetail(iorderserial){
    var iurl = "/cscenter/ordermaster/orderitemmaster.asp?orderserial="+iorderserial;
    var popwin = window.open(iurl,'poporderitemmaster','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

var ptblrow;
function chgrowcolor(obj){

	obj.parentElement.parentElement.style.background = "#FCE6E0";
    if ((ptblrow)&&(ptblrow.parentElement.parentElement)){
        ptblrow.parentElement.parentElement.style.background = "#FFFFFF";
    }
    ptblrow=obj;
}

function popDeliveryTrackingFakeBrandPop(imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingFake.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv=<%=songjangdiv%>&research=<%=research%>&orderserial=<%=orderserial%>&etcdivinc=<%=etcdivinc%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingFakepop','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    &nbsp;
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

	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    &nbsp;
		�귣��ID : <input type="text" class="text" name="makerid" value="<%= makerid %>">
		&nbsp;


		�����ȣ : <input type="text" class="text" name="songjangno" value="<%= songjangno %>">
        &nbsp;
		�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    &nbsp;
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

         &nbsp;&nbsp;|&nbsp;
        ��Ÿ�ù�����
        <input type="radio" name="etcdivinc" value="0" <%=CHKIIF(etcdivinc="0","checked","")%> >��ü
        <input type="radio" name="etcdivinc" value="1" <%=CHKIIF(etcdivinc="1","checked","")%> >��Ÿ/�� ����
        <input type="radio" name="etcdivinc" value="2" <%=CHKIIF(etcdivinc="2","checked","")%> >��Ÿ/�� �� �˻�
	</td>
</tr>

</table>
</form>

<p />

* ��ۿϷ��� ���� ����<br />
&nbsp;&nbsp; - �����ȸ �Ұ��� ��� : ������, ��Ÿ => ����� +2���� ��ۿϷ���<br />
&nbsp;&nbsp; - �����ȸ ������ ��� : �����ȸ�� �ؼ� ���� ��ۿϷ����� �츮 ��ۿϷ��Ϸ� ��, 14���� �������� �����ȸ�� �ȵǸ� 14�Ͽ� ��ۿϷ��Ϸ� ��.<br />
&nbsp;&nbsp; - ��ǰ�Ϸᰡ �ִ� ��� : �̹� ��۵Ȱɷ� �ؼ� ��ǰ�Ϸ����� ��ۿϷ��Ϸ� ��.<br />
&nbsp;&nbsp; - CS������ �ִ� ��� : �� Ŭ������ �ִ°ɷ� �����ϰ� ��ۿϷ��� �Է� ����.

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
        �˻���� : <b><%= FormatNumber(oDeliveryTrackList.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackList.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">�ֹ���ȣ</td>
    <td width="100">�����ȣ</td>
    <td width="120">�ù��</td>
    <td width="100">Digit Chk</td>
    <td width="120">�귣��ID</td>

    <td width="90">�����</td>

    <td width="90">������</td>
    <td width="90">��ۿϷ���</td>
    <td width="120">���������Ͻ�</td>
    <% if (FALSE) then %>
    <td width="80">����CS</td>
    <% end if %>
    <td width="60">ErrType</td>
    <td width="60">����</td>
    <td width="120">����</td>


</tr>
<% for i = 0 to (oDeliveryTrackList.FResultCount - 1) %>
<tr align="center" bgcolor="#FFFFFF">
    <td><a href="#" onClick="PopOrderMasterWithCallRingOrderserial('<%=oDeliveryTrackList.FItemList(i).FOrderserial %>');return false;"><%=oDeliveryTrackList.FItemList(i).FOrderserial %></a></td>
    <td><%=oDeliveryTrackList.FItemList(i).Fsongjangno %></td>
    <td><%=oDeliveryTrackList.FItemList(i).Fdivname %></td>
    <td><%=oDeliveryTrackList.FItemList(i).getDigitChkStr %></td>
    <td><a href="#" onClick="popDeliveryTrackingFakeBrandPop('<%=oDeliveryTrackList.FItemList(i).Fmakerid %>');return false;"><%=oDeliveryTrackList.FItemList(i).Fmakerid %></a></td>

    <td ><%=oDeliveryTrackList.FItemList(i).Fbeasongdate %></td>

    <td ><%=oDeliveryTrackList.FItemList(i).Fdeparturedt %></td>
    <td ><%=oDeliveryTrackList.FItemList(i).FdlvfinishDT %>
    <% if (FALSE) then %>
    <% if NOT isNULL(oDeliveryTrackList.FItemList(i).Ftrarrivedt) and (oDeliveryTrackList.FItemList(i).Ftrarrivedt<>"") then %>
        / <strong><%=oDeliveryTrackList.FItemList(i).Ftrarrivedt %></strong>
    <% end if %>
    <% end if %>
    </td>
    <td ><%=oDeliveryTrackList.FItemList(i).Ftraceupddt %></td>
    <% if (FALSE) then %>
    <td >
        <% if oDeliveryTrackList.FItemList(i).FcsCNT>0 or oDeliveryTrackList.FItemList(i).FcsFinCNT>0 then %>
            <% if oDeliveryTrackList.FItemList(i).FcsFinCNT>0 then %>
            <strong><%=oDeliveryTrackList.FItemList(i).FcsFinCNT%></strong>
            <% else %>
            <%=oDeliveryTrackList.FItemList(i).FcsFinCNT%>
            <% end if %>
            /
            <%=oDeliveryTrackList.FItemList(i).FcsCNT%>
        <% end if %>
    </td>
    <% end if %>
    <td ><%=oDeliveryTrackList.FItemList(i).getErrChkTypeName %></td>

    <td align="center">
        <a href="#" onClick="chgrowcolor(this);popDeliveryTrackingSummaryOne('<%=oDeliveryTrackList.FItemList(i).FOrderserial %>','<%=oDeliveryTrackList.FItemList(i).Fsongjangno %>','<%=oDeliveryTrackList.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackList.FItemList(i).Fmakerid %>');return false;">[����]</a>
    </td>
    <td align="center">
    <% if (oDeliveryTrackList.FItemList(i).isValidPopTraceSongjangDiv) then %>
    <a target="_dlv1" href="<%= oDeliveryTrackList.FItemList(i).getTrackURI %>">[�ù��]</a>
    <% end if %>

    <% if (oDeliveryTrackList.FItemList(i).isValidPopTraceSongjangDiv) then %>
    <a target="_dlv2" href="<%= oDeliveryTrackList.FItemList(i).getTrackNaverURI %>">[���̹�]</a>
    <% end if %>
    </td>
</tr>
<% next %>
<!--
<tr align="right" bgcolor="<%= adminColor("tabletop") %>">
    <td align="center"></td>
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
-->
<tr height="20">
    <td colspan="13" align="center" bgcolor="#FFFFFF">
        <% if oDeliveryTrackList.HasPreScroll then %>
        <a href="javascript:goPage('<%= oDeliveryTrackList.StartScrollPage-1 %>');">[pre]</a>
        <% else %>
            [pre]
        <% end if %>

        <% for i=0 + oDeliveryTrackList.StartScrollPage to oDeliveryTrackList.FScrollCount + oDeliveryTrackList.StartScrollPage - 1 %>
            <% if i>oDeliveryTrackList.FTotalpage then Exit for %>
            <% if CStr(page)=CStr(i) then %>
            <font color="red">[<%= i %>]</font>
            <% else %>
            <a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
            <% end if %>
        <% next %>

        <% if oDeliveryTrackList.HasNextScroll then %>
            <a href="javascript:goPage('<%= i %>');">[next]</a>
        <% else %>
            [next]
        <% end if %>
    </td>
</tr>
</table>

<br>
<p />
<%
SET oDeliveryTrackList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
