<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송추적 서머리 상세
' Hieditor : 2019.05.22 eastone 생성
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

<!-- 검색 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" height="60" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
    &nbsp;
		송장입력일(출고일) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        &nbsp;
        업배구분 :
        <select name="isupbea" >
            <option value="">전체
            <option value="N" <%=CHKIIF(isupbea="N","selected","")%> >텐배
            <option value="Y" <%=CHKIIF(isupbea="Y","selected","")%> >업배
        </select>
		&nbsp;
		택배사 :
        <% Call drawTrackDeliverBox("songjangdiv",songjangdiv,"Y") %>

	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    &nbsp;
		브랜드ID : <input type="text" class="text" name="makerid" value="<%= makerid %>">
		&nbsp;


		송장번호 : <input type="text" class="text" name="songjangno" value="<%= songjangno %>">
        &nbsp;
		주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    &nbsp;
        <input type="checkbox" name="MijipHaExists" value="1" <%= CHKIIF(MijipHaExists="1", "checked", "") %> > 미집하존재건 검색
        &nbsp; or &nbsp;
        <input type="checkbox" name="MiBeasongExists" value="1" <%= CHKIIF(MiBeasongExists="1", "checked", "") %> > 미배송존재건 검색
        &nbsp; | &nbsp;
        <input type="checkbox" name="errchktype" value="999" <%= CHKIIF(errchktype="999", "checked", "") %> > ERR예상건
        <select name="errchksub">
            <option value="">오류전체
            <option value="1" <%=CHKIIF(errchksub="1","selected","")%> >재추적필요
            <option value="2" <%=CHKIIF(errchksub="2","selected","")%> >추적일다름
            <option value="3" <%=CHKIIF(errchksub="3","selected","")%> >배송일<집하일
            <option value="4" <%=CHKIIF(errchksub="4","selected","")%> >Digit체크(타택배사예상)
            <option value="5" <%=CHKIIF(errchksub="5","selected","")%> >Digit체크(길이,코드)
            <option value="9" <%=CHKIIF(errchksub="9","selected","")%> >송장번호길이오류
        </select>

         &nbsp;&nbsp;|&nbsp;
        기타택배조건
        <input type="radio" name="etcdivinc" value="0" <%=CHKIIF(etcdivinc="0","checked","")%> >전체
        <input type="radio" name="etcdivinc" value="1" <%=CHKIIF(etcdivinc="1","checked","")%> >기타/퀵 제외
        <input type="radio" name="etcdivinc" value="2" <%=CHKIIF(etcdivinc="2","checked","")%> >기타/퀵 만 검색
	</td>
</tr>

</table>
</form>

<p />

* 배송완료일 산정 기준<br />
&nbsp;&nbsp; - 배송조회 불가한 경우 : 퀵서비스, 기타 => 출고일 +2일을 배송완료일<br />
&nbsp;&nbsp; - 배송조회 가능한 경우 : 배송조회를 해서 실제 배송완료일을 우리 배송완료일로 함, 14일이 지나도록 배송조회가 안되면 14일에 배송완료일로 함.<br />
&nbsp;&nbsp; - 반품완료가 있는 경우 : 이미 배송된걸로 해서 반품완료일을 배송완료일로 함.<br />
&nbsp;&nbsp; - CS내역이 있는 경우 : 고객 클래임이 있는걸로 간주하고 배송완료일 입력 안함.

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
        검색결과 : <b><%= FormatNumber(oDeliveryTrackList.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackList.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">주문번호</td>
    <td width="100">송장번호</td>
    <td width="120">택배사</td>
    <td width="100">Digit Chk</td>
    <td width="120">브랜드ID</td>

    <td width="90">출고일</td>

    <td width="90">집하일</td>
    <td width="90">배송완료일</td>
    <td width="120">최종추적일시</td>
    <% if (FALSE) then %>
    <td width="80">관련CS</td>
    <% end if %>
    <td width="60">ErrType</td>
    <td width="60">상태</td>
    <td width="120">추적</td>


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
        <a href="#" onClick="chgrowcolor(this);popDeliveryTrackingSummaryOne('<%=oDeliveryTrackList.FItemList(i).FOrderserial %>','<%=oDeliveryTrackList.FItemList(i).Fsongjangno %>','<%=oDeliveryTrackList.FItemList(i).Fsongjangdiv %>','<%=oDeliveryTrackList.FItemList(i).Fmakerid %>');return false;">[검토]</a>
    </td>
    <td align="center">
    <% if (oDeliveryTrackList.FItemList(i).isValidPopTraceSongjangDiv) then %>
    <a target="_dlv1" href="<%= oDeliveryTrackList.FItemList(i).getTrackURI %>">[택배사]</a>
    <% end if %>

    <% if (oDeliveryTrackList.FItemList(i).isValidPopTraceSongjangDiv) then %>
    <a target="_dlv2" href="<%= oDeliveryTrackList.FItemList(i).getTrackNaverURI %>">[네이버]</a>
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
