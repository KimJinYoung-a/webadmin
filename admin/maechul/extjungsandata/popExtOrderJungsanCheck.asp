<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 정산 Vs 주문내역
' Hieditor : 2018.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim i
dim research : research = requestCheckvar(request("research"),10)
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
dim page : page = requestCheckvar(request("page"),10)

dim difftp : difftp = requestCheckvar(request("difftp"),10)
dim chkerritemno : chkerritemno = requestCheckvar(request("chkerritemno"),10)

dim yyyy1, mm1
''dim fromDate, toDate, dlvyyyy, dlvmm
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
'dlvyyyy = requestCheckvar(request("dlvyyyy"),4)
'dlvmm = requestCheckvar(request("dlvmm"),2)

if (yyyy1="") then yyyy1=LEFT(NOW(),4)
if (mm1="") then mm1=MID(NOW(),6,2)
if (page="") then page=1

dim oCExtJungsanDiff
SET oCExtJungsanDiff = new CExtJungsan
oCExtJungsanDiff.FPageSize = 2000
oCExtJungsanDiff.FCurrPage = page
oCExtJungsanDiff.FRectSellSite = sellsite
oCExtJungsanDiff.FRectDlvMonth = yyyy1+"-"+mm1
oCExtJungsanDiff.FRectDiffType = difftp
if (chkerritemno<>"") then
oCExtJungsanDiff.FRectDiffType2 = 1
end if

oCExtJungsanDiff.getExtOrderJungsanDiffList

dim FormatDotNo : FormatDotNo=0
%>
<script language='javascript'>
function ssgDlvFinishSend(outmallorderserial,tenorderserial,tenitemid,tenitemoption){
	var params = "prctp=3&outmallorderserial="+outmallorderserial+"&tenorderserial="+tenorderserial+"&tenitemid="+tenitemid+"&tenitemoption="+tenitemoption+"&dlvfinishdt=2019-10-11"
 	var popwin=window.open('http://wapi.10x10.co.kr/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    
}

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sellsite%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function popJcomment(iorderserial,iitemid,iitemoption,isadd){
    var addcmt = "";
   // if (isadd){
        addcmt = prompt("정산 comment", "");
        if (addcmt == null) return;

        if (addcmt.length<1){
            alert("코멘트를 작성해주세요.");
            return;
        }

        var frm = document.frmcmt;
        frm.orderserial.value=iorderserial;
        frm.itemid.value=iitemid;
        frm.itemoption.value=iitemoption;
        frm.addcomment.value=addcmt;

        frm.submit();
   // }else{

   // }
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 제휴몰:
		<%= getJungsanXsiteComboHTML("sellsite",sellsite,"") %>
		&nbsp;
		
		* 출고월:
		<% DrawYMBox yyyy1,mm1 %>
        &nbsp;
        * 검색 조건
        <select class="select" name="difftp">
        <option value="" <%=CHKIIF(difftp="","selected","") %> >전체
        <option value="1" <%=CHKIIF(difftp="1","selected","") %> >자사기준 배송완료건만
		<option value="2" <%=CHKIIF(difftp="2","selected","") %> >송장변경 존재내역만
		<option value="3" <%=CHKIIF(difftp="3","selected","") %> >기타 택배사만 검색
        </select>
		&nbsp;
		<input type="checkbox" name="chkerritemno" <%=CHKIIF(chkerritemno<>"","checked","")%> >오차수량 있는 내역만

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>
		<%= getExtsongjangInputNOTIStr %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p  >
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oCExtJungsanDiff.FTotalcount %></b>
		&nbsp;
		<% if oCExtJungsanDiff.FTotalcount>=oCExtJungsanDiff.FPageSize then %>
        (최대 <%=FormatNumber(oCExtJungsanDiff.FPageSize,0)%> 건)
        <% end if %>
	</td>
</tr>

<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">

	<td width="80">제휴몰</td>
    <td width="50">주문수</td>
    <td width="50">교환수</td>
    <td width="50">반품수</td>

	<td width="90">주문번호</td>
	<td width="140">제휴 주문번호</td>
    <td width="80">상품코드</td>
    <td width="70">옵션코드</td>

	<td width="40">수량(합)</td>
	<td width="70">판매가(합)</td>
    <td width="70">실판매가(합)</td>

	<td width="70">출고월</td>
	<td width="70">오차수량</td>
	<td width="70">오차금액</td>

	<td width="80">택배사</td>
	<td width="90">송장번호</td>
    
	<td width="70">배송완료일</td>
	<td width="70">정산완료일</td>
	<td width="70">완료전송</td>
	<td>비고</td>

   
</tr>

<% if oCExtJungsanDiff.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td colspan="20">
        <% if (sellsite="") then %>
        [먼저 제휴몰을 선택 하세요.]
        <% else %>
        [검색결과가 없습니다.]
        <% end if %>
    </td>
</tr>
<% else %>
<% for i=0 to oCExtJungsanDiff.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCExtJungsanDiff.FItemList(i).Fsitename %></td>
	<td><%= oCExtJungsanDiff.FItemList(i).FordCnt %></td>
	<td><%= oCExtJungsanDiff.FItemList(i).FChgOrdCNT %></td>
	<td><%= oCExtJungsanDiff.FItemList(i).FretOrdCNT %></td>
	<td><a href="#" onClick="popDeliveryTrackingSummaryOne(<%= oCExtJungsanDiff.FItemList(i).ForgOrderserial %>,'<%= oCExtJungsanDiff.FItemList(i).Forgsongjangno %>',<%= oCExtJungsanDiff.FItemList(i).Forgsongjangdiv %>);return false;"><%= oCExtJungsanDiff.FItemList(i).ForgOrderserial %></a></td>
    <td><a href="#" onClick="popByExtorderserial('<%= oCExtJungsanDiff.FItemList(i).Fauthcode %>');return false;"><%= oCExtJungsanDiff.FItemList(i).Fauthcode %></a></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Fitemid %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Fitemoption %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Fitemno+oCExtJungsanDiff.FItemList(i).FMinus_itemno %></td>
    <td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FitemcostSum+oCExtJungsanDiff.FItemList(i).FMinus_itemcostSum,0) %></td>
    <td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FreducedpriceSum+oCExtJungsanDiff.FItemList(i).FMinus_reducedpriceSum,0) %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).FbeasongMonth %></td>
	<td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fjorgorderserial) then %>

		<% else %>
			<% if (oCExtJungsanDiff.FItemList(i).Fdiffitemno<>0) then %>
			<strong><%= FormatNumber(oCExtJungsanDiff.FItemList(i).Fdiffitemno,0) %></strong>
			<% else %>
			<%= FormatNumber(oCExtJungsanDiff.FItemList(i).Fdiffitemno,0) %>
			<% end if %>
		<% end if %>
	</td>
	<td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fjorgorderserial) then %>
		
		<% else %>
		<%= FormatNumber(oCExtJungsanDiff.FItemList(i).FdiffSum,0) %>
		<% end if %>
	</td>
    <td><%=getSongjangDiv2Val(oCExtJungsanDiff.FItemList(i).Forgsongjangdiv,1) %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Forgsongjangno %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Forgdlvfinishdt %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Forgjungsanfixdate %></td>
	
	<td>
		<% if (sellsite="ssg") then %>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fjorgorderserial) then %>
		<% if NOT isNULL(oCExtJungsanDiff.FItemList(i).Forgdlvfinishdt) or NOT isNULL(oCExtJungsanDiff.FItemList(i).Forgjungsanfixdate) then %>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fcomment) or InStr(oCExtJungsanDiff.FItemList(i).Fcomment,"완료")<1 then %>
		<a href="#" onClick="ssgDlvFinishSend('<%= oCExtJungsanDiff.FItemList(i).Fauthcode %>','<%=oCExtJungsanDiff.FItemList(i).ForgOrderserial%>','<%=oCExtJungsanDiff.FItemList(i).Fitemid%>','<%=oCExtJungsanDiff.FItemList(i).Fitemoption%>');return false;">[배송완료전송]</a>
		<% end if %>
		<% end if %>
		<% end if %>
		<% end if %>
	</td>
    <td>
        <% if isNULL(oCExtJungsanDiff.FItemList(i).Fcomment) or (oCExtJungsanDiff.FItemList(i).Fcomment="") then %>
            <a href="#" onClick="popJcomment('<%=oCExtJungsanDiff.FItemList(i).ForgOrderserial%>','<%=oCExtJungsanDiff.FItemList(i).Fitemid%>','<%=oCExtJungsanDiff.FItemList(i).Fitemoption%>',true);return false;"><img src="/images/icon_new.gif" alt="코멘트작성"></a>
        <% else %>
            <a href="#" onClick="popJcomment('<%=oCExtJungsanDiff.FItemList(i).ForgOrderserial%>','<%=oCExtJungsanDiff.FItemList(i).Fitemid%>','<%=oCExtJungsanDiff.FItemList(i).Fitemoption%>',false);return false;"><%=oCExtJungsanDiff.FItemList(i).Fcomment%></a>
        <% end if %>
    </td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
    <% if (FALSE) then %>
		<% if oCExtJungsanDiff.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsanDiff.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsanDiff.StartScrollPage to oCExtJungsanDiff.FScrollCount + oCExtJungsanDiff.StartScrollPage - 1 %>
			<% if i>oCExtJungsanDiff.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsanDiff.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    <% end if %>
	</td>
</tr>

</table>

<p>
<form name="frmcmt" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="addcmt">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="addcomment" value="">
</form>

<%
set oCExtJungsanDiff = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->

