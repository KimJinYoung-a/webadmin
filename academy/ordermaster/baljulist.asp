<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/order/acabaljucls.asp"-->
<%

'// 한글 한글 한글

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)

yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	date1 = dateAdd("d",-4,nowdate)
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(dateserial(yyyy2, mm2 , dd2)+1),10)
end if

dim baljumaster
baljumaster = request("baljumaster")

dim obalju,i
set obalju = New CAcademyDiyBalju
obalju.FStartdate = yyyy1 + "-" + mm1 + "-" + dd1
obalju.FEndDate = Left(CStr(Edate),10)
obalju.FMaxcount = 1000

obalju.getAcademyDiyBaljumaster

''발주상세내역
'if baljumaster<>"" then
'	obalju.getBaljuDetailList baljumaster
'end if

dim ppdate, ppcnt, sscnt, sumppcnt, sumsscnt, itemCnt, sumItemCnt
	sumppcnt = 0
	sumsscnt = 0
	sumItemCnt = 0

dim SubChulgoCount, Subdelay0chulgocnt, Subdelay1chulgocnt, Subdelay2chulgocnt, Subdelay3chulgocnt, SubCancelCnt, SubMichulgoCnt
dim SumChulgoCount, Sumdelay0chulgocnt, Sumdelay1chulgocnt, Sumdelay2chulgocnt, Sumdelay3chulgocnt, SumCancelCnt, SumMichulgoCnt

%>

<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function AnViewUpcheList(iid){
	var popwin = window.open('/admin/pop/viewupchelist.asp?iid=' + iid,'viewupchelist','width=780,height=700,scrollbars=yes');
	popwin.focus();
}

function ViewAddDetailList(){
    var popwin = window.open('/admin/ordermaster/pop_makeonorder_list.asp?research=on&menupos=44','pop_makeonorder_list','width=1100,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsPopLogisticsBaljuList(sitebaljukey) {
    var popwin = window.open("pop_logistics_baljuitemlist.asp?sitebaljukey=" + sitebaljukey,"jsPopLogisticsBaljuList","width=800,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function ViewBaljuArr(){
	var frm;
	var pass = false;
	var idxarr = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (idxarr==""){
					idxarr = frm.iidx.value;
				}else{
					idxarr = idxarr + "," + frm.iidx.value;
				}
			}
		}
	}

	window.open('/admin/pop/viewupchelist.asp?idxarr=' + idxarr,'','');
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#F4F4F4" >
	    <td rowspan="2" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
        <td align="left">
        	조회기간 : <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        </td>
        <td rowspan="2" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<!--
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
    <tr height="25">
        <td align="left">
        	<input type="button" class="button" value="선택발주아이템목록보기" onclick="ViewBaljuArr()">
        </td>
        <td align="right">
            <input type="button" class="button" value="주문제작List" onclick="ViewAddDetailList()">
        </td>
    </tr>
</table>
-->
<!-- 액션 끝 -->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=20></td>
		<td width="80">발주ID</td>
		<td>발주일시</td>
		<td>주문<br>사이트</td>
		<td>차수</td>
		<td>그룹</td>
		<td>발주<br>타입</td>
		<td>택배사</td>
		<td>총발주건</td>
		<td>(%)</td>
		<td>자체배송<br>건수</td>
		<td>상품수<br>(취소X)</td>
		<td>취소</td>
		<td>총(텐배)<br>출고건수</td>
		<td>당일<br>출고건수</td>
		<td>1일<br>지연출고</td>
		<td>2일<br>지연출고</td>
		<td>3일<br>지연출고</td>
		<td>미출고</td>
		<td>주문<br>리스트</td>
		<!--
		<td>상품<br>목록</td>
		<td>송장파일</td>
		-->
		<td>사은품<br>리스트</td>
		<td>사은품<br>합계</td>
		<td>텐배<br>주문제작</td>
		<td>발주서출력</td>
	</tr>
<% if (obalju.resultBaljucount<1) then %>
	<tr bgcolor="#FFFFFF" height="31"><td colspan="24" align="center">해당기간에 발주서없음</td></tr>
<% else %>

<% for i=0 to obalju.resultBaljucount-1 %>
<%
if ppdate<>Left(obalju.FBaljumasterList(i).FBaljudate,10) then
	ppdate = Left(obalju.FBaljumasterList(i).FBaljudate,10)
	ppcnt = 0
	sscnt = 0
	itemCnt = 0
    SubChulgoCount      = 0
    Subdelay0chulgocnt  = 0
    Subdelay1chulgocnt  = 0
    Subdelay2chulgocnt  = 0
    Subdelay3chulgocnt  = 0
    SubMichulgoCnt      = 0
    SubCancelCnt        = 0
end if

ppcnt = ppcnt + obalju.FBaljumasterList(i).FCount
sscnt = sscnt + obalju.FBaljumasterList(i).Fsongjangcnt
itemCnt = itemCnt + obalju.FBaljumasterList(i).Fitemno

SubChulgoCount      = SubChulgoCount +  obalju.FBaljumasterList(i).GetTotalChulgoCount
Subdelay0chulgocnt  = Subdelay0chulgocnt + obalju.FBaljumasterList(i).Fdelay0chulgocnt
Subdelay1chulgocnt  = Subdelay1chulgocnt + obalju.FBaljumasterList(i).Fdelay1chulgocnt
Subdelay2chulgocnt  = Subdelay2chulgocnt + obalju.FBaljumasterList(i).Fdelay2chulgocnt
Subdelay3chulgocnt  = Subdelay3chulgocnt + obalju.FBaljumasterList(i).Fdelay3chulgocnt
SubMichulgoCnt      = SubMichulgoCnt + obalju.FBaljumasterList(i).GetTenMiChulgoCount
SubCancelCnt        = SubCancelCnt + obalju.FBaljumasterList(i).FCancelCnt
%>


	<form name="frmBuyPrc_<%= obalju.FBaljumasterList(i).FBaljuID %>" method="post" >
	<input type="hidden" name="iidx" value="<%= obalju.FBaljumasterList(i).FBaljuID %>">
	<tr bgcolor="#FFFFFF" align="center">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= obalju.FBaljumasterList(i).FBaljuID %></td>
		<td align="left"><%= obalju.FBaljumasterList(i).FBaljudate %></td>
		<td><%= obalju.FBaljumasterList(i).GetExtSiteName %></td>
		<td><%= obalju.FBaljumasterList(i).Fdifferencekey %></td>
		<td><%= obalju.FBaljumasterList(i).Fworkgroup %></td>
		<td><%= obalju.FBaljumasterList(i).getBaljuTypeName %></td>
		<td><%= obalju.FBaljumasterList(i).getDeliverName %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).FCount,0) %></td>
		<td>
		<% if obalju.FBaljumasterList(i).FCount<>0 then %>
			<%= CLng(obalju.FBaljumasterList(i).Fsongjangcnt/obalju.FBaljumasterList(i).FCount*100) %> %
		<% end if %>
        </td>
        <td><%= FormatNumber(obalju.FBaljumasterList(i).Fsongjangcnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fitemno,0) %></td>
    	<td><%= FormatNumber(obalju.FBaljumasterList(i).Fcancelcnt,0) %></td>
	    <td><%= FormatNumber(obalju.FBaljumasterList(i).GetTotalChulgoCount,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay0chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay1chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).Fdelay3chulgocnt,0) %></td>
		<td><%= FormatNumber(obalju.FBaljumasterList(i).GetTenMiChulgoCount,0) %></td>
		<td>
			<!--
			<a href="popbaljulist.asp?idx=<%= obalju.FBaljumasterList(i).FBaljuID %>&songjangDiv=<%= obalju.FBaljumasterList(i).FsongjangDiv%>" target=_blank>보기</a>
			-->
		</td>
		<!--
		<td><a href="/admin/pop/viewupchelist.asp?iid=<%= obalju.FBaljumasterList(i).FBaljuID %>" target=_blank>보기</a></td>
		<td><a href="/admin/ordermaster/popsongjangmaker.asp?iid=<%= obalju.FBaljumasterList(i).FBaljuID %>" target=_blank>보기</a></td>
	    -->
	    <td>
			<!--
			<a href="/admin/ordermaster/poporder_gift.asp?baljuid=<%= obalju.FBaljumasterList(i).FBaljuID %>" target=_blank>보기</a>
			-->
		</td>
	    <td>
			<!--
			<a href="/admin/ordermaster/poporder_gift_summary.asp?menupos=1011&research=on&evt_code=&balju_code=<%= obalju.FBaljumasterList(i).FBaljuID %>&viewType=summary&isupchebeasong=N&dateview1=yes&date_display=on" target=_blank>보기</a>
			-->
		</td>
		<td>
			<!--
			<a href="/admin/ordermaster/pop_makeonorder_list.asp?menupos=1568&balju_code=<%= obalju.FBaljumasterList(i).FBaljuID %>" target="_blank">보기</a>
			-->
		</td>
		<td>
			<!--
			<input type="button" class="button" value="출력" onClick="jsPopLogisticsBaljuList(<%= obalju.FBaljumasterList(i).FBaljuID %>)">
			-->
		</td>
	</tr>
	</form>

<% if i+1<obalju.resultBaljucount then %>
	<% if (ppdate<>Left(obalju.FBaljumasterList(i+1).FBaljudate,10)) then %>
<%
sumppcnt = sumppcnt + ppcnt
sumsscnt = sumsscnt + sscnt
sumItemCnt = sumItemCnt + itemCnt

SumChulgoCount      = SumChulgoCount + SubChulgoCount
Sumdelay0chulgocnt  = Sumdelay0chulgocnt + Subdelay0chulgocnt
Sumdelay1chulgocnt  = Sumdelay1chulgocnt + Subdelay1chulgocnt
Sumdelay2chulgocnt  = Sumdelay2chulgocnt + Subdelay2chulgocnt
Sumdelay3chulgocnt  = Sumdelay3chulgocnt + Subdelay3chulgocnt
SumMichulgoCnt      = SumMichulgoCnt + SubMichulgoCnt
SumCancelCnt        = SumCancelCnt + SubCancelCnt
%>
	<tr align="center" bgcolor="#EEEEEE">
		<td></td>
		<td>소계</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber(ppcnt,0) %></td>
		<td>
		<% if ppcnt<>0 then %>
			<%= CLng(sscnt/ppcnt*100) %> %
		<% end if %>
        </td>
        <td><%= FormatNumber(sscnt,0) %></td>
		<td><%= FormatNumber(itemCnt,0) %></td>
        <td><%= FormatNumber(SubCancelCnt,0) %></td>
        <td><font color="<%= ChkIIF(sscnt-SubCancelCnt-SubMichulgoCnt<>SubChulgoCount,"#FF0000","#000000") %>"><%= FormatNumber(SubChulgoCount,0) %></font></td>
    	<td><%= FormatNumber(Subdelay0chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay1chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(Subdelay3chulgocnt,0) %></td>

		<td>
		    <% if SubMichulgoCnt<>0 then %>
		    <b><font color="red"><%= FormatNumber(SubMichulgoCnt,0) %></font></b>
		    <% else %>
		    <%= FormatNumber(SubMichulgoCnt,0) %>
		    <% end if %>
		</td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<% end if %>
<% end if %>



<% next %>

<%
sumppcnt = sumppcnt + ppcnt
sumsscnt = sumsscnt + sscnt
sumItemCnt = sumItemCnt + itemCnt

SumChulgoCount      = SumChulgoCount + SubChulgoCount
Sumdelay0chulgocnt  = Sumdelay0chulgocnt + Subdelay0chulgocnt
Sumdelay1chulgocnt  = Sumdelay1chulgocnt + Subdelay1chulgocnt
Sumdelay2chulgocnt  = Sumdelay2chulgocnt + Subdelay2chulgocnt
Sumdelay3chulgocnt  = Sumdelay3chulgocnt + Subdelay3chulgocnt
SumMichulgoCnt      = SumMichulgoCnt + SubMichulgoCnt
SumCancelCnt        = SumCancelCnt + SubCancelCnt
%>

<% end if %>
	<tr align="center" bgcolor="#EEEEEE">
		<td></td>
		<td>소계</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber(ppcnt,0) %></td>
		<td>
		<% if ppcnt<>0 then %>
			<%= CLng(sscnt/ppcnt*100) %> %
		<% end if %>
    	</td>
    	<td><%= FormatNumber(sscnt,0) %></td>
		<td><%= FormatNumber(itemCnt,0) %></td>

		<td><%= FormatNumber(SubCancelCnt,0) %></td>
		<td><%= FormatNumber(SubChulgoCount,0) %></td>
    	<td><%= FormatNumber(Subdelay0chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay1chulgocnt,0) %></td>
    	<td><%= FormatNumber(Subdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(Subdelay3chulgocnt,0) %></td>
		<td>
		    <% if SubMichulgoCnt<>0 then %>
		    <b><font color="red"><%= FormatNumber(SubMichulgoCnt,0) %></font></b>
		    <% else %>
		    <%= FormatNumber(SubMichulgoCnt,0) %>
		    <% end if %>
		</td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr align="center" bgcolor="#EEEEEE">
		<td></td>
		<td>페이지 합계</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber(sumppcnt,0) %></td>
		<td>
		<% if sumppcnt<>0 then %>
			<%= CLng(sumsscnt/sumppcnt*100) %> %
		<% end if %>
        </td>
        <td><%= FormatNumber(sumsscnt,0) %></td>
		<td><%= FormatNumber(sumItemCnt,0) %></td>

        <td><%= FormatNumber(SumCancelCnt,0) %></td>

    	<td><%= FormatNumber(SumChulgoCount,0) %></td>
    	<td><%= FormatNumber(Sumdelay0chulgocnt,0) %></td>
    	<td><%= FormatNumber(Sumdelay1chulgocnt,0) %></td>
    	<td><%= FormatNumber(Sumdelay2chulgocnt,0) %></td>
		<td><%= FormatNumber(Sumdelay3chulgocnt,0) %></td>
		<td><%= FormatNumber(SumMichulgoCnt,0) %></td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
</table>

<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
