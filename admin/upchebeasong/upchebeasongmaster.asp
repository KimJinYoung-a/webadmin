<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate, searchtype
dim dateback, research
dim cdl, cdm, cds, dispCate

nowdate = Left(CStr(now()),10)

yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
dd1     = requestCheckVar(request("dd1"),2)
yyyy2   = requestCheckVar(request("yyyy2"),4)
mm2     = requestCheckVar(request("mm2"),2)
dd2     = requestCheckVar(request("dd2"),2)
searchtype  = requestCheckVar(request("searchtype"),32)
research    = requestCheckVar(request("research"),32)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

if (research="") and (searchtype="") then searchtype="ndayexists"


if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy1,mm2, dd2-30)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim ojumun

set ojumun = new CBaljuMaster
ojumun.FRectRegStart    = LEft(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
ojumun.FRectRegEnd      = searchnextdate
ojumun.FRectDispCDL     = cdl
ojumun.FRectDispCDM     = cdm
ojumun.FRectDispCDS     = cds
ojumun.FRectDispCate	= dispCate
ojumun.DesignerDateMiBaljuMiBeasongList

dim i, mitongbocnttotal, mibaljutotal, misendtotal
dim ndaymibaljutotal, ndaymisendtotal
dim p_ndaymibaljutotal, p_ndaymisendtotal
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function PopSMS(v){
	//var popwin =
	//popwin.focus();
}

function changecontent(){
    //
}

function PopUpcheBrandInfo(v){
	var popwin = window.open("/admin/lib/popupchebrandinfo.asp?designer=" + v,"popupchebrandinfo","width=640 height=660 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpchebeasongList(imakerid, idetailstate){
    var popwin = window.open("upchebeasonglist.asp?makerid=" + imakerid + "&cdl=<%= cdl %>&cdm=<%= cdm %>&cds=<%= cds %>&disp=<%= dispCate %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>&detailstate=" + idetailstate,"PopUpchebeasongList","width=1000 height=800 scrollbars=yes resizable=yes");
    popwin.focus();
}

function popUpcheMisend(makerID, dplusOver, currState)
{
	var f = document.frm;

	var url = "/admin/upchebeasong/upchemibeasonglist.asp?menupos=246&MisendState=";
	url += "&makerID=" + makerID;
	url += "&dplusOver=" + dplusOver;
	url += "&currState=" + currState;
	url += "&cdl=<%=cdl%>";// + f.cdl.value;
	url += "&cdm=<%=cdm%>";
	url += "&cds=<%=cds%>";
	url += "&disp=<%=dispCate%>";
	url += "&yyyy1=<%=yyyy1%>";// + f.yyyy1.value;
	url += "&yyyy2=<%=yyyy2%>";// + f.yyyy2.value;
	url += "&mm1=<%=mm1%>";// + f.mm1.value;
	url += "&mm2=<%=mm2%>";// + f.mm2.value;
	url += "&dd1=<%=dd1%>";// + f.dd1.value;
	url += "&dd2=<%=dd2%>";// + f.dd2.value;

    var window_width = 1024;
    var window_height = 800;
	var popwin = window.open( url , "cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

function popUpcheMisendTop(makerID, dplusOver, currState)
{
	var f = document.frm;

	var url = "/admin/upchebeasong/upchemibeasonglist.asp?menupos=246&MisendState=";
	url += "&makerID=" + makerID;
	url += "&Dtype=topN";
	url += "&dplusOver=" + dplusOver;
	url += "&currState=" + currState;
	url += "&cdl=<%=cdl%>";// + f.cdl.value;
	url += "&cdm=<%=cdm%>";
	url += "&cds=<%=cds%>";
	url += "&disp=<%=dispCate%>";
	url += "&yyyy1=<%=yyyy1%>";// + f.yyyy1.value;
	url += "&yyyy2=<%=yyyy2%>";// + f.yyyy2.value;
	url += "&mm1=<%=mm1%>";// + f.mm1.value;
	url += "&mm2=<%=mm2%>";// + f.mm2.value;
	url += "&dd1=<%=dd1%>";// + f.dd1.value;
	url += "&dd2=<%=dd2%>";// + f.dd2.value;

    var window_width = 1024;
    var window_height = 800;
	var popwin = window.open( url , "cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			검색기간(주문일기준) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			구분 :
				<input type=radio name=searchtype value="ndayexists" <% if searchtype="ndayexists" then response.write "checked" %> >2일 미발주 또는 4일 미출고 존재
				<input type=radio name=searchtype value="" <% if searchtype="" then response.write "checked" %> >전체
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<!--
			브랜드 카테고리: <% DrawSelectBoxCategoryLarge "cdl",cdl %>
			-->
			전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2">브랜드ID</td>
		<td rowspan="2">브랜드명</td>
		<td rowspan="2">업체명</td>
		<td rowspan="2">카테고리</td>
		<td width="60" rowspan="2">통보이전<br>(결제완료)</td>
		<td width="160" colspan="3">미확인내역</td>
		<td width="160" colspan="3">미출고내역</td>
		<td width="40" rowspan="2">전체<br>내역</td>
		<td width=100 rowspan="2">배송담당연락처</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<!--td width="40">D+0</td -->
		<td width="40">D+1<br>이하</td>
		<td width="40"><b>D+2<br>이상</b></td>
		<td width="40">총건수</td>
		<td width="40">D+3<br>이하</td>
		<td width="40"><b>D+4<br>이상</b></td>
		<!--td width="40"><b>D+9<br>이상</b></td-->
		<td width="40">총건수</td>
	</tr>

	<% for i=0 to ojumun.FResultCount-1 %>
	<% if (searchtype<>"ndayexists") or ((searchtype="ndayexists") and ((ojumun.FMasterItemList(i).FNDayMiBaljuCnt>0) or (ojumun.FMasterItemList(i).FNDayMiBeasongCnt>0))) then %>
    <%
    mitongbocnttotal    = mitongbocnttotal + ojumun.FMasterItemList(i).Fmitongbocnt
    mibaljutotal        = mibaljutotal + ojumun.FMasterItemList(i).FMiBalJuCount
    misendtotal         = misendtotal + ojumun.FMasterItemList(i).FMiBeasongCount

    ndaymibaljutotal    = ndaymibaljutotal + ojumun.FMasterItemList(i).FNDayMiBaljuCnt
    ndaymisendtotal     = ndaymisendtotal + ojumun.FMasterItemList(i).FNDayMiBeasongCnt

    p_ndaymibaljutotal    = p_ndaymibaljutotal + ojumun.FMasterItemList(i).FP_NDayMiBaljuCnt
    p_ndaymisendtotal     = p_ndaymisendtotal + ojumun.FMasterItemList(i).FP_NDayMiBeasongCnt
    %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><a href="javascript:PopUpcheBrandInfoEdit('<%= ojumun.FMasterItemList(i).FMakerid %>');"><%= ojumun.FMasterItemList(i).FMakerid %></a></td>
		<td><%= ojumun.FMasterItemList(i).FSocNameKor %></td>
		<td><%= ojumun.FMasterItemList(i).FCompanyName %></td>
		<td><%= ojumun.FMasterItemList(i).Fcatename %></td>
	    <!--td></td-->
	    <td><a href="javascript:popUpcheMisend('<%= ojumun.FMasterItemList(i).FMakerid %>','','0');"><%= ojumun.FMasterItemList(i).Fmitongbocnt %></a></td>
	    <td><%= ojumun.FMasterItemList(i).FP_NDayMiBaljuCnt %></td>
	    <td>
			<a href="javascript:popUpcheMisend('<%= ojumun.FMasterItemList(i).FMakerid %>','2','2');">
	        <% if ojumun.FMasterItemList(i).FNDayMiBaljuCnt>0 then %>
	        <strong><%= ojumun.FMasterItemList(i).FNDayMiBaljuCnt %></strong>
	        <% else %>
	        <%= ojumun.FMasterItemList(i).FNDayMiBaljuCnt %>
	        <% end if %>
			</a>
	    </td>
	    <td><a href="javascript:popUpcheMisend('<%= ojumun.FMasterItemList(i).FMakerid %>','','2');"><%= ojumun.FMasterItemList(i).FMiBalJuCount %></a></td>

	    <td><%= ojumun.FMasterItemList(i).FP_NDayMiBeasongCnt %></td>
	    <td>
			<a href="javascript:popUpcheMisend('<%= ojumun.FMasterItemList(i).FMakerid %>','4','3');">
	        <% if ojumun.FMasterItemList(i).FNDayMiBeasongCnt>0 then %>
	        <strong><%= ojumun.FMasterItemList(i).FNDayMiBeasongCnt %></strong>
	        <% else %>
	        <%= ojumun.FMasterItemList(i).FNDayMiBeasongCnt %>
	        <% end if %>
			</a>
	    </td>
	    <!--td></td-->
	    <td><a href="javascript:popUpcheMisend('<%= ojumun.FMasterItemList(i).FMakerid %>','','3');"><%= ojumun.FMasterItemList(i).FMiBeasongCount %></a></td>
	    <td><a href="javascript:popUpcheMisend('<%= ojumun.FMasterItemList(i).FMakerid %>','','');"><img src="/images/icon_arrow_link.gif" border="0"></a></td>
	    <td><%= ojumun.FMasterItemList(i).FDeliverHp %></td>
	</tr>
	<% end if %>
	<% next %>
	<tr align=center bgcolor="#FFFFFF">
		<td>총계</td>
		<td></td>
		<td></td>
		<td></td>
		<td><a href="javascript:<%= CHKIIF(mitongbocnttotal<=300,"popUpcheMisendTop","popUpcheMisend") %>('','','0');"><%= FormatNumber(mitongbocnttotal,0) %></a></td>
		<td><%= FormatNumber(p_ndaymibaljutotal,0) %></td>
		<td><a href="javascript:<%= CHKIIF(ndaymibaljutotal<=300,"popUpcheMisendTop","popUpcheMisend") %>('','2','2');"><%= FormatNumber(ndaymibaljutotal,0) %></a></td>
		<td><a href="javascript:<%= CHKIIF(mibaljutotal<=300,"popUpcheMisendTop","popUpcheMisend") %>('','','2');"><%= FormatNumber(mibaljutotal,0) %></a></td>
		<td><%= FormatNumber(p_ndaymisendtotal,0) %></td>
		<td><a href="javascript:<%= CHKIIF(ndaymisendtotal<=300,"popUpcheMisendTop","popUpcheMisend") %>('','4','3');"><%= FormatNumber(ndaymisendtotal,0) %></a></td>
		<td><a href="javascript:<%= CHKIIF(misendtotal<=300,"popUpcheMisendTop","popUpcheMisend") %>('','','3');"><%= FormatNumber(misendtotal,0) %></a></td>
		<td></td>
		<td></td>
	</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
