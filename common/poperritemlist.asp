<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%

dim itemgubun, itemid, itemoption, designer

itemgubun   = requestCheckVar(request("itemgubun"),2)
itemid      = requestCheckVar(request("itemid"),10)
itemoption  = requestCheckVar(request("itemoption"),4)
designer    = requestCheckVar(request("designer"),32)


if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

fromdate = request("fromdate")
todate = request("todate")

if fromdate<>"" then
	yyyy1 = Left(fromdate,4)
	mm1 = Mid(fromdate,6,2)
	dd1 = Mid(fromdate,9,2)
else
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
end if

if todate<>"" then
	yyyy2 = Left(todate,4)
	mm2 = Mid(todate,6,2)
	dd2 = Mid(todate,9,2)
else
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
end if



if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateSerial(yyyy2, mm2, dd2+1))


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = fromdate
osummarystock.FRectEndDate	 = todate
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
osummarystock.FRectMakerid =  designer

if ((itemid<>"") or (designer<>"")) then
	osummarystock.GetDailyErrItemList
end if
dim i, totitemno, totrealerrno

totitemno=0
totrealerrno=0


%>
<script language='javascript'>
function editThis(yyyymmdd,itemgubun,itemid,itemoption,ii){
    var frm = document.frmAct;
    var frmArr = document.frmArr;
    var errbaditemno;
    var errrealcheckno;

    if (!frmArr.errbaditemno){
        return;
    }

    if (frmArr.errbaditemno.length){
        errbaditemno = frmArr.errbaditemno[ii].value
        errrealcheckno = frmArr.errrealcheckno[ii].value
    }else{
        errbaditemno = frmArr.errbaditemno.value
        errrealcheckno = frmArr.errrealcheckno.value
    }

    if (confirm('관리자 메뉴 - 수정하시겠습니까?')){
        frm.mode.value="editdailyerrlog";
        frm.yyyymmdd.value=yyyymmdd;
        frm.itemgubun.value=itemgubun;
        frm.itemid.value=itemid;
        frm.itemoption.value=itemoption;
        frm.errbaditemno.value=errbaditemno;
        frm.errrealcheckno.value=errrealcheckno;

        frm.submit();
    }
}

function inputDummiErr(comp,itemgubun,itemid,itemoption){
    var bufFrm = comp.form;
    var frm = document.frmAct;
    if (bufFrm.bufYYYYMMDD.value.length!=10){
        alert('에러 날짜를 입력하세요.');
        return;
    }

    if (confirm('관리자 메뉴 - 입력 하시겠습니까?')){
        frm.mode.value="dummidailyerrlog";
        frm.yyyymmdd.value=bufFrm.bufYYYYMMDD.value;
        frm.itemgubun.value=itemgubun;
        frm.itemid.value=itemid;
        frm.itemoption.value=itemoption;
        frm.errbaditemno.value=0;
        frm.errrealcheckno.value=0;

        frm.submit();
    }
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <form name="frm" method="get" onsubmit="return false;">

		<tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td valign="top" bgcolor="F4F4F4">
        	상품코드 :
        	 <input type="text" name="itemgubun" value="<%= itemgubun %>" Maxlength="2" size="2">
        	 <input type="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">
        	 <input type="text" name="itemoption" value="<%= itemoption %>" Maxlength="4" size="4">
        	검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        </td>
        <td valign="top" align="right" bgcolor="F4F4F4">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" onclick="document.frm.submit();"></a>
        </td>
	</tr>

	<tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td valign="top" bgcolor="F4F4F4">
        	브랜드명 : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
        </td>
        <td valign="top" align="right" bgcolor="F4F4F4">
        <% if (C_ADMIN_AUTH) then %>
        빈에러로그입력
        <input type="text" name="bufYYYYMMDD" value="" size="10" maxlength="10">
        <input type="button" value="입력" onClick="inputDummiErr(this,'<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
        <% end if %>
        </td>
	</tr>
	</form>

</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmArr">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">등록일</td>
      <td width="90">브랜드ID</td>
      <td width="30">배송<br>구분</td>
      <td width="30">매입<br>구분</td>
      <td width="25">구분</td>
      <td width="50">상품코드</td>
      <td>아이템명</td>
      <td width="150">옵션</td>
      <td width="50">소비자가</td>
      <td width="80">불량</td>
      <td width="80">(실사)<br>오차</td>
      <% if (C_ADMIN_AUTH) then %>
      <td width="40">수정</td>
      <% end if %>
    </tr>
    <% for i=0 to osummarystock.FResultCount -1 %>
        <% totitemno = totitemno + osummarystock.FItemList(i).Ferrbaditemno %>
        <% totrealerrno = totrealerrno + osummarystock.FItemList(i).Ferrrealcheckno %>
    <tr align="center" bgcolor="#FFFFFF">
      <td><%= osummarystock.FItemList(i).FYYYYMMDD %></td>
      <td><%= osummarystock.FItemList(i).FMakerid %></td>
      <td><%= osummarystock.FItemList(i).GetdeliverytypeName %></td>
      <td><%= osummarystock.FItemList(i).GetMwDivName %></td>
      <td><%= osummarystock.FItemList(i).Fitemgubun %></td>
      <td><%= osummarystock.FItemList(i).Fitemid %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemName %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= FormatNumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
      <td><%= osummarystock.FItemList(i).Ferrbaditemno %>
      <% if (C_ADMIN_AUTH) then %>
      <input type="text" name="errbaditemno" value="<%= osummarystock.FItemList(i).Ferrbaditemno %>" size="3">
      <% end if %>
      </td>
      <td><%= osummarystock.FItemList(i).Ferrrealcheckno %>
      <% if (C_ADMIN_AUTH) then %>
      <input type="text" name="errrealcheckno" value="<%= osummarystock.FItemList(i).Ferrrealcheckno %>" size="3">
      <% end if %>
      </td>
      <% if (C_ADMIN_AUTH) then %>
      <td><input type="button" value="수정" onClick="editThis('<%= osummarystock.FItemList(i).FYYYYMMDD %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).Fitemoption %>',<%= i %>)"></td>
      <% end if %>
    </tr>
	<% next %>
    <tr align="center" bgcolor="#EEEEEE">
      <td>ToTal</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td><%= totitemno %></td>
      <td><%= totrealerrno %></td>
      <% if (C_ADMIN_AUTH) then %>
      <td></td>
      <% end if %>
    </tr>
</form>
</table>

<form name="frmAct" method="post" action="/admin/stock/stockrefresh_process.asp">
<input type="hidden" name="mode" value="editdailyerrlog">
<input type="hidden" name="yyyymmdd" value="">
<input type="hidden" name="itemgubun" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="errbaditemno" value="">
<input type="hidden" name="errrealcheckno" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->