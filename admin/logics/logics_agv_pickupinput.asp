<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : 이상구 생성
'           2020.05.12 정태훈 수정
'           2020.05.20 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%
dim storeid, divcode, scheduledt, vatcode, chargeid, chargename, comment, storemarginrate
dim ArrShopInfo, currencyunit, currencyChar, loginsite, shopdiv, sqlStr, company_no, ischulgonotdisp
dim pickingStationCd, title

	chargeid = session("ssBctid")
	chargename = session("ssBctCname")
	comment = html2db(request("comment"))
	title = request("title")
    pickingStationCd = request("pickingStationCd")

ischulgonotdisp=false

dim itemgubunarr, itemidarr, itemoptionarr
dim itemnamearr, itemoptionnamearr
dim sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr, mwdivarr

dim itemgubun, itemid, itemoption
dim itemname, itemoptionname
dim sellcash, suplycash, buycash, itemno, designer, mwdiv

itemgubunarr = request("itemgubunarr")
itemidarr	= request("itemidarr")
itemoptionarr = request("itemoptionarr")
itemnamearr		= request("itemnamearr")
itemoptionnamearr = request("itemoptionnamearr")
sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
buycasharr = request("buycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")
mwdivarr = request("mwdivarr")

%>
<script>
function Items2Array()
{
	var frm;

	frmMaster.itemgubunarr.value = "";
	frmMaster.itemidarr.value = "";
	frmMaster.itemoptionarr.value = "";
	frmMaster.itemnamearr.value = "";
	frmMaster.itemoptionnamearr.value = "";
	frmMaster.itemnoarr.value = "";
	frmMaster.designerarr.value = "";
	frmMaster.mwdivarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (!IsInteger(frm.itemno.value)){
				alert('갯수는 정수만 가능합니다.');
				frm.itemno.focus();
				return;
			}

			frmMaster.itemgubunarr.value = frmMaster.itemgubunarr.value + frm.itemgubun.value + "|";
			frmMaster.itemidarr.value = frmMaster.itemidarr.value + frm.itemid.value + "|";
			frmMaster.itemoptionarr.value = frmMaster.itemoptionarr.value + frm.itemoption.value + "|";
			frmMaster.itemnamearr.value = frmMaster.itemnamearr.value + frm.itemname.value + "|";
			frmMaster.itemoptionnamearr.value = frmMaster.itemoptionnamearr.value + frm.itemoptionname.value + "|";
			frmMaster.itemnoarr.value = frmMaster.itemnoarr.value + frm.itemno.value + "|";
			frmMaster.designerarr.value = frmMaster.designerarr.value + frm.desingerid.value + "|";
			frmMaster.mwdivarr.value = frmMaster.mwdivarr.value + frm.mwdiv.value + "|";
		}
	}

}

function removeDuplicate() {
	var itemgubunarr, itemidarr, itemoptionarr, itemnamearr, itemoptionnamearr, sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr, mwdivarr;
	var i, j;

	itemgubunarr = frmMaster.itemgubunarr.value.split("|");
	itemidarr = frmMaster.itemidarr.value.split("|");
	itemoptionarr = frmMaster.itemoptionarr.value.split("|");
	itemnamearr = frmMaster.itemnamearr.value.split("|");
	itemoptionnamearr = frmMaster.itemoptionnamearr.value.split("|");
	itemnoarr = frmMaster.itemnoarr.value.split("|");
	designerarr = frmMaster.designerarr.value.split("|");
	mwdivarr = frmMaster.mwdivarr.value.split("|");

	frmMaster.itemgubunarr.value = "";
	frmMaster.itemidarr.value = "";
	frmMaster.itemoptionarr.value = "";
	frmMaster.itemnamearr.value = "";
	frmMaster.itemoptionnamearr.value = "";
	frmMaster.itemnoarr.value = "";
	frmMaster.designerarr.value = "";
	frmMaster.mwdivarr.value = "";

	for (i = 0; i < itemgubunarr.length; i++) {
		if ((itemgubunarr[i] != "XX") && (itemgubunarr[i] != "")) {
			for (j = i + 1; j < itemgubunarr.length; j++) {
				if ((itemgubunarr[i] == itemgubunarr[j]) && (itemidarr[i] == itemidarr[j]) && (itemoptionarr[i] == itemoptionarr[j])) {
					itemgubunarr[j] = "XX";
					itemnoarr[i] = itemnoarr[i]*1 + itemnoarr[j]*1;
				}
			}

			frmMaster.itemgubunarr.value = frmMaster.itemgubunarr.value + itemgubunarr[i] + "|";
			frmMaster.itemidarr.value = frmMaster.itemidarr.value + itemidarr[i] + "|";
			frmMaster.itemoptionarr.value = frmMaster.itemoptionarr.value + itemoptionarr[i] + "|";
			frmMaster.itemnamearr.value = frmMaster.itemnamearr.value + itemnamearr[i] + "|";
			frmMaster.itemoptionnamearr.value = frmMaster.itemoptionnamearr.value + itemoptionnamearr[i] + "|";
			frmMaster.itemnoarr.value = frmMaster.itemnoarr.value + itemnoarr[i] + "|";
			frmMaster.designerarr.value = frmMaster.designerarr.value + designerarr[i] + "|";
			frmMaster.mwdivarr.value = frmMaster.mwdivarr.value + mwdivarr[i] + "|";
		}
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,imwdiv){
	if (iidx!='0'){
		alert('주문서가 일치하지 않습니다. 다시시도해 주세요.');
		return;
	}

    //출고가 기본 0원처리
    var arrsuplycash = isuplycash.split("|");
    isuplycash = "";
    for (var i=0;i<arrsuplycash.length;i++){
        if(i==0){
            isuplycash =  parseInt(arrsuplycash[i])*0;
        }else{
        isuplycash = isuplycash + "|" + parseInt(arrsuplycash[i])*0;
        }
    }

	Items2Array();

	frmMaster.itemgubunarr.value = frmMaster.itemgubunarr.value + igubun;
	frmMaster.itemidarr.value = frmMaster.itemidarr.value + iitemid;
	frmMaster.itemoptionarr.value = frmMaster.itemoptionarr.value + iitemoption;
	frmMaster.itemnoarr.value = frmMaster.itemnoarr.value + iitemno;
	frmMaster.itemnamearr.value = frmMaster.itemnamearr.value + iitemname;
	frmMaster.itemoptionnamearr.value = frmMaster.itemoptionnamearr.value + iitemoptionname;
	frmMaster.designerarr.value = frmMaster.designerarr.value + iitemdesigner;
	frmMaster.mwdivarr.value = frmMaster.mwdivarr.value + imwdiv;

	removeDuplicate();

	frmMaster.submit();
}

function AddItems(frm){
	var popwin;
	var suplyer, shopid;
	var priceGbn;

	popwin = window.open('/admin/newstorage/popjumunitemNew.asp?suplyer=&changesuplyer=Y&shopid=10x10&idx=0&priceGbn=orgprice','chulgoinputadd','width=1280,height=960,scrollbars=yes,resizable=no');
	popwin.focus();
}

function ApplyMargin() {
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			frm.suplycash.value = 1 * frm.sellcash.value * (100 - frmMaster.storemarginrate.value) / 100;
		}
	}
}

function SubmitForm() {
	var frm = document.frmMaster;

    if (frm.pickingStationCd.value == "") {
        alert("피킹스테이션을 선택하세요.");
        return;
    }

    if (frm.title.value == "") {
        alert("제목을 선택하세요.");
        return;
    }

    if (confirm("저장하시겠습니까?") != true) {
        return;
	}

    Items2Array();

    frm.mode.value = "write";
    frm.action = "logics_agv_pickup_process.asp";
    frm.submit();

}

function tempSave(){
	var frm = document.frmMaster;

	if (frm.storeid.value == "") {
        alert("출고처를 선택하세요.");
        return;
    }

	if ( (frm.storeid.value == "promotion") ) {		//  || (frm.storeid.value == "etcsales")
		alert("출고처 promotion 는 선택할 수 없습니다.");
		//alert("출고처 promotion, etcsales 는 선택할 수 없습니다.");
        return;
	}

    Items2Array();

	frm.mode.value = "temp";
    frm.action = "chulgoedit_process.asp";
    frm.submit();
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
        	<font color="red"><strong>피킹지시입력</strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->

	<form name="frmMaster" method="post" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="chargeid" value="<%= chargeid %>">
	<input type="hidden" name="chargename" value="<%= chargename %>">
	<input type="hidden" name="vatcode" value="<%= vatcode %>">

	<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
	<input type="hidden" name="itemidarr" value="<%= itemidarr %>">
	<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
	<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
	<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
	<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
	<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
	<input type="hidden" name="buycasharr" value="<%= buycasharr %>">
	<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
	<input type="hidden" name="designerarr" value="<%= designerarr %>">
	<input type="hidden" name="mwdivarr" value="<%= mwdivarr %>">
    <tr align="center" bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">IDX</td>
		<td width=400 align="left"></td>
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">스테이션</td>
		<td align="left">
            <% Call drawSelectStationByStationGubun("PICK", "pickingStationCd", pickingStationCd) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">등록일</td>
		<td align="left"><%= Now() %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">등록자</td>
		<td align="left"><%= chargename %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">전송상태</td>
		<td align="left"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"></td>
		<td align="left"></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">제목</td>
		<td align="left" colspan="3">
            <input type="text" class="text" size="80" name="title" value="<%= title %>">
        </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
		<td colspan="3" align="left"><textarea class="textarea" name="comment" cols=80 rows=6><%= comment %></textarea></td>
	</tr>
</table>
<%

itemgubunarr = split(itemgubunarr,"|")
itemidarr	= split(itemidarr,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
mwdivarr = split(mwdivarr,"|")

dim cnt, i

cnt = ubound(itemidarr)
if cnt < 0 then cnt = 0
dim selltotal, suplytotal, buytotal
selltotal = 0
suplytotal = 0
buytotal = 0

%>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="9">
			<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
			        	<font color="red"><strong>상품목록</strong></font>
	        		</td>
	        		<td align="right">
	        			총건수:  <%= cnt %>
			        	&nbsp;
			        	<input type="button" class="button" value=" 상품추가 " onClick="AddItems(frmMaster)">
	        		</td>
	        	</tr>
	        </table>
		</td>
	</tr>
	</form>
	<!-- 상단바 끝 -->

    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">바코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="60">수량</td>
	</tr>
	<% for i=0 to cnt-1 %>
	<form name="frmBuyPrc_<%= i %>" method="post" action="">
	<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
	<input type="hidden" name="itemid" value="<%= itemidarr(i) %>">
	<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
	<input type="hidden" name="itemname" value="<%= itemnamearr(i) %>">
	<input type="hidden" name="itemoptionname" value="<%= itemoptionnamearr(i) %>">
	<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">
	<input type="hidden" name="mwdiv" value="<%= mwdivarr(i) %>">

	<tr bgcolor="#FFFFFF">
		<td align=center >
		<% if mwdivarr(i)="M" then %>
		<font color="#EE4444"><%= itemgubunarr(i) %>-<%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %>-<%= itemoptionarr(i) %></font>
		<% elseif mwdivarr(i)="U" then %>
		<font color="#4444EE"><%= itemgubunarr(i) %>-<%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %>-<%= itemoptionarr(i) %></font>
		<% else %>
		<%= itemgubunarr(i) %>-<%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %>-<%= itemoptionarr(i) %>
		<% end if %>
		</td>
		<td ><%= itemnamearr(i) %></td>
		<td ><%= itemoptionnamearr(i) %></td>

		<td align=right><input type="text" class="text" name="itemno" value="<%= itemnoarr(i) %>"  size="4" maxlength="4"></td>
	</form>
	<% next %>

	<% if (cnt>0) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td colspan="2" align="center">
		<td></td>
	</tr>
	<% end if %>

</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
	<tr height="25"  >
		<td colspan="15" align="center">
			<% if (cnt>0) then %>
			<input type="button" class="button" value=" 저 장 " onclick="SubmitForm()">
        	<% else %>
        	저장할 내역이 없습니다.
        	<% end if %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
