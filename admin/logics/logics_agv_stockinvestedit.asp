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
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%
dim storeid, divcode, scheduledt, vatcode, chargeid, chargename, comment, storemarginrate
dim ArrShopInfo, currencyunit, currencyChar, loginsite, shopdiv, sqlStr, company_no, ischulgonotdisp
dim pickingStationCd, title
dim masteridx

chargeid = session("ssBctid")
chargename = session("ssBctCname")

masteridx = requestCheckVar(request("idx"), 32)

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

dim oPickupMaster
set oPickupMaster = new CAGVItems
	oPickupMaster.FRectMasterIdx = masteridx
	oPickupMaster.GetStockInvestMasterOne

dim oPickupDetail
set oPickupDetail = new CAGVItems
	oPickupDetail.FRectMasterIdx = masteridx
	oPickupDetail.FPageSize = 20000
	oPickupDetail.GetStockInvestDetailList


dim IsEditAvailable : IsEditAvailable = True
if Not IsNull(oPickupMaster.FOneItem.Fstatus) then
    if (oPickupMaster.FOneItem.Fstatus >= 50) then
        '// 전송완료 이후 수정불가
        IsEditAvailable = False
    end if
end if


dim i, j, k

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
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
<% if Not IsEditAvailable then %>
	alert('전송완료 이후에는 수정할 수 없습니다.');
	return;
<% end if %>

    var frmDetail = document.frmDetail;
    var frm;

	frmDetail.itemgubunarr.value = igubun;
	frmDetail.itemidarr.value = iitemid;
	frmDetail.itemoptionarr.value = iitemoption;
	frmDetail.itemnoarr.value = iitemno;

	frmDetail.mode.value = "adddetail";
	frmDetail.action = "logics_agv_stockinvest_process.asp";
	frmDetail.submit();
}

function AddItems(frm){
	var popwin;
	var suplyer, shopid;
	var priceGbn;

	popwin = window.open('/admin/newstorage/popjumunitemNew.asp?suplyer=&changesuplyer=Y&shopid=10x10&idx=0&priceGbn=orgprice&skipChkItemNo=Y','chulgoinputadd','width=1280,height=960,scrollbars=yes,resizable=no');
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
    frm.action = "logics_agv_stockinvest_process.asp";
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

function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmDetail;

	if (frm.chk.length) {
		for (var i = 0; i < frm.chk.length; i++) {
			frm.chk[i].checked = bool;
			AnCheckClick(frm.chk[i]);
		}
	} else {
		frm.chk.checked = bool;
		AnCheckClick(frm.chk);
	}
}

function DelDetail(masterfrm,iid){
<% if Not IsEditAvailable then %>
	alert('전송완료 이후에는 수정할 수 없습니다.');
	return;
<% end if %>

    var frmDetail = document.frmDetail;
	var frm;
	var found = false;
	for (var i = 0; i < frmDetail.elements.length; i++) {
		frm = frmDetail.elements[i];
		if (frm.name == "chk") {
			if (frm.checked == true) {
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("선택된 상품을 삭제합니다.") == true) {
			frmDetail.mode.value = "deldetail";
			frmDetail.action = "logics_agv_stockinvest_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("선택된 상품이 없습니다.");
	}
}

function jsSaveForm() {
<% if Not IsEditAvailable then %>
	alert('전송완료 이후에는 수정할 수 없습니다.');
	return;
<% end if %>
    var frm, frmMaster, frmDetail, i;
    var didxarr, itemnoarr;

    frmMaster = document.frmMaster;
    frmDetail = document.frmDetail;

    frm = frmMaster;
    if (frm.pickingStationCd.value == "") {
        alert("피킹스테이션을 선택하세요.");
        return;
    }

    if (frm.title.value == "") {
        alert("제목을 선택하세요.");
        return;
    }

    didxarr = '';
    itemnoarr = '';

	if (frmDetail.chk.length) {
		for (var i = 0; i < frmDetail.chk.length; i++) {
			didxarr = didxarr + ',' + frmDetail.chk[i].value;
		}
	} else {
		didxarr = didxarr + ',' + frmDetail.chk.value;
	}

    frmMaster.didxarr.value = didxarr;

    if (confirm("저장하시겠습니까?") != true) {
        return;
	}

    frmMaster.action = "logics_agv_stockinvest_process.asp";
	frmMaster.mode.value = "modi";
	frmMaster.submit();
}

function jsGotoList() {
    location.replace('logics_agv_stockInvestList.asp?menupos=<%= menupos %>');
}

function DelMaster() {
<% if Not IsEditAvailable then %>
	alert('전송완료 이후에는 수정할 수 없습니다.');
	return;
<% end if %>
    var frmMaster = document.frmMaster;

    if (confirm("삭제하시겠습니까?") != true) {
        return;
	}

    frmMaster.action = "logics_agv_stockinvest_process.asp";
	frmMaster.mode.value = "delmaster";
	frmMaster.submit();
}

function jsCallAjax(url) {
	$.ajax({
		url: url,
		type: 'get',
		crossDomain: true,
		data: {},
		dataType: 'json',
		success: function(data) {
			if (data.resultCode == '00') {
				$.each(data.resultData.skuList, function(idx, val) {
					$('#agvstock_' + val.skuCd).text(val.totalQty*1 + val.adjustQty*1);
				});
			} else {
				alert(data.resultMessage);
			}
		},
		error: function(jqXHR, textStatus, ex) {
			alert(textStatus + "," + ex + "," + jqXHR.responseText);
		}
	});
}

function jsUpdateAgvStockInfo() {
    var url, url2;
    var skuCdArray = '';
    var frmDetail = document.frmDetail;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'https://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=currstockListView&skuCdArray=';

	if (frmDetail.skuCd && frmDetail.skuCd.length) {
        url2 = url;

		for (var i = 0; i < frmDetail.skuCd.length; i++) {
			skuCdArray = skuCdArray + ',' + frmDetail.skuCd[i].value;
            if ((i > 0) && (((i % 100) == 0) || (frmDetail.skuCd.length == (i+1)))) {
                url = url2 + skuCdArray;

                jsCallAjax(url);

                skuCdArray = '';
            }

		}

        return;
    } else if (frmDetail.skuCd) {
        skuCdArray = skuCdArray + ',' + frmDetail.skuCd.value;
	} else {
		return;
	}

    url = url + skuCdArray;

    jsCallAjax(url);
}

// AGV 재고조회
function jsGetStockState() {
    var url;
    var skuCdArray = '';
    var frmDetail = document.frmDetail;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'https://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=currstockList&skuCdArray=';

	if (frmDetail.skuCd && frmDetail.skuCd.length) {
		for (var i = 0; i < frmDetail.skuCd.length; i++) {
			skuCdArray = skuCdArray + ',' + frmDetail.skuCd[i].value;
		}
    } else if (frmDetail.skuCd) {
        skuCdArray = skuCdArray + ',' + frmDetail.skuCd.value;
	} else {
		return;
	}

    //alert(skuCdArray);
    return;

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                // alert('업데이트되었습니다.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
}

// 재고조사 전송
function jsSendStockInvest() {
    var url;
    var skuCdArray = '';
    var frmDetail = document.frmDetail;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'https://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=agvstockinvest&masteridx=<%= masteridx %>';

    if (confirm('전송하시겠습니까?') != true) { return; }

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                alert('전송되었습니다.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
}

// 재고조사 전송취소
function jsSendStockInvestCancel() {
    var url;
    var skuCdArray = '';
    var frmDetail = document.frmDetail;

    <% IF application("Svr_Info")="Dev" THEN %>
    url = 'http://testwapi.10x10.co.kr';
    <% ELSE %>
    url = 'https://wapi.10x10.co.kr';
    <% END IF %>

    url = url + '/agv/api.asp?mode=agvstockinvestdel&masteridx=<%= masteridx %>';

    if (confirm('취소하시겠습니까?') != true) { return; }

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                alert('취소되었습니다.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
}

function PopModiRackCode(mode) {
	var frm;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";

	var makerid = "";

	var selecteditemcount = 0;

	if (CheckSelected() != true) {
		alert("선택아이템이 없습니다.[0]");
		return;
	}

    frm = document.frmDetail;
    if (frm.chk && frm.chk.length) {
        for (var i = 0; i < frm.chk.length; i++) {
            if (frm.chk[i].checked == true) {
                itemgubunarr = itemgubunarr + frm.itemgubun[i].value + "|";
				itemidarr = itemidarr + frm.itemid[i].value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption[i].value + "|";
            }
        }
    } else if (frm.chk) {
        if (frm.chk.checked == true) {
            itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
			itemidarr = itemidarr + frm.itemid.value + "|";
			itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
        }
    } else {
        alert("선택된 상품이 없습니다.[1]");
		return;
    }

	if (itemgubunarr == "") {
		alert("선택된 상품이 없습니다.[1]");
		return;
	}

    var popwin;
	var url = "/admin/stock/popMultiRackCode.asp";

	document.frmActPop.mode.value=mode;
	document.frmActPop.itemgubunarr.value=itemgubunarr;
	document.frmActPop.itemidadd.value=itemidarr;
	document.frmActPop.itemoptionarr.value=itemoptionarr;

    popwin = window.open("", "PopModiRackCode","width=300,height=150,scrollbars=yes,resizable=yes");
    popwin.focus();
    document.frmActPop.action=url;
    document.frmActPop.target="PopModiRackCode";
    document.frmActPop.submit();
}

function CheckSelected(){
    var frmDetail = document.frmDetail;
	var frm;
	var found = false;

	for (var i = 0; i < frmDetail.elements.length; i++) {
		frm = frmDetail.elements[i];
		if (frm.name == "chk") {
			if (frm.checked == true) {
				found = true;
            } else {
			}
		}
	}

	if (!found) {
		return false;
	}
	return true;
}

function PopAgvStockOut(masteridx) {
    var popwin = window.open('pop_interface_agv_stockout.asp?idx=' + masteridx,'PopAgvStockOut' + masteridx,'width=800, height=600, resizabled=yes, scrollbars=yes');
	popwin.focus();
}

$(document).ready(function(){
    jsUpdateAgvStockInfo();
});

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
    <input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="chargeid" value="<%= chargeid %>">

    <input type="hidden" name="didxarr" value="">
    <input type="hidden" name="itemnoarr" value="">

	<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
	<input type="hidden" name="itemidarr" value="<%= itemidarr %>">
	<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
	<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
	<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
	<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
	<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
	<input type="hidden" name="buycasharr" value="<%= buycasharr %>">

	<input type="hidden" name="designerarr" value="<%= designerarr %>">
	<input type="hidden" name="mwdivarr" value="<%= mwdivarr %>">
    <tr align="center" bgcolor="#FFFFFF" height="30">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">IDX</td>
		<td width=400 align="left"><%= oPickupMaster.FOneItem.Fidx %></td>
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">스테이션</td>
		<td align="left">
            <% Call drawSelectStationByStationGubun("PICK", "pickingStationCd", oPickupMaster.FOneItem.FstationCd) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">등록일</td>
		<td align="left"><%= oPickupMaster.FOneItem.Fregdate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">등록자</td>
		<td align="left"><%= oPickupMaster.FOneItem.Freguserid %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">전송상태</td>
		<td align="left">
            <%= oPickupMaster.FOneItem.getStatusName %>
            <% if IsNull(oPickupMaster.FOneItem.Fstatus) or (oPickupMaster.FOneItem.Fstatus < 50) then %>
            <input type="button" class="button" value=" 전송하기 " onclick="jsSendStockInvest()">
            <% end if %>
            <input type="button" class="button" value=" 전송취소 " onclick="jsSendStockInvestCancel()">
        </td>
		<td bgcolor="<%= adminColor("tabletop") %>">재고조사지시 번호</td>
		<td align="left"><%= oPickupMaster.FOneItem.FinventorySurveyOrderId %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">제목</td>
		<td align="left" colspan="3">
            <input type="text" class="text" size="80" name="title" value="<%= oPickupMaster.FOneItem.Ftitle %>">
        </td>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">작업지시코드</td>
		<td align="left" colspan="3">
            <%= oPickupMaster.FOneItem.FrequestNo %>
        </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
		<td colspan="3" align="left"><textarea class="textarea" name="comment" cols=80 rows=6><%= oPickupMaster.FOneItem.Fcomment %></textarea></td>
	</tr>
</table>

<p />

<!--
<div style="width: 100%; height: 50px; display: flex; justify-content: center; align-items: center;">
    <input type="button" class="button" name="stock_index_print" value="선택상품 상품 랙코드수정" onclick="PopModiRackCode('modiitem');">
	&nbsp;&nbsp;
	<input type="button" class="button" name="stock_index_print" value="선택상품 [옵션별] 랙코드수정" onclick="PopModiRackCode('modiopt');">
    &nbsp;&nbsp;
	<input type="button" class="button" name="stock_index_print" value="AGV결품내역" onclick="PopAgvStockOut(<%= masteridx %>);">
</div>
-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="10">
			<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
			        	<font color="red"><strong>상품목록</strong></font>
	        		</td>
	        		<td align="right">
	        			총건수:  <%= oPickupDetail.FResultCount %>
			        	&nbsp;
			        	<input type="button" class="button" value=" 상품추가 " onClick="AddItems(frmMaster)" <%= CHKIIF(IsEditAvailable, "", "disabled") %>>
	        		</td>
	        	</tr>
	        </table>
		</td>
	</tr>
	</form>
	<!-- 상단바 끝 -->
	<form name="frmDetail" method="post" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnoarr" value="">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width=20><Input Type="checkbox" name="ckall" onClick="ckAll(this)"></td>
        <td width="150">상품코드</td>
		<td width="200">브랜드</td>
        <td>상품명</td>
		<td>옵션명</td>
        <td width="60">AGV<br />(현)재고</td>
        <td>비고</td>
	</tr>
    <% for i = 0 to oPickupDetail.FResultCount -1 %>
	<tr bgcolor="#FFFFFF">
        <input type="hidden" name="itemgubun" value="<%= oPickupDetail.FItemList(i).Fitemgubun %>">
	    <input type="hidden" name="itemid" value="<%= oPickupDetail.FItemList(i).FItemId %>">
	    <input type="hidden" name="itemoption" value="<%= oPickupDetail.FItemList(i).FItemOption %>">
        <td align="center"><input type=checkbox name=chk value="<%= oPickupDetail.FItemList(i).Fidx %>" onClick="AnCheckClick(this);"></td>
        <td align="center"><%= oPickupDetail.FItemList(i).Fitemgubun %>-<%= CHKIIF(oPickupDetail.FItemList(i).FItemId>=1000000,Format00(8,oPickupDetail.FItemList(i).FItemId),Format00(6,oPickupDetail.FItemList(i).FItemId)) %>-<%= oPickupDetail.FItemList(i).FItemOption %></td>
        <td><%= oPickupDetail.FItemList(i).Fmakerid %></td>
		<td><%= oPickupDetail.FItemList(i).Fitemname %></td>
        <td><%= oPickupDetail.FItemList(i).Fitemoptionname %></td>
        <td align="center"><div id="agvstock_<%= oPickupDetail.FItemList(i).FskuCd %>">-</div></td>
        <td>
            <input type="hidden" name="skuCd" value="<%= oPickupDetail.FItemList(i).FskuCd %>">
		</td>
    </tr>
	<%
	if i mod 3000 = 0 then
		Response.Flush		' 버퍼리플래쉬
	end if

	next
	%>
    </form>
</table>

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
	<tr height="25"  >
		<td colspan="15" align="center">
            <% if IsEditAvailable then %>
            <input type="button" class="button" value=" 저 장 하 기 " onclick="jsSaveForm()">
            &nbsp;
    		<input type="button" class="button" value=" 목록으로 " onclick="jsGotoList()">
            &nbsp;
    		<input type="button" class="button" value=" 선택삭제 " onclick="DelDetail()">
            &nbsp;
    		<input type="button" class="button" value=" 삭제하기 " onclick="DelMaster()">
            <% else %>
            전송완료 이후에는 수정할 수 없습니다.
            <input type="button" class="button" value=" 목록으로 " onclick="jsGotoList()">
            <% end if %>
		</td>
	</tr>
</table>
<form name="frmActPop" method="post" action="" style="margin:0px;">
<input type="hidden" name="suplyer" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidadd" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="refergubun" value="agvinterface">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
