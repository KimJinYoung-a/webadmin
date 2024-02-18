<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 재고파악
' History : 2010.04.02 한용민 수정 
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<%
Dim NowDate : NowDate=Left(now(),10)
Dim isWorkingState 
dim shopid, makerid, stTakingIdx
shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
stTakingIdx  = RequestCheckVar(request("stTakingIdx"),10)

if (stTakingIdx="") then stTakingIdx="0"

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

''업체
if (C_IS_Maker_Upche) then
    makerid = session("ssBctid")
end if

dim oOffStockTaking
set oOffStockTaking = new CStockTaking
oOffStockTaking.FRectShopID       = shopid
oOffStockTaking.FRectMakerID      = makerid

if (stTakingIdx<>"0") then
    oOffStockTaking.FRectIdx = stTakingIdx
    oOffStockTaking.getOneStockTaking
elseif ((shopid<>"") and (makerid<>"")) then
    oOffStockTaking.getRecentStockTaking
    
    if (oOffStockTaking.FResultCount>0) then stTakingIdx = oOffStockTaking.FOneItem.FstTakingIdx
end if

dim i

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if

isWorkingState = (stTakingIdx="0")
if (Not isWorkingState) then
    isWorkingState = oOffStockTaking.FOneItem.isWorkingState
end if
%>
<style>
#divView {width:100%;}
</style>
<script language='javascript'>
function goStockInput(stTakingIdx){
    location.href="/common/offshop/shop_brandcurrentstock_byjobkey.asp?idx="+stTakingIdx+"&sType=stTaking";
    
}

function playding(dingname){
    //alert(dingname);
    
    var v = document.getElementById(dingname);
    if (v.IsPlaying()){
        var v = document.getElementById(dingname+"2");
        //v.StopPlay();
    }
    v.Play(); //StopPlay
    
    document.frm.itembarcode.select();
    
    
}

// 키코드 접수
function keyCode(e) {
	var result;
	if(window.event)
		result = window.event.keyCode;
	else if(e)
		result = e.which;
	return result;
}



function getOnLoad(){
    document.frm.itembarcode.focus();
    <% if (stTakingIdx>"0") then %>
    jsView('<%= shopid %>','<%= makerid %>','');
    <% end if %>
}

window.onload=getOnLoad;

function selectFinish(ibarcode){
    document.frm.itembarcode.value=ibarcode;
    jsView('<%= shopid %>','<%= makerid %>',ibarcode);
}

function popOffItemList(comp){
    var frmname = comp.form.name;
    var compname = comp.name;
    var popUrl = "/common/offshop/popshopitemSelect.asp?makerid=<%= makerid %>&seltype=one&frmname="+frmname+"&compname="+compname;
    
    var popwin = window.open(popUrl,'popshopitemSelect','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function jsView(shopid,makerid,itembarcode){
    <% if (Not isWorkingState) then %>
    alert('재고파악중 상태가 아닙니다.');
    
    if (itembarcode!="") return;
    <% end if %>
	$.ajax({
		type: "POST",
		url: "/common/offshop/getBarcodeStockTaking.asp",
		data: "shopid="+shopid+"&makerid="+makerid+"&itembarcode="+itembarcode+"&stTakingIdx=<%= stTakingIdx %>",
		dataType: "text",
		//timeout : 1000,
		error: function(){
			html = "/common/offshop/getBarcodeStockTaking.asp?shopid="+shopid+"&makerid="+makerid+"&itembarcode="+itembarcode+"&stTakingIdx=<%= stTakingIdx %>";
			$("#divView").html(html);
		},
		success: function(html){
			$("#divView").html(html);
			//getHistory("CS",orderSerial,"");
		}
	});
}


//-------------------------------------------------------------------------------


function ModiPreStock(){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return;
    <% end if %>
    
    var frm = document.frmArr;
    var ischecked = false;
    var i = 0;
    
    if (!frm.cksel) return;
    
    if (frm.cksel.length){
        for (i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                ischecked = true;
                if (!IsInteger(frm.Arrrealstock[i].value)){
                    alert('정수만 가능합니다.');
                    frm.Arrrealstock[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            ischecked = true;
            if (!IsInteger(frm.Arrrealstock.value)){
                alert('정수만 가능합니다.');
                frm.Arrrealstock.focus();
                return;
            }
        }
    }
    
    if (!(ischecked)){
        alert('선택된 상품이 없습니다.');
        return;
    }
    
    if (confirm('재고파악 내역을 수정하시겠습니까?')){
        frm.submit();
    }
}


function nextStockStep(nVal){
    var frm = document.frmup;
    if (document.frmStockDt.stockdate){
        var stockdate = document.frmStockDt.stockdate.value;
    }else{
        var stockdate = "NULL"
    }
    frm.stStatus.value = nVal;
    frm.stockdate.value = stockdate;
    
    if (nVal==3){
        var confirmStr = "재고 반영 요청 하시겠습니까?";
    }else if (nVal==0){
        var confirmStr = "재고파악중 상태로 변경 하시겠습니까?";
    }else{
        var confirmStr = "수정 하시겠습니까?";
    }
    
    if (confirm(confirmStr)){
        frm.submit();
    }
}

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function popOffItemEdit(ibarcode){
    <% if C_IS_SHOP then %>
        return;
    <% elseif C_IS_Maker_Upche then %>
        var popwin = window.open('/designer/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
        popwin.focus();
    <% else %>
	    var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	    popwin.focus();
	<% end if %>
}

function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return; //업체위탁 상품인 경우?
    <% else %>
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
	    popwin.focus();
	<% end if %>
}

function PopOFFBrandStockSheet(){
    
    var shopid = document.frm.shopid.value;
    var makerid = document.frm.makerid.value;
    var centermwdiv = "";//document.frm.centermwdiv.value;
    var usingyn= document.frm.usingyn.value;
    
    if ((shopid.length<1)||(makerid.length<1)){
        alert('먼저 매장과 브랜드를 선택후 출력해 주세요.');
        return;
    }
    
    var popwin;
    
    popwin = window.open('/common/pop_offbrandstockprint.asp?shopid=' + shopid + '&makerid=' + makerid + '&centermwdiv=' + centermwdiv + '&usingyn=' + usingyn ,'pop_offbrandstockprint','width=1000,height=600,scrollbars=yes,resizable=yes')
    popwin.focus();
}

function CheckThis(i){
    var frm = document.frmArr;
    if (frm.cksel.length){
        frm.cksel[i].checked = true;
        AnCheckClick(frm.cksel[i]);
    }else{
        frm.cksel.checked = true;
        AnCheckClick(frm.cksel);
    }
}


</script>
<script language="javascript" src="/js/jquery-1.4.2.min.js"></script>

<OBJECT id=ding name=ding classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="1" HEIGHT="1">
 <PARAM NAME="movie" VALUE="/images/swf/ding.swf">
 <PARAM NAME="quality" VALUE="high">
 <PARAM NAME="bgcolor" VALUE="#FFFFFF">
 <param name="play" value="false">
 <EMBED src="ding.swf" quality="high" play="false" bgcolor="#FFFFFF" WIDTH="1" HEIGHT="1" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
</OBJECT>

<OBJECT id=ding2 name=ding2 classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="1" HEIGHT="1">
 <PARAM NAME="movie" VALUE="/images/swf/ding.swf">
 <PARAM NAME="quality" VALUE="high">
 <PARAM NAME="bgcolor" VALUE="#FFFFFF">
 <param name="play" value="false">
 <EMBED src="ding.swf" quality="high" play="false" bgcolor="#FFFFFF" WIDTH="1" HEIGHT="1" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
</OBJECT>

<OBJECT id=chord name=chord classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="1" HEIGHT="1">
 <PARAM NAME="movie" VALUE="/images/swf/chord.swf">
 <PARAM NAME="quality" VALUE="high">
 <PARAM NAME="bgcolor" VALUE="#FFFFFF">
 <param name="play" value="false">
 <EMBED src="chord.swf" quality="high" play="false" bgcolor="#FFFFFF" WIDTH="1" HEIGHT="1" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
</OBJECT>

<OBJECT id=chord2 name=chord2 classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="1" HEIGHT="1">
 <PARAM NAME="movie" VALUE="/images/swf/chord.swf">
 <PARAM NAME="quality" VALUE="high">
 <PARAM NAME="bgcolor" VALUE="#FFFFFF">
 <param name="play" value="false">
 <EMBED src="chord.swf" quality="high" play="false" bgcolor="#FFFFFF" WIDTH="1" HEIGHT="1" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
</OBJECT>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" onsubmit="return false;">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    <input type="hidden" name="shopid" value="<%= shopid %>">
		    <input type="hidden" name="makerid" value="<%= makerid %>">
		    매장 : <strong><%= shopid %></strong>
		    &nbsp;&nbsp;&nbsp;&nbsp;
		    브랜드 : <strong><%= makerid %></strong>
		    &nbsp;&nbsp;&nbsp;&nbsp;
			상품바코드 : 
			<input type="text" class="text" name="itembarcode" value="" size="20" maxlength="32" onKeyPress="if (keyCode(event) == 13) jsView('<%= shopid %>','<%= makerid %>',this.value); ">
			
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="수동검색" onClick="popOffItemList(frm.itembarcode);">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<% if (oOffStockTaking.FResultCount<1) then %>
		    <span id="jobText">재고작업 내역 없음.</span>
		<% else %>
		    <span id="jobText">
		        작업번호 : <strong><%= oOffStockTaking.FOneItem.FstTakingIdx %></strong>
		        &nbsp;
		        최초작업자 : <%= oOffStockTaking.FOneItem.FregUserID %>
		        &nbsp;
		        작업상태 : <%= oOffStockTaking.FOneItem.getStatusName %>
		    </span>
		    
		<% end if%>
		</td>
	</tr>
	
	</form>
</table>
<!-- 검색 끝 -->


<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" > 
    <form name="frmStockDt">
	<tr height="30">
		<td align="left">
			<% if (C_ADMIN_AUTH=true) then %>
			
	        <% end if %>
	        
        	
		    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
                (업체위탁 계약 매장만 재고 수정 가능)
            <% else %>
                <% if CStr(stTakingIdx)<>"0" then %>
                    <% if (oOffStockTaking.FOneItem.FstStatus=0) then %>
                    <input type="button" class="button" value="선택상품 재고파악 내역 수정" onClick="ModiPreStock();">
                    <% elseif (oOffStockTaking.FOneItem.FstStatus=3) then %>
                    <input type="button" class="button" value="재고파악중 상태로 변경" onClick="nextStockStep(0);">
                    <% end if %>
                <% end if %> 
            <% end if %> 
		</td>
		<td align="right">
		    <% if CStr(stTakingIdx)<>"0" then %>
    		    <% if (oOffStockTaking.FOneItem.FstStatus=0) then %>
    		    재고파악일시 : <input type="text" class="text" name="stockdate" value="<%= NowDate %>" size=11 readonly ><a href="javascript:calendarOpen(frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
    			<input type="button" class="button" name="stock_sheet_print" value="재고파악 완료 및 재고 반영 요청" onclick="nextStockStep(3);"> 
    			<% elseif (oOffStockTaking.FOneItem.FstStatus=3) then %>
			    <input type="button" class="button" name="stock_sheet_print" value="재고입력 으로 이동" onclick="goStockInput(<%= stTakingIdx %>);"> 
			    <% end if %>
			<% end if %>
		</td>
	</tr>
	</form>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td align="right">
    </td>
</tr>
</table>



<div id="divList"></div>

<div id="divView"></div>


<%
set oOffStockTaking = Nothing
%>
<form name="frmup" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="stockTakingNext">
<input type="hidden" name="stTakingIdx" value="<%= stTakingIdx %>">
<input type="hidden" name="stStatus" value="">
<input type="hidden" name="stockdate" value="">
</form>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" --> 