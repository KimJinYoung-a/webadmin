<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->

<%
'Matchitemoption 만 업데이트
Dim mallid, itemid, vIsOption, mode, itemoption, outmallitemid, outmallitemname, outmallPrice, outmallSellYn
Dim itemOptionname, outmallitemOptionname, itemname, sellyn, sellcash, i
	mallid	= requestCheckvar(request("mallid"),32)
	itemid	= requestCheckvar(request("itemid"),10)
	itemoption= requestCheckvar(request("itemoption"),4)

Dim oxItem
set oxItem = new CxSiteTempLinkItem
	oxItem.FRectItemID = itemid
	oxItem.FRectSellSite = mallid
	oxItem.FRectItemOption= itemoption
	
	If itemid <> "" Then
		oxItem.getOnexSiteTempLinkItem
	End If

if oxItem.fresultCount>0 then 
    mode="edit"
    outmallitemid   = oxItem.FOneItem.Foutmallitemid
    outmallitemname = oxItem.FOneItem.Foutmallitemname
    outmallitemOptionname = oxItem.FOneItem.FoutmallitemOptionname
    outmallPrice    = oxItem.FOneItem.FoutmallPrice
    outmallSellYn   = oxItem.FOneItem.FoutmallSellYn
    itemname        = oxItem.FOneItem.Fitemname
    itemOptionname  = oxItem.FOneItem.FitemOptionname
    sellyn          = oxItem.FOneItem.Fsellyn
    sellcash        = oxItem.FOneItem.Fsellcash
end if

if (itemid="") then
    mode="add"
end if
%>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script language="javascript" src="/js/jquery-1.6.2.min.js"></script>
<script language="javascript">

function selThisItem(iitemid){
    var frm = document.frmXItem;
    var selopt ='';

	selopt = $("#vOpt").val();
    if (selopt){
        if (selopt.length!=4){
            alert('옵션을 선택 하세요.');
            return;
        }
    }else{
        selopt = '0000';
    }
    
    frm.itemid.value=iitemid;
    frm.itemoption.value=selopt;
    
}

function searchItem(frm){
    var xitemname = escape(frm.outmallitemname.value);
    //var xitemname = encodeURI(frm.outmallitemname.value);
    //var xitemname = frm.outmallitemname.value;
    if (xitemname.length<1){
        alert('상품명을 입력후 검색하세요.');
        frm.outmallitemname.focus();
        return;
    }
    
    $("#divView").html('');
    
    $.ajax({
		type: "POST",
		url: "/admin/etc/orderInput/ajxMatchXsiteItem.asp",
		data: "outmallitemname="+xitemname,
		dataType: "text",
		//timeout : 1000,
		error: function(){
			html = "/admin/etc/orderInput/ajxMatchXsiteItem.asp?"+"outmallitemname="+xitemname;
			$("#divView").html(html);
		},
		success: function(html){
			$("#divView").html(html);
			
		}
	});
}


function searchItem2(frm){
    var tenitemid = escape(frm.itemid.value);
    if (tenitemid.length<1){
        alert('상품코드 입력후 검색하세요.');
        frm.itemid.focus();
        return;
    }
    
    $("#divView").html('');
    
    $.ajax({
		type: "POST",
		url: "/admin/etc/orderInput/ajxMatchXsiteItem.asp",
		data: "tenitemid="+tenitemid,
		dataType: "text",
		//timeout : 1000,
		error: function(){
			html = "/admin/etc/orderInput/ajxMatchXsiteItem.asp?"+"tenitemid="+tenitemid;
			$("#divView").html(html);
		},
		success: function(html){
			$("#divView").html(html);
			
		}
	});
}
function ModiXItem(){
	var frm = document.frmXItem;
	
	if (frm.itemid.value.length<1){
	    alert('TEN 상품번호 필수 입니다.')
	    frm.itemid.focus();
	    return;
	}
	
	if ((frm.outmallitemid.value.length<1)&&(frm.outmallitemname.value.length<1)){
	    alert('제휴 상품번호 또는 제휴 상품명 중 하나는 필수 입력 값입니다.')
	    frm.outmallitemid.focus();
	    return;
	}
	
	//반디앤루이스 제외 필수 값.
	if (frm.itemoption.value.length!=4){
	    if (frm.mallid.value!="bandinlunis11111111" ){			//&& frm.mallid.value!="mintstore"
	        alert('bandinlunis, mintstore 몰 제외 옵션코드 필수값.')
    	    frm.itemoption.focus();
    	    return;
	    }
	}
	
	if (frm.outmallPrice.value.length<1){
	    alert('제휴 판매가격 필수 입니다.')
	    frm.outmallPrice.focus();
	    return;
	}
	
	if ((!frm.outmallSellYn[0].checked)&&(!frm.outmallSellYn[1].checked)&&(!frm.outmallSellYn[2].checked)){
	    alert('제휴 판매 여부를 선택하세요.')
	    frm.outmallSellYn[0].focus();
	    return;
	}
	
	if (confirm('저장 하시겠습니까?')){
	    frm.submit();
	}
}

function DelXItem(){
    var frm = document.frmXItem;
	
	if (frm.itemid.value.length<1){
	    alert('TEN 상품번호 필수 입니다.')
	    frm.itemid.focus();
	    return;
	}
	
	if (confirm('삭제 하시겠습니까?')){
	    frm.mode.value="del";
	    frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
<form name="frmXItem" method="post" action="partneritemlink_process.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr>
	<td width="120" align="right" class="td_br_tablebar">몰 구분:</td>
	<td class="td_br" colspan="2"><%= mallid %></td>
</tr>
<tr>
	<td width="120" align="right" class="td_br_tablebar">TEN 상품번호:</td>
	<td class="td_br" colspan="2">
	<% if mode="add" then %>
	    <input type="text" name="itemid" value="" size="6" maxlength="9">(필수) <input type="button" value="Search" onClick="searchItem2(document.frmXItem);">
	<% else %>
	    <input type="text" name="itemid" value="<%= oxItem.FOneItem.FItemId %>" size="6" maxlength="9" readonly class="text_ro">
	<% end if %>
	</td>
</tr>
<% 
if mallid="bandinlunis11111111"  then 		'/or mallid="mintstore"
%>
    <input type="hidden" name="itemoption">
    <input type="hidden" name="p_itemoption">
<% else %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">TEN 옵션번호:</td>
	<td class="td_br" colspan="2">
	<% if mode="add" then %>
	    <input type="text" name="itemoption" value="" size="6" maxlength="9">
	<% else %>
	    <input type="hidden" name="p_itemoption" value="<%= oxItem.FOneItem.Fitemoption %>">
	    <input type="text" name="itemoption" value="<%= oxItem.FOneItem.Fitemoption %>" size="4" maxlength="4" >
	<% end if %>
	( 필수 )
	</td>
</tr>
<% end if %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">제휴 상품번호:</td>
	<td class="td_br" colspan="2">
	    <input type="text" name="outmallitemid" value="<%= outmallitemid %>" size="20" maxlength="20">
	    (주문 리스트 엑셀에 제휴 상품번호가 있는경우 필수 입력)
	</td>
</tr>
<tr>
	<td width="120" align="right" class="td_br_tablebar">제휴 상품명:</td>
	<td class="td_br">
	    <input type="text" name="outmallitemname" value="<%= outmallitemname %>" size="40" maxlength="50">
	    (주문 리스트 엑셀에 제휴 상품번호가 없는경우 필수 입력)
	    <% if mode="add" then %>
	    <br>
	    <table border=0 cellspacing=2 cellpadding=2>
	    <tr>
	        <td><input type="button" value="Search" onClick="searchItem(document.frmXItem);"></td>
	        <td><div id="divView"></div></td>
	    </tr>
	    </table>
	    <% end if %>
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.Fitemname %></td>
	<% end if %>
</tr>
<% 
if mallid="bandinlunis11111111" then 		'/ or mallid="mintstore"
%>
    <input type="hidden" name="outmallitemOptionname" >
<% else %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">제휴 옵션명:</td>
	<td class="td_br">
	    <input type="text" name="outmallitemOptionname" value="<%= outmallitemOptionname %>" size="40" maxlength="50">
	   
	    <% if mallid="hottracks" then %>
	    	<br>교보핫트랙스의 경우 엑셀 상품명에 있는 상품명과 옵션명을 나누어서 어드민에 입력해주세요
	    <% end if %>
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.FitemOptionname %></td>
	<% end if %>
</tr>	
<% end if %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">제휴 판매가:</td>
	<td class="td_br">
	    <input type="text" name="outmallPrice" value="<%= outmallPrice %>" size="10" maxlength="10">
	    (관리상 필요 - 필수)
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.Fsellcash %> &nbsp;</td>
	<% end if %>
</tr>		
<tr>
	<td width="120" align="right" class="td_br_tablebar">제휴 판매여부:</td>
	<td class="td_br">
	    <input type="radio" name="outmallSellYn" value="Y" <%= CHKIIF(outmallSellYn="Y","checked","") %> >판매중
	    <input type="radio" name="outmallSellYn" value="N" <%= CHKIIF(outmallSellYn="N","checked","") %> >판매안함
	    <input type="radio" name="outmallSellYn" value="X" <%= CHKIIF(outmallSellYn="X","checked","") %> >판매종료
	    (관리상 필요 - 필수)
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.Fsellyn %> &nbsp;</td>
	<% end if %>
</tr>	
<tr>
	<td align="center" colspan="3" class="td_br">
	<% If mode = "add" Then %>
	    <input type="button" class="button" value="추가" onClick="ModiXItem();">
	<% Else %>
		<input type="button" class="button" value="수정" onClick="ModiXItem()">
		&nbsp;
		<input type="button" class="button" value="삭제" onClick="DelXItem()">
		&nbsp;
		<input type="button" class="button" value="닫기" onClick="self.close()">
	<% End If %>
	</td>
</tr>
</form>	
</table>

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=200>
		제휴상품명과 제휴옵션명으로<br>매칭하는 제휴몰 : 
	</td>
	<td align="left">
		<% GetItemMaeching_itemname_itemoptionname_list() %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<%
SET oxItem= Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->