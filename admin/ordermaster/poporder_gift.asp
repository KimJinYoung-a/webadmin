<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%

dim baljuid, page, isupchebeasong
dim research

baljuid         = request("baljuid")
page            = request("page")
isupchebeasong  = request("isupchebeasong")
research        = request("research")

if page="" then page=1
if research="" and isupchebeasong="" then isupchebeasong="N"

dim oOrderGift
set oOrderGift = new COrderGift
oOrderGift.FPageSize =1000
oOrderGift.FCurrPage = page
oOrderGift.FRectisupchebeasong = isupchebeasong
oOrderGift.FRectBaljuid = baljuid
oOrderGift.GetOrderGiftList

dim i
%>
<script language='javascript'>
function setGiftCode(gift_code){
    frmAct.gift_code.value = gift_code;
}

function ActReGiftMaker(frm){
    var gift_code = frm.gift_code.value;
    
    if (gift_code.length<1){
        alert('Gift 코드를 넣어주세요.');
        frm.gift_code.focus();
        return;
    }else{
        if (confirm('Gift(' + gift_code + ') 전체 사은품 내역을 재작성 하시겠습니까? \n\n미출고된 전체 주문건에 대해 재작성 됩니다.')){
            frm.submit();
        }
    }
}

function ActBaljuGift(frm){
    return;
    
    var baljuid = frm.baljuid.value;
    var evt_code = frm.evt_code.value;
    
    
    if (baljuid.length<1){
        alert('출고지시 번호를 입력 후 검색 하세요..');
        document.frm.baljuid.focus();
        return;
    }
    
    
    if (evt_code.length<1){
        //if (confirm('해당 출고지시(' + baljuid + ')의 전체 사은품 내역을 재작성 하시겠습니까?')){
        //    frm.submit();
        //}
        alert('이벤트 코드를 넣어주세요.');
        frm.evt_code.focus();
        return;
    }else{
        if (confirm('해당 출고지시(' + baljuid + ')의 이벤트(' + evt_code + ') 사은품 내역을 재작성 하시겠습니까?')){
            frm.submit();
        }
    }
}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			출고지시번호 : <input type="text" class="text" name="baljuid" value="<%= baljuid %>" size="10" maxlength="10">
        	&nbsp;
        	배송구분:
        	<input type="radio" name="isupchebeasong" value=""  <% if isupchebeasong="" then response.write "checked" %> >전체
        	<input type="radio" name="isupchebeasong" value="N" <% if isupchebeasong="N" then response.write "checked" %> >텐배
        	<input type="radio" name="isupchebeasong" value="Y" <% if isupchebeasong="Y" then response.write "checked" %> >업배
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frmAct" method="post" action="poporder_gift_process.asp" onSubmit="ActReGiftMaker(frmAct); return false;">
	<input type="hidden" name="baljuid" class="text_ro" size="6" value="<%= baljuid %>">
	<tr>
		<td align="left">
			출고지시번호 : <b><%= baljuid %></b>
			&nbsp;
            기프트번호 : <input type="text" name="gift_code" class="text" size="6" value="" >
            &nbsp;
            <% if (session("ssBctID")="icommang") or (session("ssBctID")="tozzinet") or (session("ssBctID")="coolhas") or (session("ssBctID")="kobula") then %>
            <input type="button" class="button" value="사은품목록재작성" onclick="ActReGiftMaker(frmAct);">
        	(tbl_order_gift)
        	<% end if %>
		</td>
		<td align="right">
		
		</td>
	</tr>
	</form>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">출고지시ID</td>
		<td width="80">주문번호</td>
		<td width="50">DAS</td>
		<td width="50">EVENT<br>ID</td>
		<td>이벤트명</td>
		<td width="50">GiftID</td>
		<td>사은품명</td>
		<td width="50">배송<br>구분</td>
		<td width="100">기간</td>
		<td>조건</td>
	</tr>
	<% for i=0 to oOrderGift.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><%= oOrderGift.FItemList(i).FBaljuID %></td>
	    <td><%= oOrderGift.FItemList(i).Forderserial %></td>
	    <td><%= oOrderGift.FItemList(i).Fdasindex %></td>
	    <td><%= oOrderGift.FItemList(i).Fevt_code %></td>
	    <td align="left"><%= oOrderGift.FItemList(i).Fevt_name %></td>
	    <td><a href="javascript:setGiftCode('<%= oOrderGift.FItemList(i).Fgift_code %>');"><%= oOrderGift.FItemList(i).Fgift_code %></a></td>
	    <td align="left"><%= oOrderGift.FItemList(i).Fgiftkind_name %></td>
	    <td>
	    <% if oOrderGift.FItemList(i).Fisupchebeasong="Y" then %>  
	    업체
	    <% else %>
	    텐배
	    <% end if %>  
	    </td>
	    
	    <td>
	        <%= oOrderGift.FItemList(i).Fevt_startdate %>
	        ~ <br>
	        <%= oOrderGift.FItemList(i).Fevt_enddate %>
	    </td>
	    <td>
	        <%= oOrderGift.FItemList(i).GetEventConditionStr %>
	    </td>
	</tr>
	<% next %>
</table>

<%
set oOrderGift = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->