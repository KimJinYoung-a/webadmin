<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
rw "사용중지메뉴-관리자문의 요망"
response.end

dim itembarcode
itembarcode = requestCheckVar(request("itembarcode"),20)

dim itemgubun, itemid, itemoption

'if (Len(itembarcode)=12) then
'	itemgubun 	= left(itembarcode,2)
'	itemid		= CLng(mid(itembarcode,3,6))
'	itemoption	= right(itembarcode,4)
'else
'	itemgubun = "10"
'	itemid = itembarcode
'	itemoption  = "0000"
'end if

if BF_IsMaybeTenBarcode(itembarcode) then
    itemgubun 	= BF_GetItemGubun(itembarcode)
	itemid 		= BF_GetItemId(itembarcode)
	itemoption 	= BF_GetItemOption(itembarcode)
else
	itemgubun = "10"
	itemid = itembarcode
	itemoption  = "0000"
end if


dim oitembar
set oitembar = new CItemBarCode
oitembar.FRectItemGubun = itemgubun
oitembar.FRectItemID = itemid
'''oitembar.FRectItemoption = itemoption
if itemid<>"" then
	oitembar.getItemBarcodeInfo
end if


dim i
%>
<script language='javascript'>
function InputRackcode(frm){
	if (frm.itemrackcode.value.length!=4){
		alert('상품 랙코드를 정확히 입력하세요. 4자리');
		frm.itemrackcode.focus();
		return;
	}

	if (confirm('상품 랙코드를 저장하시겠습니까?')){
		frm.submit();
	}
}

function Research(frm){
	frm.submit();
}

function InputBarcode(comp){
	var barcode = comp.value;
	var frm = document.frmsavebar;

	if ((barcode.substr(0,2)=='10')&&(barcode.length==12)){
		alert('등록할 수 없는 바코드 형식입니다.');
		comp.focus();
		return;
	}

	if (barcode.length<8){
		alert('바코드를 정확히 입력하세요.');
		comp.focus();
		return;
	}

	if ((frm.itemgubun.value!="10")&&(frm.itemgubun.value!="90")&&(frm.itemgubun.value!="70")){
		alert('상품구분오류 - 관리자 문의요망');
		return;
	}

	if (frm.itemid.value.length<1){
		alert('상품코드가 정의되지 않았습니다. - 상품검색후 사용하세요.');
		document.frmbar.itembarcode.focus();
		return;
	}

	frm.itemoption.value = comp.id;
	frm.publicbarcode.value = barcode;

    if (confirm('범용 바코드를 저장하시겠습니까?')){
		frm.submit();
	}

/*	??
	if (frm.confirmbarcode.value.length<1){
	    alert('확인을 위해 다시한번 바코드를 입력해주세요.');
	    frm.confirmbarcode.value = barcode;
	    comp.value ='';
	    comp.focus();

	}else{
	    if (frm.confirmbarcode.value==frm.publicbarcode.value){
        	if (confirm('범용 바코드를 저장하시겠습니까?')){
        		frm.submit();
        	}
        }else{
            frm.confirmbarcode.value = "";
            frm.publicbarcode.value = "";

            alert('바코드가 일치하지 않습니다. 처음부터 다시 시도해 주세요.');
            comp.value ='';
            comp.focus();
        }
    }
*/
}

function GetOnLoad(){
	<% if oitembar.FResultCount>0 then %>
	    if (document.frmbar.publicbar_<%= itemoption %>) {
	        document.frmbar.publicbar_<%= itemoption %>.select();
    	    document.frmbar.publicbar_<%= itemoption %>.focus();
        }

	<% else %>
	    document.frmbar.itembarcode.select();
	    document.frmbar.itembarcode.focus();
	<% end if %>
}
window.onload=GetOnLoad;
</script>

  <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<form name="frmbar" method=get>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;<strong>상품바코드입력</strong></font>
				    </td>
				    <td align="right">
						<input type="text" class="text"  name="itembarcode" value="<%= itembarcode %>" size=17 maxlength=14 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
        				<input type="button" class="button" value="검색" onclick="Research(frmbar)" >
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- 상단바 끝 -->

<% if oitembar.FResultCount>0 then %>
  	<tr bgcolor="#FFFFFF">
    	<td width="80" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
   	<td colspan="2"><%= oitembar.FItemList(0).FbrandName %>(<%= oitembar.FItemList(0).Fmakerid %>)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
    	<td colspan="2"><%= oitembar.FItemList(0).FItemName %></td>
    </tr>
	<% for i=0 to oitembar.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">옵션명(<%= oitembar.FItemList(i).Fitemoption %>)</td>
		<% if oitembar.FItemList(i).FitemoptionName="" then %>
		<td>옵션없음</td>
		<% else %>
			<% if itemoption=oitembar.FItemList(i).Fitemoption then %>
			<td><b><%= oitembar.FItemList(i).FitemoptionName %></b></td>
			<% else %>
			<td><%= oitembar.FItemList(i).FitemoptionName %></td>
			<% end if %>
		<% end if %>

		<td align="right">
		<% if oitembar.FItemList(i).Fitemoption=itemoption then %>
			<input type="text" class="text" id="<%= oitembar.FItemList(i).Fitemoption %>" name="publicbar_<%= oitembar.FItemList(i).Fitemoption %>" value="<%= oitembar.FItemList(i).FPublicBarcode %>" size=20 maxlength=20 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>); return false;}">
			<input type="button" class="button" value="등록" onclick="InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>)">
		<% else %>
		    <input type="text" class="text" id="<%= oitembar.FItemList(i).Fitemoption %>" name="publicbar_<%= oitembar.FItemList(i).Fitemoption %>" value="<%= oitembar.FItemList(i).FPublicBarcode %>" size=20 maxlength=20 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>); return false;}" disabled >
			<input type="button" class="button" value="등록" onclick="InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>)" disabled >
		<% end if %>
		</td>
	<% next %>
	</tr>
	<tr bgcolor="#FFFFFF">
    	<td bgcolor="<%= adminColor("tabletop") %>">이미지</td>
    	<td colspan="2"><img src="<%= oitembar.FItemList(0).FImageList %>" width="100" height="100" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
    </tr>
	</form>
	<!--
	<form name="frmitemrackcode" method=post  action="/warehouse/itemrackcode_process.asp">
	<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="mode" value="byitem">
    <tr bgcolor="#FFFFFF">
    	<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
    	<td colspan="2"><%= FormatNumber(oitembar.FItemList(0).FSellcash,0) %></td>
    </tr>

    <tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">상품랙코드</td>
    	<td colspan="2">
    		<input type="text" class="text" name="itemrackcode" value="<%= oitembar.FItemList(0).Fitemrackcode %>" size="4" maxlength="4" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputRackcode(frmitemrackcode); return false;}">
    		<input type="button" class="button" value="저장" onclick="InputRackcode(frmitemrackcode);">
    		&nbsp;
    		(브랜드랙코드 : <%= oitembar.FItemList(0).Fprtidx %>)
    	</td>
    </tr>
    </form>
    -->
<% else %>
	<tr bgcolor="#FFFFFF">
    	<td colspan="3" align="center">
    		검색결과가 없습니다

    		<!-- <br>
    		현재 10코드(온라인등록상품)만 등록이 가능합니다.
    		<br>90코드의 경우 오프상품관리를 이용하세요. -->
    	</td>
    </tr>
<% end if %>


</table>


<%
set oitembar = Nothing
%>
<form name="frmsavebar" method=post action="barcode_input_process.asp">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="publicbarcode" value="">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->