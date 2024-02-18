<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim itemgubun,itemid, barcode, itemoption
barcode	  = request("barcode")

'response.write "<script>location.replace('/admin/offshop/popoffitemedit.asp?barcode=" + barcode + "');</script>"
'dbget.close()	:	response.End

itemgubun = Left(barcode,2)
itemid	  = CLng(Mid(barcode,3,6))
itemoption = Right(barcode,4)

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FRectItemgubun = itemgubun
ioffitem.FRectItemId = itemid
ioffitem.FRectItemOption = itemoption
ioffitem.GetOffOneItem

dim IsOnlineItem
IsOnlineItem = (itemgubun="10")

dim i
%>
<script language='javascript'>
function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function EditItemImage(frm){
	if (frm.file1.value.length<1){
		alert('이미지를 선택 하세요.');
		return;
	}

	if (confirm('이미지를 저장하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>
<% if (application("Svr_Info")="Dev") then %>	
<form name="frmedit" method=post action="http://testpartner.10x10.co.kr/linkweb/dooffitemimageedit.asp" enctype="MULTIPART/FORM-DATA">
<% else %>
<form name="frmedit" method=post action="http://partner.10x10.co.kr/linkweb/dooffitemimageedit.asp" enctype="MULTIPART/FORM-DATA">
<% end if %>
<input type=hidden name=barcode value="<%= barcode %>">
<input type=hidden name=itemgubun value="<%= itemgubun %>">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoption value="<%= itemoption %>">

<tr bgcolor="#DDDDFF">
	<td width=100>바코드</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.GetBarcode %>
	<br><font color="#AAAAAA">(90오프라인전용, 80이벤트 ,70소모품, 95가맹점개별매입판매)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>상품명</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<%= ioffitem.FOneItem.Fshopitemname %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>옵션명</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fshopitemoptionname %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>범용바코드</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fextbarcode %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>사용유무</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<% if ioffitem.FOneItem.Fisusing="Y" then %>
	사용함
	<% else %>
	사용안함
	<% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>판매가</td>
	<td bgcolor="#FFFFFF" colspan=5>
	    <%= FormatNumber(ioffitem.FOneItem.Fshopitemprice,0) %>
	</td>
	
</tr>

<tr bgcolor="#DDDDFF">
	<td width=100>등록일</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fregdate %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>최종수정일</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fupdt %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100 valign=top>오프상품<br>이미지</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=file name=file1 class="input_01" size=20 onchange="AttachImage(ioffimgmain,this)">(400 x 400 px)</div>
	<img name="ioffimgmain" src="<%= ioffitem.FOneItem.FOffImgMain %>" width=300 height=300>
	<br>

	<% if ioffitem.FOneItem.FOffImgList<>"" then %>
	<img src="<%= ioffitem.FOneItem.FOffImgList %>" width=100 height=100>
	<br>
	<% end if %>

	<% if ioffitem.FOneItem.FOffImgSmall<>"" then %>
	<img src="<%= ioffitem.FOneItem.FOffImgSmall %>" width=50 height=50>
	<% end if %>
	</td>
</tr>
</form>
<% if IsOnlineItem then %>
<script language='javascript'>
alert('오프라인 상품만 수정 가능합니다.');
window.close();
</script>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align=center><input type=button value=" 저  장 " onclick="EditItemImage(frmedit)" class="input_01"></td>
</tr>
<% end if %>
</table>
<%
set ioffitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->