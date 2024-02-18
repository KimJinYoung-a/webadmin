<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%

function getObjStr(v)
	dim reStr
	reStr = "<OBJECT" + vbCrlf
	reStr = reStr + "id=iaxobject" + CStr(v)  + vbCrlf
	reStr = reStr + "classid='clsid:A4F3A486-2537-478C-B023-F8CCC41BF29D'" + vbCrlf
	reStr = reStr + "codebase='http://partner.10x10.co.kr/cab/tenbarShow.cab#version=1,0,0,3'" + vbCrlf
	reStr = reStr + "width=130" + vbCrlf
	reStr = reStr + "height=18" + vbCrlf
	reStr = reStr + "align=center" + vbCrlf
	reStr = reStr + "hspace=0" + vbCrlf
	reStr = reStr + "vspace=0" + vbCrlf
	reStr = reStr + ">" + vbCrlf
	reStr = reStr + "</OBJECT>" + vbCrlf

	getObjStr = reStr
end function

dim designer, onlyipgo, only90, only10
designer = request("designer")
onlyipgo = request("onlyipgo")
only90 = request("only90")
only10 = request("only10")


dim obarcode
set obarcode = new COffShopItem
obarcode.FPageSize=200
obarcode.FRectDesigner = designer
obarcode.FRectIpGoOnly = onlyipgo
obarcode.FRectOnly90 = only90
obarcode.FRectOnly10 = only10

if (designer<>"") and ((only90="on") or (only10="on"))  then
	obarcode.GetShopPrintList
end if

const Cols = 3
dim i,j

%>

<table width="700" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			디자이너:<% drawSelectBoxDesignerwithName "designer",designer  %>
			&nbsp;&nbsp;
			<input type="checkbox" name="onlyipgo" <% if onlyipgo="on" then response.write "checked" %>>입고된 내역만
			&nbsp;
			<input type="checkbox" name="only90" <% if only90="on" then response.write "checked" %>>샾전용상품
			<input type="checkbox" name="only10" <% if only10="on" then response.write "checked" %>>온라인상품
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<%
dim Rows
Rows = (obarcode.Fresultcount) \ Cols
if ((obarcode.Fresultcount) mod Cols)>0 then Rows= Rows+1
%>
<table width="700" border="0" cellpadding="5" cellspacing="0" >
<tr><td align="right"><a href="javascript:window.print();">print</a></td></tr>
</table>
<table width="650" border="0" cellpadding="0" cellspacing="0" bordercolor="#999999" >
<% if Rows=0 then %>
<tr height="80">
<td align="center" class="a"> 검색 결과가 없습니다.</td>
</tr>
<% else %>
<% for i=0 to Rows-1 %>
<tr height="80">
	<% for j=0 to Cols-1 %>
	<% if i*Cols+j>obarcode.Fresultcount-1 then %>
	<td></td>
	<% else %>
	<td  valign="top">
		<table border=0 cellspacing=0 width="50">
		<tr valign="top">
			<td valign="top">
				<table width="50" height="50" border="0" cellpadding="1" cellspacing="0" bordercolor="#999999">
				<tr>
				  <% if obarcode.FItemList(i*Cols+j).FImageSmall<>"" then %>
				  <td><img src="<%= obarcode.FItemList(i*Cols+j).FImageSmall %>" width="50" height="50" border="1" bordercolor="gray"></td>
				  <% else %>
				  <td width="50" height="50" ><img src="http://image.10x10.co.kr/image/whiteblank.gif" width="50" height="50" border="1" bordercolor="gray"></td>
				  <% end if %>
				</tr>
				</table>
			</td>
			<td width="30">&nbsp;&nbsp;</td>
			<td valign="top">
			  	  <table width="126" cellspacing="0" border=0 cellpadding="0" >
			  	  <tr>
			  	  	<td>
			  	  		<%= getObjStr(i*Cols+j) %>
			  	  	</td>
			  	  </tr>
			  	  <tr>
			  	  	<td>
			  	  		<font size=1">
			  	  		<%= obarcode.FItemList(i*Cols+j).GetBarCode %>
			  	  		<br>
			  	  		<br>
			  	  		<%= obarcode.FItemList(i*Cols+j).FShopItemName %>
			  	  		<br>
			  	  		<% if obarcode.FItemList(i*Cols+j).FShopItemOptionName<>"" then %>
			  	  		<%= obarcode.FItemList(i*Cols+j).FShopItemOptionName %>
			  	  		<br>
			  	  		<% end if %>
			  	  		<%= FormatNumber(obarcode.FItemList(i*Cols+j).Fshopitemprice,0) %>
			  	  		</font>
			  	  	</td>
			  	  </tr>
			  	  </table>
			  </td>
		</tr>
		</table>
	</td>
	<% end if %>
	<% next %>
</tr>
<% next %>
<% end if %>
</table>

<script language='javascript'>
<% for i=0 to obarcode.Fresultcount-1 %>
iaxobject<%= i %>.ShowBarCode(30,'<%= obarcode.FItemList(i).getbarCode %>',2);
<% next %>
</script>

<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->