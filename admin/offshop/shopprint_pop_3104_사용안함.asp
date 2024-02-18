<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%

function getObjStr(v)
	dim reStr
	reStr = "<OBJECT" + vbCrlf
	reStr = reStr + "id=iaxobject" + CStr(v)  + vbCrlf
	reStr = reStr + "classid='clsid:A4F3A486-2537-478C-B023-F8CCC41BF29D'" + vbCrlf
	reStr = reStr + "codebase='http://partner.10x10.co.kr/cab/tenbarShow.cab#version=1,0,0,3'" + vbCrlf
	reStr = reStr + "width=113" + vbCrlf
	reStr = reStr + "height=15" + vbCrlf
	reStr = reStr + "align=top" + vbCrlf
	reStr = reStr + "hspace=0" + vbCrlf
	reStr = reStr + "vspace=0" + vbCrlf
	reStr = reStr + ">" + vbCrlf
	reStr = reStr + "</OBJECT>" + vbCrlf

	getObjStr = reStr
end function

dim designer, notupche, onoffgubun, ckall
designer = request("designer")
notupche = request("notupche")
onoffgubun = request("onoffgubun")
ckall = request("ckall")


ckall = replace(ckall,", ","','")
ckall = "'" + ckall + "'"


dim obarcode
set obarcode = new COffShopItem
obarcode.FPageSize = 200
obarcode.FRectDesigner = designer
obarcode.FRectonoffgubun = onoffgubun
obarcode.FRectOnlyTenbeasong = notupche
obarcode.FRectArraryBarCode = ckall

obarcode.GetShopPrintList


const Cols = 3
dim i,j

%>

<%
dim Rows
Rows = (obarcode.Fresultcount) \ Cols
if ((obarcode.Fresultcount) mod Cols)>0 then Rows= Rows+1
%>

<STYLE TYPE="text/css">
<!-- .break {page-break-before: always;} -->
</STYLE>


<table width="740" border="1" cellpadding="0" cellspacing="0" align=center bordercolor="#CCCCCC" >
<% if Rows=0 then %>
<tr>
	<td align="center" class="a"> 검색 결과가 없습니다.</td>
</tr>
<% else %>
<!--
<tr height="11">
	<td colspan=3 align="center"></td>
</tr>
-->
<% for i=0 to Rows-1 %>
<% if (i<>0) and ((i mod 9)=0) then %>
</table>

<div CLASS="break"></div>

<table width="740" border="1" cellpadding="0" cellspacing="0" align=center bordercolor="#CCCCCC" >
<tr height="1">
	<td colspan=3 align="center"></td>
</tr>
<tr height="114" align="center">
<% else %>
<tr height="114" align="center">
<% end if %>
	<% for j=0 to Cols-1 %>
	<% if i*Cols+j>obarcode.Fresultcount-1 then %>
	<td></td>
	<% else %>
	<td valign="center">
		<table width="220" height="67" align="center" valign="center" border="0" cellspacing="0" cellpadding="0">
		<tr height=67 valign="top">
			<td width="67" align="left">
				<table width="67" height="67" border="0" cellpadding="0" cellspacing="0" >
				<tr>
				  <% if obarcode.FItemList(i*Cols+j).FImageList<>"" then %>
				  <td><img src="<%= obarcode.FItemList(i*Cols+j).FImageList %>" width="67" height="67" border="1"></td>
				  <% else %>
				  <td><img src="http://image.10x10.co.kr/image/whiteblank.gif" width="67" height="67" border="1" ></td>
				  <% end if %>
				</tr>
				</table>
			</td>
			<td align="left">
				<table width="153" height="69" align="center" valign="center" border="0" cellspacing="0" cellpadding="0">
				<tr width="153" valign="top">
					<td height="44" align="left">
					<font size="1">
					<b><%= Left(obarcode.FItemList(i*Cols+j).FShopItemName,50) %></b>
					<br>
					<% if obarcode.FItemList(i*Cols+j).FShopItemOptionName<>"" then %>
					<%= Left(obarcode.FItemList(i*Cols+j).FShopItemOptionName,50) %>
					<br>
					<% end if %>
					</font>
					</td>
				</tr>
				<tr>
					</td>
					<td height="25" align="right" valign="bottom">
					<font size="1">
					<b><%= obarcode.FItemList(i*Cols+j).Fmakerid %></b><br>
					<b><%= FormatNumber(obarcode.FItemList(i*Cols+j).Fshopitemprice,0) %></b>
					</font>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="15">
			<td align="left" valign="bottom">
				<b><strong><%= obarcode.FItemList(i*Cols+j).FShopItemid %></strong></b>
			</td>
			<td align="right" valign="bottom">
				<font size="1">
			  	<%= obarcode.FItemList(i*Cols+j).GetBarCode %>
			  	</font>
			  	<br>
				<%= getObjStr(i*Cols+j) %>
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
//window.onload=window.print();
</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->