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
	reStr = reStr + "width=115" + vbCrlf
	reStr = reStr + "height=10" + vbCrlf
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


const Cols = 4
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


<table width="745" border="1" cellpadding="0" cellspacing="0" align=center bordercolor="#CCCCCC" >
<% if Rows=0 then %>
<tr>
	<td align="center" class="a"> 검색 결과가 없습니다.</td>
</tr>
<% else %>
<tr height="11">
	<td colspan=4 align="center"></td>
</tr>
<% for i=0 to Rows-1 %>
<% if (i<>0) and ((i mod 9)=0) then %>
</table>

<div CLASS="break"></div>

<table width="745" border="1" cellpadding="0" cellspacing="0" align=center bordercolor="#CCCCCC" >
<tr height="11">
	<td colspan=4 align="center"></td>
</tr>
<tr height="102" align="center">
<% else %>
<tr height="102" align="center">
<% end if %>
	<% for j=0 to Cols-1 %>
	<% if i*Cols+j>obarcode.Fresultcount-1 then %>
	<td></td>
	<% else %>
	<td valign="center">
		<table width="165" height="85" align="center" valign="center" border="0" cellspacing="0" cellpadding="0">
		<tr height=50 valign="top">
			<td width="50" align="left">
				<table width="50" height="50" border="0" cellpadding="0" cellspacing="0" >
				<tr>
				  <% if obarcode.FItemList(i*Cols+j).FImageSmall<>"" then %>
				  <td><img src="<%= obarcode.FItemList(i*Cols+j).FImageSmall %>" width="50" height="50" border="1"></td>
				  <% else %>
				  <td><img src="http://image.10x10.co.kr/image/whiteblank.gif" width="50" height="50" border="1" ></td>
				  <% end if %>
				</tr>
				</table>
			</td>
			<td width="120" align="left">
				<font size="1">
				<b><%= Left(obarcode.FItemList(i*Cols+j).FShopItemName,36) %></b>
				<br>
				<% if obarcode.FItemList(i*Cols+j).FShopItemOptionName<>"" then %>
				<%= Left(obarcode.FItemList(i*Cols+j).FShopItemOptionName,36) %>
				<br>
				<% end if %>
				</font>
			</td>
		</tr>
		<tr height=5>
			<td colspan=2 valign=top>
				<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="105"><font size="1"><b><%= obarcode.FItemList(i*Cols+j).Fmakerid %></b></font></td>
					<td width="60" align="right"><font size="1"><b><%= FormatNumber(obarcode.FItemList(i*Cols+j).Fshopitemprice,0) %></b></font></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height=25 valign="top">
			<td colspan=2 align="center">
				<%= getObjStr(i*Cols+j) %>
				<br>
			  	<font size="1">
			  	<%= obarcode.FItemList(i*Cols+j).GetBarCode %> &nbsp&nbsp
			  	</font>
			  	<b><%= obarcode.FItemList(i*Cols+j).FShopItemid %></b>

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