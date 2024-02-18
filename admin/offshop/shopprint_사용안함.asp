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
	reStr = reStr + "height=35" + vbCrlf
	reStr = reStr + "align=top" + vbCrlf
	reStr = reStr + "hspace=0" + vbCrlf
	reStr = reStr + "vspace=2" + vbCrlf
	reStr = reStr + ">" + vbCrlf
	reStr = reStr + "</OBJECT>" + vbCrlf

	getObjStr = reStr
end function

dim designer, notupche
dim onoffgubun
designer = request("designer")
notupche = request("notupche")
onoffgubun = request("onoffgubun")

dim obarcode
set obarcode = new COffShopItem
obarcode.FPageSize = 200
obarcode.FRectDesigner = designer
obarcode.FRectonoffgubun = onoffgubun
obarcode.FRectOnlyTenbeasong = notupche

obarcode.GetShopPrintList

const Cols = 4
dim i,j

%>
<script language='javascript'>
function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function shopprint_pop(designer,onlyipgo,only90,only10,onlynew)
{
	var popwin = window.open("","shopprint_pop","width=800 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
	frmarr.action = "/admin/offshop/shopprint_pop.asp";
	frmarr.target = "shopprint_pop";
	frmarr.submit();
}

function shopprint_pop_3104(designer,onlyipgo,only90,only10,onlynew)
{
	var popwin = window.open("","shopprint_pop","width=800 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
	frmarr.action = "/admin/offshop/shopprint_pop_3104.asp";
	frmarr.target = "shopprint_pop";
	frmarr.submit();
}
</script>


<table width="750" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			디자이너:<% drawSelectBoxDesignerwithName "designer",designer  %>
			<br>
			<input type="radio" name="onoffgubun" value="on" <% if onoffgubun="on" then response.write "checked" %>>온라인상품
			<input type="checkbox" name="notupche" <% if notupche="on" then response.write "checked" %>>업체배송포함안함

			&nbsp;&nbsp;&nbsp;
			<input type="radio" name="onoffgubun" value="off" <% if onoffgubun="off" then response.write "checked" %>>오프라인상품

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
<table width="750" border="0" cellpadding="5" cellspacing="0" >
<tr>
	<td align="right">
	<a href="javascript:shopprint_pop('<%= designer %>','<%= notupche %>','<%= onoffgubun %>')";><b>[3102]print</b></a>
	&nbsp;&nbsp;
	<a href="javascript:shopprint_pop_3104('<%= designer %>','<%= notupche %>','<%= onoffgubun %>')";><b>[3104]print</b></a>
	</td>
</tr>
</table>

<table width="750" border="0" cellpadding="1" cellspacing="1" bgcolor=#3d3d3d class="a">
<tr bgcolor="#DDDDFF">
	<td width=20></td>
	<td>이미지</td>
	<td>브랜드</td>
	<td>코드</td>
	<td>상품명</td>
	<td></td>
	<td>가격</td>
<!--	<td>전시<br>여부</td>
	<td>판매<br>여부</td>	-->
	<td>사용<br>여부</td>
</tr>
<form name=frmarr method=post action="">
<input type=hidden name="designer" value="<%= designer %>">
<input type=hidden name="notupche" value="<%= notupche %>">
<input type=hidden name="onoffgubun" value="<%= onoffgubun %>">
<% if obarcode.Fresultcount<1 then %>
<tr height=30 bgcolor="#FFFFFF">
<td align="center" colspan=11> 검색 결과가 없습니다.</td>
</tr>
<% else %>
<% for i=0 to obarcode.Fresultcount-1 %>
<tr height=50 bgcolor="#FFFFFF">
	<td width=20><input type="checkbox" name="ckall" onClick="AnCheckClick(this)" value="<%= obarcode.FItemList(i).GetBarCode %>"></td>
	<td width="50" valign="top">
		<table width="50" height="50" border="0" cellpadding="1" cellspacing="0" bordercolor="#000000">
		<tr>
		  <% if obarcode.FItemList(i).FImageSmall<>"" then %>
		  <td><img src="<%= obarcode.FItemList(i).FImageSmall %>" width="50" height="50" border="1" bordercolor="gray"></td>
		  <% else %>
		  <td width="50" height="50" ><img src="http://image.10x10.co.kr/image/whiteblank.gif" width="50" height="50" border="1" bordercolor="gray"></td>
		  <% end if %>
		</tr>
		</table>
	</td>
	<td><%= obarcode.FItemList(i).Fmakerid %></td>
	<td><%= obarcode.FItemList(i).GetBarCode %></td>
	<td><%= obarcode.FItemList(i).FShopItemName %></td>
	<td>
		<% if obarcode.FItemList(i).FShopItemOptionName<>"" then %>
	 	<%= obarcode.FItemList(i).FShopItemOptionName %>
	 	<% end if %>
	</td>
	<td  align="right"><%= FormatNumber(obarcode.FItemList(i).Fshopitemprice,0) %></td>
	<td  align="center"><%= obarcode.FItemList(i).Fisusing %></td>
</tr>
<% next %>
<% end if %>
</form>
</table>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->