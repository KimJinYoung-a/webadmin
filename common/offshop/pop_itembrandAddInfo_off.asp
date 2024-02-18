<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2013.02.21 한용민 생성
' Description : 브랜드상품 추가
'				input - actionURL(db 처리에 필요한 파라미터까지 포함) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim i, makerid, shopid, acURL, brandcount
	makerid    = RequestCheckVar(request("makerid"),32)
	shopid    = RequestCheckVar(request("shopid"),32)
	acURL	= request("acURL")

if shopid = "" then
	response.write "<script>alert('매장ID 가 없습니다'); self.close();</script>"
	response.end
end if

if makerid<>"" then
	brandcount = getcontractbranditemcount(shopid,makerid)
end if
%>

<script language="javascript">

function jsSerach(){

	frm.target = "";
	frm.action = "";
	frm.submit();
}

function insertbranditem(){	
	
	if( confirm('해당 상품을 모두 추가 하시겠습니까?') ){
		frm.target = "FrameCKP";
		frm.action = "<%=acURL%>";
		frm.submit();

		opener.history.go(0);	
		//window.close();
	
	}else{
		return;	
	}	
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="mode" value="bi">
<input type="hidden" name="acURL" value="<%=acURL%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
	</td>
	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
	</td>
</tr>    
</table>

<br>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
<tr valign="bottom">       
    <td align="left">
    	※매장(<%=shopid%>)과 계약된 상품만 검색 됩니다.
    </td>
    <td align="right">
    </td>        
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="20">
		<% if brandcount<>"" then %>
			검색결과 : <b><%= brandcount %></b>개 상품이 검색 되었습니다.
			<% if brandcount <> 0 then %>
				<input type="button" value="모두추가(<%= brandcount %>건)" onClick="insertbranditem()" class="button">
			<% end if %>
		<% else %>
			브랜드를 입력해 주세요
		<% end if %>
	</td>		
</tr>
</table>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="800" height="100"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
