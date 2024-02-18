<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 재고 일괄 등록 페이지
' History : 2008.04.24 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->

<%
dim evt_code
	evt_code = request("evt_code")

dim oip, i
	set oip = new Cauctionlist
	oip.frectevt_code = evt_code
	oip.feventitem_list()				
%>

<!-- #include virtual="/admin/auction/auction.js"-->

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="itemid">
<input type="hidden" name="mode">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			EventCode : <input type=text size=10 name="evt_code" value="<%= evt_code %>">
		</td>
	</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if oip.ftotalcount > 0 then %>
				<input type="button" class="button" value="저장" onclick="event_add(frm);">
			<% end if %>
		</td>
		<td align="right">
		</td>
	</tr>	
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.ftotalcount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17">
			검색결과 : <b><%= oip.ftotalcount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
   		<td align="center">Image</td>
		<td align="center">ItemCode</td>
		<td align="center">Maker</td>
		<td align="center">상품명</td>
		<td align="center">Option</td>
    </tr>
  
	<% for i=0 to oip.ftotalcount - 1 %>
		<form name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->
		<input type="hidden" name="mode">	
    	<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	   		<td align="center"><img src="<%= oip.flist(i).FImageSmall %>" width=40 height=40></td>
			<td align="center"><%= oip.flist(i).fitemid %>
			<input type="hidden" name="itemid" value="<%= oip.flist(i).fitemid %>">
			</td>
			<td align="center"><%= oip.flist(i).fmakerid %></td>
			<td align="center"><%= oip.flist(i).fitemname %></td>
			<td align="center"><%= oip.flist(i).fitemoptionname %></td>
    	</tr>   
		</form>
	<% next %>
	
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
	
</table>

<iframe frameboarder=0 height=0 width=0 name="view" id="view"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->