<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  수정및 입력
' History : 2007.09.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->

<%
dim idx, itemid , i
	idx = request("idx")

dim oip
	set oip = new Cauctionlist        			'클래스 지정
	oip.fauctionedit()		

%>
				
	<script language="javascript">
	function sendit(){
	if(document.frm.auction_cate_code.value==""){
	alert("옥션카테고리명을 입력하세요.")
	document.frm.auction_cate_code.focus();
	}
	else if(document.frm.auction_realsel.value==""){
	alert("옥션에 등록 하실 수량을 입력하세요.")
	document.frm.auction_realsel.focus();
	}
	else if(document.frm.auction_realsel.value==0){
	alert("옥션에 등록 하실 수량 1개 이상 입력하세요.")
	document.frm.auction_realsel.focus();
	}
	else if(document.frm.auction_isusing.value==""){
	alert("옥션에 등록 여부를 입력 하세요.")
	document.frm.auction_isusing.focus();
	}
	else
	document.frm.submit();
	}
	</script>
	
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" value="등록" onclick="sendit();" class="button">
		</td>
	</tr>
</form>	
</table>
<!-- 액션 끝 -->
	
<!--상품테이블시작-->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form method="get" name="frm" action="auctionedit_submit.asp">  	
	<tr bgcolor="#FFFFFF">
		<td rowspan=5><input type="hidden" name="mode"><img src="<%= oip.flist(0).FImageList %>" width="100" height="100"></td>
		<td><font size=2>페이지번호 :</font></td>
		<td><font size=2><%= idx %><input type="hidden" name="idx" value="<%= idx %>"></font></td>
		<td><font size=2>아이템 옵션 : </font></td>
		<td><font size=2><%= oip.flist(0).ten_option %></font>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>상품번호 :</font></td> 
		<td><font size=2><%= oip.flist(0).ten_itemid %><input type="hidden" name="ten_itemid" value="<%= oip.flist(0).ten_itemid %>"></font></td>
		<td><font size=2>옥션카테고리명 :</font></td> 
		<td><font size=2><input type="text" name="auction_cate_code" value="<%= oip.flist(0).auction_cate_code %>" size="13"></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>상품명 : </font></td>
		<td><font size=2><%= oip.flist(0).ten_itemname %></font></td>
		<td><font size=2>브랜드 :</font></td>
		<td><font size=2><%= oip.flist(0).ten_makerid %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>텐재고여부 : </font></td>
		<td>
			<font size=2>
			<% if oip.flist(0).ten_jaego >= 10 then
			response.write "Y"
			else
			response.write "N"
			end if %></font>
		</td>
		<td><font size=2>텐배이텐재고 : </font></td>
		<td><font size=2><%= oip.flist(0).ten_jaego %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>옥션등록여부 : </font></td>
		<td><font size=2><input type="text" name="auction_isusing" value="<%= oip.flist(0).auction_isusing %>" size="2"> ex)y,n</font></td>
		<td><!--<font size=2>옥션등록수량 : </font>--></td> 
		<td><font size=2><input type="hidden" name="auction_realsel" value="1" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=1>이미지경로 : </td>
		<td colspan=5><%= oip.flist(0).FImageList %></td>	
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=7>상품 상세 정보 : </td>	
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=7><%= nl2br(oip.flist(0).ten_itemcontent) %></td>
	</tr>
	</form>
</table>
<!--상품테이블끝-->
	
<!-- #include virtual="/lib/db/dbclose.asp" -->