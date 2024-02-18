<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  옥션 재고 등록
' History : 2007.09.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->

<!--수정모드시작-->
<% 
dim idx,auction_cate_code,auction_realsel,auction_isusing,ten_jaego_isusing,ten_jaego,itemid
dim i,sql111,oip1
	idx = request("idx") 				'인덱스값을 받아온다.
	auction_cate_code = request("auction_cate_code") 				'카테고리명
	auction_realsel = request("auction_realsel") 				'옥션 등록수량
	auction_isusing = request("auction_isusing") 				'옥션등록
	ten_jaego = request("ten_jaego")
	itemid = request("ten_itemid")
	
	if ten_jaego > 10 then
		ten_jaego_isusing = "y"
	else
		ten_jaego_isusing = "n"	
	end if
		
%>

<%
	dim sql50,sql51,ten_auction_option_rect,ten_auction_option,ten_auction_cnt_rect,ten_auction_cnt
		sql50 = "update [db_item].[dbo].tbl_auction set auction_cate_code = "& auction_cate_code &" ,auction_realsel = "& auction_realsel &", auction_isusing = '"& auction_isusing &"'"	& VbCrlf
		sql50 = sql50 & " where idx = " & idx 
		response.write sql50
		dbget.execute sql50
%>	
<!--수정모드끝-->
			
<!-- #include virtual="/lib/db/dbclose.asp" -->

	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>


