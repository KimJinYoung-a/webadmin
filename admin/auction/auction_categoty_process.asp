<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� ī�װ�
' History : 2008.03.04 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->
<%	
dim idx,idxsum ,itemid, rectitemid , category_gubun
	category_gubun = request("category_gubun")
	idxsum = request("idx")
	idx = left(idxsum,len(idxsum)-1)
%>	
<% 
dim sql

	sql = "update [db_item].dbo.tbl_auction set" 
	sql = sql & " auction_cate_code = '"& category_gubun &"'"
	sql = sql & " where ten_itemid in ("&idx&")"
	response.write sql&"<br>" & category_gubun
	dbget.execute sql

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script language="javascript">
opener.location.reload();
self.close();
</script>
