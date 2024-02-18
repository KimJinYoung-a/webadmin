<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/incShoplinkerFunction.asp"-->
<%
Dim i, itemid, divid, tmp, strSQL, mallname, mallprdno, malluserid
itemid = request("subcmd")
divid = request("cmdparam")
If divid = "" or itemid = "" Then Response.end

strSQL = ""
strSQL = strSQL & " select D.mall_name, D.mall_product_id, D.mall_user_id  "
strSQL = strSQL & " From db_item.dbo.tbl_Shoplinker_OutmallControl AS C "
strSQL = strSQL & " join db_item.dbo.tbl_Shoplinker_Outmall as D on C.mall_user_id = D.mall_user_id and C.mall_name = D.mall_name "
strSQL = strSQL & " where D.itemid = '"&itemid&"'  "
strSQL = strSQL & " ORDER BY D.mall_name DESC, D.mall_product_id DESC "
rsget.Open strSQL,dbget,1
if Not(rsget.EOF or rsget.BOF) then
	i=0
	Do until rsget.EOF
		mallname = rsget("mall_name")
		mallprdno = rsget("mall_product_id")
		malluserid = "[ " &rsget("mall_user_id")& " ]"
		Select Case mallname
			Case "GS홈쇼핑(eshop)"		mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://www.gsshop.com/prd/prd.gs?prdid="&mallprdno&"','vies',''); >"
			Case "(주)위즈위드"			mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://www.wizwid.com/CSW/handler/wizwid/kr/MallProduct-Start?AssortID="&mallprdno&"','vies',''); >"
			Case "엔조이뉴욕(미러스)"	mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://www.njoyny.com/shop/goods_view.jsp?pid="&mallprdno&"','vies',''); >"
			Case "디앤샵"				mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://www.dnshop.com/front/product/ProductDetail?PID="&mallprdno&"','vies',''); >"
			Case "가방팝"				mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://bizest.gabangpop.co.kr/app/product/detail/"&Split(mallprdno,"-")(0)&"/"&Split(mallprdno,"-")(1)&"','vies',''); >"
			Case "PLAYER"				mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://www.player.co.kr/v3/category/detail.php?goods_no="&Split(mallprdno,"-")(0)&"&goods_sub="&Split(mallprdno,"-")(1)&"','vies',''); >"
			Case "HOTTRACKS"			mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://www.hottracks.co.kr/ht/product/detail?barcode="&mallprdno&"','vies',''); >"
			Case "신세계닷컴2.0"		mallname = "<input type='button' class='button' style='cursor:hand' value='"&mallname&"' onClick=javascript:window.open('http://mall.shinsegae.com/item/item.do?method=viewItemDetail&item_id="&mallprdno&"','vies',''); >"
		End Select
		tmp = tmp & mallname & malluserid & "<Br>"
		rsget.MoveNext
		i=i+1
	Loop
end if
rsget.Close
%>
<script language="">
	var divid = '<%=divid%>';
	parent.eval("document.all."+divid).innerHTML = "<%=tmp%>";
</script>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
