<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim i, j , orderno ,oordermaster, oorderdetail
	orderno = requestCheckVar(request("orderno"),16)

set oordermaster = new COrder
	oordermaster.frectorderno = orderno
	oordermaster.frectcancelyn = "N"
	oordermaster.fQuickSearchOrderMaster

set oorderdetail = new COrder
	oorderdetail.frectorderno = orderno
	oorderdetail.frectcancelyn = "N"
	oorderdetail.fQuickSearchOrderDetail
%>

<table width="701" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<% if oordermaster.ftotalcount>0 then %>
<tr>
	<td width="701" style="padding-top:15">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td style="border:1px solid #E1E1E1">
				<img src="http://fiximage.10x10.co.kr/web2007/cs_center/receipt_top.gif">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				<!-- 구매자 정보 시작 -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
				<tr>
				<td style="padding-top:15">
				<font color="#000000"><b>* 주문정보</b></font></td>
				</tr>
				<tr>
				<td>
					<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
					<tr>
						<td width="70" height="25" bgcolor="#F7F7F7" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1">
							<font color="#000000">주문번호</font>
						</td>
						<td width="70" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1">
							<%= oordermaster.FOneItem.Forderno %>
						</td>
						<td width="70" bgcolor="#F7F7F7" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1">
							<font color="#000000">주문일자</font>
						</td>
						<td width="70" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1">
							<%= left(oordermaster.FOneItem.FRegDate,10) %>
						</td>
					</tr>
					</table>
				</td>
				</tr>
				</table>
				<!-- 구매자 정보 끝 -->
			</td>
		</tr>
		<tr>
			<td>
				<!-- 구매상품 정보 시작 -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
				<tr>
					<td style="padding-top:15">
						<font color="#000000"><b>* 주문내역</b></font>
					</td>
				</tr>
				<tr>
					<td style="padding:2 0 0 0 ;border-top:1px solid #E1E1E1;border-bottom:2px solid #D4E4D1"  background="http://fiximage.10x10.co.kr/web2007/cs_center/top_bg.gif" height="30">
						<table width="100%" cellspacing="0" cellpadding="0" height="10" class="a">
						<tr>
							<td width="60" height="0" style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">상품코드</font></div></td>
							<td style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">상품명[옵션]</font></div></td>
							<td width="65" height="0" style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">판매가</font></div></td>
							<td width="30" height="0" style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">수량</font></div></td>
							<td width="60" height="0" style="padding:0 5 0 5"><div align="center"><font color="#000000">소계금액</font></div></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<% for i=0 to oorderdetail.FResultCount-1 %>

						<% if oorderdetail.FItemList(i).Fitemid <>0 then %>

							<table width="100%" style="border-bottom:1px solid #DCDCDC" cellspacing="0" cellpadding="0" class="a">
							<tr>
								<td width="60" align="center" valign="middle" style="padding:0 5 0 5">
									<%= oorderdetail.FItemList(i).fitemgubun%>-<%=CHKIIF(oorderdetail.FItemList(i).fitemid>=1000000,Format00(8,oorderdetail.FItemList(i).fitemid),Format00(6,oorderdetail.FItemList(i).fitemid))%>-<%=oorderdetail.FItemList(i).fitemoption %>
									<br>
									<% if oorderdetail.FItemList(i).Fisupchebeasong="N" then %>
										텐바이텐
									<% elseif oorderdetail.FItemList(i).Fisupchebeasong="Y" then %>
										<font color="red">업체개별</font>
									<% end if %>
								</td>
								<td align="left" valign="middle" style="padding:0 5 0 5">
									<%= oorderdetail.FItemList(i).FItemName %>
									<br>
									<font color="blue"><%= oorderdetail.FItemList(i).FItemoptionName %></font>
								</td>
								<td width="65" align="right" valign="middle" style="padding:0 5 0 5">
									<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
										<%= FormatNumber(oorderdetail.FItemList(i).fsellprice,0) %> 원
									<% else %>
										<font color="red">취소</font>
									<% end if %>
								</td>
								<td width="30" align="center" valign="middle" style="padding:0 5 0 5">
									<%= oorderdetail.FItemList(i).FItemNo %>
								</td>
								<td width="60" align="right" valign="middle" style="padding:0 5 0 5">
									<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
										<%= FormatNumber((oorderdetail.FItemList(i).fsellprice * oorderdetail.FItemList(i).FItemNo),0) %> 원
									<% else %>
										<font color="red">취소</font>
									<% end if %>
								</td>
							</tr>
							</table>
						<% end if %>

						<% next %>
					</td>
				</tr>
				</table>
				<!-- 구매상품 정보 끝 -->
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="left">
		<table width="700" border="0" align="left" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td height="80" align="left" valign="bottom" style="border-top:1px solid #dddddd"><img src="http://fiximage.10x10.co.kr/web2007/cs_center/receipt_bottom.gif"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" style="padding:15">
		<a href="javascript:window.print();"><img src="http://fiximage.10x10.co.kr/web2007/cs_center/print_btn.gif" border="0"></a>
	</td>
</tr>
<% else %>
<tr>
	<td align="center">
		주문내역이 없습니다
	</td>
</tr>
<% end if %>
</table>

<%
set oorderdetail = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->