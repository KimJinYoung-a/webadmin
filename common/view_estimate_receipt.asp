<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader_xhtml.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim i, j
dim getPdfDownLinkUrlAdm
dim orderserial, ekey,webImgUrl, webImgSSLUrl
orderserial = requestCheckvar(request("vos"),16)
ekey =  requestCheckvar(request("ekey"),32)

if (ekey="") then
    response.write "��ȣȭ Ű�� �ùٸ��� �ʽ��ϴ�.1"
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5(orderserial))) then
    response.write "��ȣȭ Ű�� �ùٸ��� �ʽ��ϴ�.2"
    response.end
end if

IF application("Svr_Info")="Dev" THEN
	webImgUrl		= "http://testwebimage.10x10.co.kr"
	webImgSSLUrl	= "http://testwebimage.10x10.co.kr"
ELSE
	webImgUrl		= "http://webimage.10x10.co.kr"
	webImgSSLUrl	= "http://webimage.10x10.co.kr"
END IF

dim oordermaster, oorderdetail
dim addparam

set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial
oordermaster.QuickSearchOrderMaster

if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if

set oorderdetail = new COrderMaster
oorderdetail.FRectOldOrder = oordermaster.FRectOldOrder
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

dim totalBeasongPay : totalBeasongPay = 0

%>
<style>
		#orderPrint .baseTable th, #orderPrint .baseTable td {padding:4px 0;}
		.companyInfo {margin:40px 0 0 0;}
		.popWrap .popHeader {background:transparent; text-align:center; font-size:20px;}
		h1 {font-size:24px; font-weight:bold;}
		h1 img {display:none;}
		h1 span {display:block;}
		.fs14 {font-size:12px;}
		.estimateTo {font-size:11px;}
		.estimateTo dl {overflow:hidden; _zoom:1; margin-top:3px;}
		.estimateTo dl.tMar20 {margin-top:10px;}
		.estimateTo dl dt, .estimateTo dl dd {float:left;}
		.estimateTo dl dt {width:70px;}
		#orderPrint h2 {font-size:12px; font-weight:bold; margin:20px 0 5px 0;}
	</style>
<script type="text/javascript">
	function jsGoPDF(iUri){
		  var popwin = window.open(iUri,'dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes');
	}
</script>
	<div class="heightgird" id="orderPrint"><!-- 2013.09.24 : id="mediaPrint" �߰� -->
		<div class="popWrap">
			<div class="popHeader">
				<h1><span>�� �� ��</span></h1>
			</div>
			<div class="popContent">
				<!-- content -->
				<div class="mySection">
					<div class="orderDetail">
						<div class="overHidden">
							<div class="ftLt estimateTo" style="width:45%;">
								<dl>
									<dt>�� �� �� :</dt>
									<dd><%= left(oordermaster.FOneItem.FRegDate,10) %></dd>
								</dl>
								<dl>
									<dt>ȸ �� �� :</dt>
									<dd><%= oordermaster.FOneItem.FBuyName %></dd>
								</dl>
								<dl>
									<dt>�� �� �� :</dt>
									<dd><%= oordermaster.FOneItem.FBuyName %></dd>
								</dl>
								<dl>
									<dt>��ȭ��ȣ :</dt>
									<dd><%= oordermaster.FOneItem.FBuyPhone %>/ <%= oordermaster.FOneItem.FBuyHp %></dd>
								</dl>
								<dl>
									<dt>�ѽ���ȣ :</dt>
									<dd></dd>
								</dl>
								<dl>
									<dt>�� �� �� :</dt>
									<dd><%=oordermaster.FOneItem.Fbuyemail%></dd>
								</dl>
								<dl class="tMar20">
									<dt><strong>�ֹ���ȣ :</strong></dt>
									<dd><strong><%= oordermaster.FOneItem.FOrderSerial %>
										 <% If oordermaster.FOneItem.IsForeignDeliver Then %>
									 		 (<strong>�ؿܹ��</strong>)
								 		 <% End If %></strong></dd>
								</dl>
								<dl>
									<dt><strong>�հ�ݾ� :</strong></dt>
									<dd><strong>�� <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></strong></dd>
								</dl>
							</div>
							<div class="ftRt" style="width:55%;">
								<div class="title">
									<h2 class="ftLt" style="margin-top:0;">������ ����</h2>
								</div>
								<table class="baseTable btmLine">
									<caption>������ ����</caption>
									<colgroup>
										<col width="18%" /> <col width="30%" /> <col width="18%" /> <col width="30%" />
									</colgroup>
									<tbody>
									<tr>
										<th scope="row">��Ϲ�ȣ</th>
										<td colspan="3">211-87-00620</td>
									</tr>
									<tr>
										<th scope="row">��ȣ</th>
										<td>(��)�ٹ�����</td>
										<th scope="row">��ǥ��</th>
										<td>������</td>
									</tr>
									<tr>
										<th scope="row">������</th>
										<td colspan="3">(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����</td>
									</tr>
									<tr>
										<th scope="row">����</th>
										<td>����, ���Ҹ� ��</td>
										<th scope="row">����</th>
										<td>���ڻ�ŷ� ��</td>
									</tr>
									<tr>
										<th scope="row">�����</th>
										<td></td>
										<th scope="row">����ó</th>
										<td> </td>
									</tr>
									</tbody>
								</table>
							</div>
						</div>

						<div class="title">
							<h2>�ֹ���ǰ����</h2>
						</div>
						<table class="baseTable btmLine">
						<caption>�ֹ���ǰ���� ���</caption>
						<colgroup>
							<col width="90" /> <col width="*" /> <col width="90" /> <col width="60" /> <col width="100" />
						</colgroup>
						<thead>
						<tr>
							<th scope="col">��ǰ�ڵ�</th>
							<th scope="col">��ǰ��[�ɼǸ�]</th>
							<th scope="col">�ܰ�</th>
							<th scope="col">����</th>
							<th scope="col">�ݾ�</th>
						</tr>
						</thead>
						<tbody>
					 <% for i=0 to oorderdetail.FResultCount-1 %>
						 <% if oorderdetail.FItemList(i).Fitemid <>0 then %>
						<tr>
							<td><%= oorderdetail.FItemList(i).FItemid %></td>
							<td class="lt">
								<div><%= oorderdetail.FItemList(i).FItemName %></div>
								<div><font color="blue"><%= oorderdetail.FItemList(i).FItemoptionName %></font></div>
							</td>
							<td><% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
							<% if (oorderdetail.FItemList(i).IsSaleItem) then %>
                                    <strike><%= FormatNumber(oorderdetail.FItemList(i).Forgitemcost,0) %></strike><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %><br>
                                    <strong class="crRed"><%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %></strong><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
                                <% else %>
                                    <% if (oorderdetail.FItemList(i).IsItemCouponAssignedItem) then %>
                                    <strike><%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %></strike><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
                                    <% else %>
                                    <%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
                                    <% end if %>
                                <% end if %>

                                <% if (oorderdetail.FItemList(i).IsItemCouponAssignedItem) then %>
                                    <br><strong class="crGrn"><%= FormatNumber(oorderdetail.FItemList(i).FItemCost,0) %>��</strong>
                                <% else %>

                                <% end if %>

                                <% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
                                <p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %></p>
                                <% end if %>
                        <% else %>
                        	<font color="red">���</font>
                        <% end if %></td>
							<td><%= oorderdetail.FItemList(i).FItemNo %></td>
							<td>
								<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
									<%= FormatNumber(oorderdetail.FItemList(i).FItemCost*oorderdetail.FItemList(i).FItemNo,0) %> <%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
									<% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
									<p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %></p>
									<% end if %>
								 <% else %>
		                        	<font color="red">���</font>
		                        <% end if %>
                       	 </td>
						</tr>
						<%
							else
								if (oorderdetail.FItemList(i).Fcancelyn <> "Y") then
									totalBeasongPay = totalBeasongPay + oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo
								end if
							end if
						next
						%>
						</tbody>
						<tfoot>
						<tr>
							<td colspan="5">
								<div class="orderTotal">
									�� �հ�ݾ� : <strong><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></strong>��
									<% if (totalBeasongPay <> 0) then %>
									(��ۺ� : <%= FormatNumber(totalBeasongPay,0) %>��)
									<% end if %>
								</div>
							</td>
						</tr>
						</tfoot>
						</table>
					</div>

					<dl class="tMar30">
						<dt class="bBdr2 bPad05 fs14"><strong>�������</strong></dt>
						<dd class="tPad05 tBdr3">1. �հ�ݾ��� �ΰ��� ���� �ݾ��Դϴ�.</dd>
						<dd>2. �� ������ ��ۺ� �����մϴ�.</dd>
						<dd>3. ������ȿ�Ⱓ�� �����Ϸκ��� 15���Դϴ�.</dd>
					</dl>

					<div class="companyInfo">
						<p><img src="http://fiximage.10x10.co.kr/web2020/my10x10/img_company_info.png" alt="company info" /></p>
					</div>
				</div>
				<!-- //content -->
			</div>
		</div>
	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
