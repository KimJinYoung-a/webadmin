<%@ language=vbscript %>
<% option explicit %>
<%
''///////// ���������� ���� �������� ����� ���� ������ �Դϴ�. ����� ��ġ����		'/�븸 ����
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader_xhtml.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim i, j
dim getPdfDownLinkUrlAdm
dim orderserial
orderserial = requestCheckvar(request("orderserial"),16)

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
addparam = "vos="&orderserial&"&ekey="&md5(orderserial)
if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnEstimateReceiptPdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnEstimateReceiptPdf.asp?"&addparam
end if

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
									<dt>Estimate :</dt>
									<dd>
										<% if orderserial="17032852294" then %>
											2017-03-24
										<% else %>
											<%= left(oordermaster.FOneItem.FRegDate,10) %>
										<% end if %>
									</dd>
								</dl>
								<dl>
									<dt>Corporate name :</dt>
									<dd>
										<% if orderserial="19010952388" then %>
											Chaeyeon Park
										<% else %>
											<%= oordermaster.FOneItem.FBuyName %>
										<% end if %>
									</dd>
								</dl>
								<dl>
									<dt>party speaker :</dt>
									<dd>
										<% if orderserial="19010952388" then %>
											Chaeyeon Park
										<% else %>
											<%= oordermaster.FOneItem.FBuyName %>
										<% end if %>
									</dd>
								</dl>
								<dl>
									<dt>HP :</dt>
									<dd><%= oordermaster.FOneItem.FBuyPhone %>/ <%= oordermaster.FOneItem.FBuyHp %></dd>
								</dl>
								<dl>
									<dt>�ѽ���ȣ :</dt>
									<dd></dd>
								</dl>
								<dl>
									<dt>EMAIL :</dt>
									<dd><%=oordermaster.FOneItem.Fbuyemail%></dd>
								</dl>
								<dl class="tMar20">
									<dt><strong>Order number :</strong></dt>
									<dd><strong><%= oordermaster.FOneItem.FOrderSerial %>
										 <% If oordermaster.FOneItem.IsForeignDeliver Then %>
									 		 (<strong>Overseas delivery</strong>)
								 		 <% End If %></strong></dd>
								</dl>
								<dl>
									<dt><strong>Total amount :</strong></dt>

									<% if orderserial="19010952388" then %>
										<dd><strong>94543 KRW</strong></dd>
									<% else %>
										<dd><strong>�� <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></strong></dd>
									<% end if %>
								</dl>
							</div>
							<div class="ftRt" style="width:55%;">
								<div class="title">
									<h2 class="ftLt" style="margin-top:0;">Provider Information</h2>
								</div>
								<table class="baseTable btmLine">
									<caption>Provider Information</caption>
									<colgroup>
										<col width="18%" /> <col width="30%" /> <col width="18%" /> <col width="30%" />
									</colgroup>
									<tbody>
									<tr>
										<th scope="row">Registration number</th>
										<td colspan="3">211-87-00620</td>
									</tr>
									<tr>
										<th scope="row">Mutual</th>
										<td>10x10</td>
										<th scope="row">CEO</th>
										<td>Choi Eun Hee</td>
									</tr>
									<tr>
										<th scope="row">Address</th>
										<td colspan="3">2F of the Free Building in Dongsung-dong, Jongno-gu, Seoul</td>
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
							<td>
								<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
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
                        		<% end if %>
							</td>
							<% if orderserial="16071803386" and oorderdetail.FItemList(i).Fitemid="1335555" then %>
								<td><%= oorderdetail.FItemList(i).FItemNo+6 %></td>
							<% else %>
								<td><%= oorderdetail.FItemList(i).FItemNo %></td>
							<% end if %>
							<td>
								<% if orderserial="16071803386" and oorderdetail.FItemList(i).Fitemid="1335555" then %>
									<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
										<%= FormatNumber(oorderdetail.FItemList(i).FItemCost*(oorderdetail.FItemList(i).FItemNo+6),0) %> <%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
										<% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
											<p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %></p>
										<% end if %>
									<% else %>
			                        	<font color="red">���</font>
			                        <% end if %>
								<% else %>
									<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
										<%= FormatNumber(oorderdetail.FItemList(i).FItemCost*oorderdetail.FItemList(i).FItemNo,0) %> <%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
										<% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
											<p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %></p>
										<% end if %>
									<% else %>
			                        	<font color="red">���</font>
			                        <% end if %>
								<% end if %>
							</td>
						</tr>
						<%
							else
								if orderserial="16071803386" and oorderdetail.FItemList(i).Fitemid="1335555" then
									if (oorderdetail.FItemList(i).Fcancelyn <> "Y") then
										totalBeasongPay = totalBeasongPay + oorderdetail.FItemList(i).FreducedPrice*(oorderdetail.FItemList(i).FItemNo+6)
									end if
								else
									if (oorderdetail.FItemList(i).Fcancelyn <> "Y") then
										totalBeasongPay = totalBeasongPay + oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo
									end if
								end if
							end if
						next
						%>
						</tbody>
						<tfoot>
						<tr>
							<td colspan="5">
								<div class="orderTotal">
									<% if orderserial="19010952388" then %>
										Total amount : <strong>94543 KRW</strong>
										<% if (totalBeasongPay <> 0) then %>
										(delivery charge : <%= FormatNumber(totalBeasongPay,0) %> KRW)
										<% end if %>
									<% else %>
										�� �հ�ݾ� : <strong><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></strong>��
										<% if (totalBeasongPay <> 0) then %>
										(��ۺ� : <%= FormatNumber(totalBeasongPay,0) %>��)
										<% end if %>
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
						<p><img src="http://fiximage.10x10.co.kr/web2013/my10x10/img_company_info2.gif" alt="�ٹ����� 10X10 / �Ǹ�ó �ȳ� : (��)�ٹ����� ����ڵ�Ϲ�ȣ : 211-87-00620 / ��ǥ�̻� : ������ / ������ : (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� / �ٹ����� �����;ȳ� TEL : 1644-6030 / AM 09 :00~PM 06:00 ���ɽð� PM 12:00~01:00 �ָ�,������ �޹� / E-mail : customer@10x10.co.kr " /></p>
					</div>

					<div class="btnArea tMar30 ct">
						<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">�μ��ϱ�</button>
						<!--<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF ��ȯ</button>-->
					</div>
				</div>
				<!-- //content -->
			</div>
		</div>
		<div class="popFooter">
			<div class="btnArea">
				<button type="button" class="btn btnS1 btnGry2" onclick="window.close();">�ݱ�</button>
			</div>
		</div>
	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
