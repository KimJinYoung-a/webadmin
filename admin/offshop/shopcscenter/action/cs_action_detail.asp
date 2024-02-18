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
dim i ,masteridx, divcd, currstate ,ocsaslist ,orefund ,oordermaster ,OCsDetail
dim OCsDelivery  ,OCsConfirm ,deliverydivname ,reqname ,reqphone ,reqhp ,zipcode ,reqzipcode ,reqzipaddr ,reqetcaddr
	masteridx = requestCheckVar(request("masteridx"),10)

set ocsaslist = New COrder
	ocsaslist.FRectCsAsID = masteridx

	if (masteridx<>"") then
	    ocsaslist.fGetOneCSASMaster()
	end if

set oordermaster = new COrder

	if (ocsaslist.FResultCount > 0) then
	    oordermaster.FRectmasteridx = ocsaslist.FOneItem.fmasteridx
	    oordermaster.fQuickSearchOrderMaster

	    divcd = ocsaslist.FOneItem.FDivCD
	    currstate = ocsaslist.FOneItem.Fcurrstate
	end if

set OCsDetail = new COrder
	OCsDetail.FRectCsAsID = masteridx

	if ocsaslist.FResultCount>0 then
	    OCsDetail.fGetCsDetailList
	end if

set OCsDelivery = new COrder
	OCsDelivery.FRectCsAsID = masteridx

	if ocsaslist.FResultCount>0 then
	    OCsDelivery.fGetOneCsDeliveryItem

	    if OCsDelivery.Ftotalcount>0 then
		    reqname = OCsDelivery.FOneItem.Freqname
		    reqphone = OCsDelivery.FOneItem.Freqphone
		    reqhp = OCsDelivery.FOneItem.Freqhp
		    zipcode = OCsDelivery.FOneItem.Freqzipcode
		    reqzipaddr = OCsDelivery.FOneItem.Freqzipaddr
		    reqetcaddr = OCsDelivery.FOneItem.Freqetcaddr
		end if
	end if

if divcd = "A030" then
	deliverydivname = "A/S완료후수령지"
elseif divcd = "A031" then
	deliverydivname = "A/S업체"
end if
%>

<script type='text/javascript'>

function PopCSMailTest(masteridx){
    var popwin = window.open('/admin/offshop/cscenter/action/cs_action_mail_view.asp?masteridx=' + masteridx,'cs_action_mail_view','width=600,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function CardCancelProcess(iid){
    var popwin = window.open('pop_CardCancel.asp?id=' + iid,'PopCardCancelProcess','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function regConfirmMsg(iid,fin){
    var popwin = window.open('pop_ConfirmMsg.asp?id=' + iid + '&fin=' + fin,'regConfirmMsg','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopCSAddUpchejungsanEdit(iid){
    var popwin = window.open('pop_AddUpchejungsanEdit.asp?id=' + iid ,'AddUpchejungsanEdit','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
<form name="frmdetail" onsubmit="return false;">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<% if ocsaslist.FResultCount>0 then %>
<tr>
	<td>
		<% getcurrstate_table ocsaslist.FOneItem.Fcurrstate,divcd %>
	</td>
</tr>
</table>

<br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr valign="top" height="450">
	<td>
		<!-- 접수 정보 -->
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
	        <td colspan="4">
	            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
	            	<tr>
		    		    <td>
		    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>접수 정보</b>
		    		    	&nbsp;[<%= ocsaslist.FOneItem.shopGetAsDivCDName %>]
		    		    	&nbsp;<%= ocsaslist.FOneItem.Forderno %>
		    		    </td>
		    		    <td align="right" >
		    		        <input class="button" type="button" value="정보수정" onclick="javascript:PopCSActionEdit('','editreginfo','<%= masteridx %>');" >
		    		    </td>
		    		 </tr>
    		    </table>
		    </td>
		</tr>
		<tr>
		    <td width="100" bgcolor="<%= adminColor("topbar") %>">접수자</td>
		    <td width="150" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fwriteuser %></td>
		    <td width="100" bgcolor="<%= adminColor("topbar") %>">접수일시</td>
		    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fregdate %></td>
		</tr>
		<tr height="20">
		    <td bgcolor="<%= adminColor("topbar") %>">제목</td>
		    <td colspan="3" bgcolor="#F4F4F4"><input type="text" class="text_ro" name="title" value="<%= ocsaslist.FOneItem.FTitle %>" size="68" maxlength="60" ReadOnly></td>
		</tr>
		<tr bgcolor="#F4F4F4">
		    <td bgcolor="<%= adminColor("topbar") %>">접수내용</td>
		    <td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea_ro" name="contents_jupsu" cols="68" rows="8" ReadOnly><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea></td>
		</tr>
		</table>
		<!-- 접수 정보 -->
		<br>
		<table width="100%" border=0 align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="FFFFFF">
			<td valign="top" width="50%">

				<!-- 접수시 주소정보 시작-->
				<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>배송주소 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		        <% if (divcd="A030") or (divcd="A031") or (OCsDelivery.FResultCount>0) then %>
					    		        <% if (currstate="B001") then %>
					    		        	<input class="button" type="button" value="주소변경" onclick="popEditCsDelivery('<%= masteridx %>');" >
										<% elseif C_ADMIN_AUTH or C_OFF_AUTH then %>
											<input class="button" type="button" value="주소변경(관리자모드)" onclick="popEditCsDelivery('<%= masteridx %>');" >
					    		        <% else %>
					    		        	<input class="button" type="button" value="주소변경불가" onclick="alert('접수상태에서만 변경가능 합니다.');" >
					    		        <% end if %>
				    		        <% end if %>
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<% if divcd = "A030" or divcd = "A031" then %>
				<!-- 배송정보 -->
				<tr>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("pink") %>"><%= deliverydivname %></td>
				    <td width="50" bgcolor="<%= adminColor("pink") %>">고객명</td>
				    <td width="100" bgcolor="#FFFFFF"><%= reqname %></td>
				    <td width="50" bgcolor="<%= adminColor("pink") %>">연락처</td>
				    <td bgcolor="#FFFFFF"><%= reqphone %> / <%= reqhp %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("pink") %>">주소</td>
				    <td colspan="3" bgcolor="#FFFFFF">[<%= reqzipcode %>] <%= reqzipaddr %> &nbsp;<%= reqetcaddr %></td>
				</tr>
				<% end if %>
				</table>
			</td>
			<td valign="top">
				<!-- 처리 정보 시작-->
				<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>처리 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		    	<% if currstate <> "B007" then %>
				    		        	<input class="button" type="button" value="최종완료처리" onclick="PopCSActionFinish('','finishreginfo','<%= masteridx %>');" >
				    		        <% end if %>
				    		    </td>
				    		 </tr>
		    			</table>
		    		</td>
		    	</tr>
				<tr>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">처리자</td>
				    <td width="80" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Ffinishuser %><% if isnull(ocsaslist.FOneItem.Ffinishuser) then %>미처리<% end if %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">처리일시</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Ffinishdate %><% if isnull(ocsaslist.FOneItem.Ffinishuser) then %>미처리<% end if %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">관련송장</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
					        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
					        <a href="<%= DeliverDivTrace(Trim(ocsaslist.FOneItem.Fsongjangdiv)) %><%= ocsaslist.FOneItem.Fsongjangno %>" target="_blank">추적</a>

					        <% if currstate <> "B007" then %>
				            	<input type="button" class="button" value="수정" onClick="changeSongjang('<%= masteridx %>');">
				            <% elseif C_ADMIN_AUTH or C_OFF_AUTH then %>
				            	<input class="button" type="button" value="수정(관리자모드)" onClick="changeSongjang('<%= masteridx %>');">
							<% end if %>
				        <% end if %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">처리내용</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				        <textarea class="textarea_ro" name="contents_finish" cols="48" rows="8"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
				    </td>
				</tr>
				<!--<tr bgcolor="<%'= adminColor("pink") %>">
				    <td rowspan="2">처리관련<br>고객오픈<br>내용입력</td>
				    <td colspan="3">
				    	<input class="text" type="text" name="opentitle" value="<%'= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr bgcolor="<%'= adminColor("pink") %>">
				    <td colspan="3">
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%'= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>-->
				</table>
				<!-- 처리 정보 끝-->
			</td>
		</tr>
		</table>

	</td>
</tr>
<tr>
	<td valign="top">
		<!-- 접수시 상품정보 시작-->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
	        <td>
	            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
	            	<tr>
		    		    <td>
		    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품 정보</b> (접수 된 내역)
		    		    </td>
		    		    <td align="right" >
		    		    </td>
		    		 </tr>
    		    </table>
		    </td>
		</tr>
		<tr valign="top" bgcolor="<%= adminColor("topbar") %>">
		   	<td>
		   		<table width="100%" height="200" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
	            	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
	            		<td style="width:30px; border-right:1px solid <%= adminColor("tablebg") %>;">구분</td>
		    		    <td style="width:40px; border-right:1px solid <%= adminColor("tablebg") %>;">CODE</td>
		    		    <td style="border-right:1px solid <%= adminColor("tablebg") %>;">상품명[옵션]</td>
		    		    <td style="width:50px; border-right:1px solid <%= adminColor("tablebg") %>;">판매가</td>
		    		    <td style="width:30px;">수량</td>
		    		</tr>
		    		<tr>
                        <td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
                    </tr>
                    <% for i=0 to OCsDetail.FResultCount-1 %>
                    <tr height="25" align="center" bgcolor="#FFFFFF" >
		    			<td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"></td>
		    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;">
		    		    	<%=OCsDetail.FItemList(i).fitemgubun%>-<%=CHKIIF(OCsDetail.FItemList(i).fitemid>=1000000,Format00(8,OCsDetail.FItemList(i).fitemid),Format00(6,OCsDetail.FItemList(i).fitemid))%>-<%=OCsDetail.FItemList(i).fitemoption%>
		    		    </td>
		    		    <td align="left" style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemname %>[<%= OCsDetail.FItemList(i).Fitemoptionname %>]</td>
		    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).fsellprice %></td>
		    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fregitemno %></td>
		    		</tr>
                    <% next %>
                    <tr bgcolor="#FFFFFF">
                        <td colspan="6"></td>
                    </tr>
    		    </table>
		   	</td>
		</tr>
		</table>
		<!-- 접수시 주문정보 끝-->

	</td>
</tr>
<% else %>
	<tr height="50" colspan=20>
	    <td align="center">[ 선택된 처리AS 가 없습니다. 먼저 처리 내역을 선택하세요 ]</td>
	</tr>
<% end if %>
</form>
<table>

<%
set ocsaslist   = Nothing
set oordermaster = Nothing
set OCsDetail = Nothing
set OCsDelivery = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->