<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.09 한용민 생성
'###########################################################
%>
<% if (IsDisplayCSMaster = true) then %>
<tr >
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">접수구분</td>
            <td bgcolor="#FFFFFF">
		    	<font style='line-height:100%; font-size:15px; color:blue; font-family:돋움; font-weight:bold'><%= GetCSCommName_off("Z001", divcd) %></font>
		    	&nbsp;
                <% if (Not IsStatusRegister) then %>
			    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:돋움; font-weight:bold'>[<%= ocsaslist.FOneItem.shopGetCurrstateName %>]</font>
			    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
						<font style='line-height:100%; font-size:15px; color:#FF0000; font-family:돋움; font-weight:bold'>- 삭제된 내역</font>
			    	<% end if %>
		    	<% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">주문번호<br>(일렬번호)</td>
            <td bgcolor="#FFFFFF" width="200" >
                <%= orderno %><Br>(<%= oordermaster.FOneItem.fmasteridx %>)
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= oordermaster.FOneItem.shopIpkumDivColor %>"><%= oordermaster.FOneItem.shopIpkumDivName %></font>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수자</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                    <%= session("ssbctid") %>
                <% else %>
                    <%= ocsaslist.FOneItem.Fwriteuser %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자</td>
            <td bgcolor="#FFFFFF">
               <%= oordermaster.FOneItem.Fbuyname %>
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수일시</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                	<%= now() %>
                <% else %>
                	<%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자정보</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FBuyname %>                 
                 [<%= oordermaster.FOneItem.FBuyHp %>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수제목</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= GetDefaultTitle_off(divcd,"", masteridx, orderno) %>" size="56" maxlength="56">
                <% else %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">수령인정보</td>
            <td bgcolor="#FFFFFF">
                 <%= oordermaster.FOneItem.FReqName %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FReqHp %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">접수내용</td>
            <td bgcolor="#FFFFFF" rowspan="2">
            	<textarea <% if IsStatusFinishing then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">배송지정보</td>
            <td bgcolor="#FFFFFF" valign="top">
            	[<%= oordermaster.FOneItem.FReqZipCode %>]<br>
                <%= oordermaster.FOneItem.FReqZipAddr %><br>
                <%= oordermaster.FOneItem.FReqAddress %>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">관련택배정보</td>
            <td bgcolor="#FFFFFF" valign="top">
            	<!-- 코딩 확인할것 -->
            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
			        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
			        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
			        <% dim ifindurl : ifindurl = fnGetSongjangURL(ocsaslist.FOneItem.Fsongjangdiv) %>
			        <% if (ocsaslist.FOneItem.Fsongjangdiv="24") then %>
                		<a href="javascript:popDeliveryTrace('<%= ifindurl %>','<%= ocsaslist.FOneItem.Fsongjangno %>');">추적</a>
                	<% else %>
			            <a href="<%= ifindurl + ocsaslist.FOneItem.Fsongjangno %>" target="_blank">추적</a>
			        <% end if %>
			        <input type="button" class="button" value="수정" onClick="changeSongjang('<%= csmasteridx %>');">
		        <% end if %>
            </td>
        </tr>
        <% if (IsStatusFinishing) or (IsUpcheConfirmState) or (IsStatusFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">처리내용</td>
            <td bgcolor="#FFFFFF">
            <% if (IsUpcheConfirmState) then %>
            	<textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% else %>
            	<textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% end if %>
            </td>
            <td bgcolor="<%= adminColor("pink") %>" align="center">처리관련<br>고객오픈<br>내용입력</td>
            <td bgcolor="#FFFFFF">
            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
            	<tr>
				    <td>
				    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr>
				    <td>
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>
				</table>
			</td>
        </tr>
        <% end if %>
		<% 
		'/취소시
		if divcd = "A008" then
			'/출고완료가 아닐경우
			if OrderMasterState <> "8" then
		%>        
        <!--<tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">취소주문번호<Br>(마이너스주문)</td>
            <td bgcolor="#FFFFFF" colspan="3">
				<input type="text" size=20 name="cancelorderno" value="<%'= oordermaster.FOneItem.fcancelorgorderno %>">
            </td>
        </tr>-->
		<%
			end if
		end if
		%>
        </table>
	</td>
</tr>
<% end if %>
