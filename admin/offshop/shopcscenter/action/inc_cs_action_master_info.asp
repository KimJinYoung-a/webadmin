<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>

<% if (IsDisplayCSMaster = true) then %>
<tr >
    <td >
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
		<tr>
			<td>
				<% getcurrstate_table currstate,divcd %>
			</td>
		</tr>
		</table>
		
		<br>    
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
                <%= oordermaster.FOneItem.forderno %><Br>(<%= oordermaster.FOneItem.fmasteridx %>)
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
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
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수일시</td>
            <td bgcolor="#FFFFFF">
                <% if (IsStatusRegister) then %>
                	<%= now() %>
                <% else %>
                	<%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
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
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수내용</td>
            <td bgcolor="#FFFFFF">
            	<textarea <% if IsStatusFinishing then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
            </td>
        </tr>
        <% if divcd = "A030" or divcd = "A031" then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">배송정보</td>
            <td bgcolor="#FFFFFF" colspan=3>
				<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr bgcolor="#FFFFFF">
				    <td width=100 bgcolor="<%= adminColor("topbar") %>">
				    	<font color="red"><%=deliverydivname%></font>
				    </td>
				    <td>
				    	<% if divcd = "A030" then %>
				    		<% DrawdeliveryCombo "deliverybox" ,deliverybox ," onchange='deliverych(this.value,searchfrm)'" ,"'1'" ,"Y" ,"frmaction" %>
				    	<% elseif divcd = "A031" then %>
				    		<%= ocsaslist.FOneItem.Fmakerid %>
				    	<% end if %>
				    </td>
				    <td width=100 bgcolor="<%= adminColor("topbar") %>">*수령인명</td>
				    <td><input type="text" class="text" name="reqname" value="<%= ocsaslist.FOneItem.FReqName %>"></td>				    
				</tr>
				<tr bgcolor="#FFFFFF">
				    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
				    <td><input type="text" class="text" name="reqphone" value="<%= ocsaslist.FOneItem.FReqPhone %>"></td>
				    <td bgcolor="<%= adminColor("topbar") %>">*핸드폰</td>
				    <td>
				    	<input type="text" class="text" name="reqhp" value="<%= ocsaslist.FOneItem.FReqHp %>">
				    	<input type="button" name="buyhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend_off('<%= ocsaslist.FOneItem.Freqhp %>','<%= ocsaslist.FOneItem.Fmasteridx %>','<%= oordermaster.FOneItem.forderno %>','');">
				    </td>				    
				</tr>
				<tr bgcolor="#FFFFFF">
				    <td valign="top" bgcolor="<%= adminColor("topbar") %>">*수령주소</td>
				    <td>
				        <input type="text" class="text" name="reqzipcode" value="<%= ocsaslist.FOneItem.FReqZipCode %>" size="7"  readonly><!-- id="[on,off,7,7][우편번호]" -->
				        <input type="button" class="button" value="검색" onClick="FnFindZipNew('frmaction','A')">
						<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmaction','A')">
				        <% '<input type="button" class="button" value="검색(구)" onClick="PopSearchZipcode('frmaction')"> %>
				        <Br><input type="text" class="text" name="reqzipaddr" id="[on,off,1,64][주소]" size="35" value="<%= ocsaslist.FOneItem.FReqZipAddr %>">
				        <Br><input type="text" class="text" name="reqaddress" id="[on,off,1,200][주소]" size="35" value="<%= ocsaslist.FOneItem.FReqAddress %>">
				    </td>
				    <td bgcolor="<%= adminColor("topbar") %>">기타사항</td>
				    <td>
				        <textarea class="textarea" rows="3" cols="35" name="comment" id="[off,off,off,off][기타사항]"><%= ocsaslist.FOneItem.FComment %></textarea>
					</td>				    
				</tr>
				<tr bgcolor="#FFFFFF">
				    <td bgcolor="<%= adminColor("topbar") %>">이메일</td>
				    <td colspan=3>
				    	<input type="text" class="text" name="reqemail" value="<%= ocsaslist.FOneItem.freqemail %>">
				    </td>
				</tr>
				</table>
            </td>
        </tr>
		<% end if %>
        <% if (IsStatusFinishing) or (IsUpcheConfirmState) or (IsStatusFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">처리내용</td>
            <td bgcolor="#FFFFFF" colspan=3>
	            <% if (IsUpcheConfirmState) then %>
	            	<textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
	            <% else %>
	            	<textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
	            <% end if %>
            </td>
            <!--<td bgcolor="<%= adminColor("pink") %>" align="center">처리관련<br>고객오픈<br>내용입력</td>
            <td bgcolor="#FFFFFF">
            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
            	<tr>
				    <td>
				    	<input class="text" type="text" name="opentitle" value="<%'= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr>
				    <td>
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%'= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>
				</table>
			</td>-->
        </tr>
        <% end if %>

        </table>
	</td>
</tr>
<% end if %>
