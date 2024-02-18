<tr >
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">접수상품</td>
            <td colspan="3" bgcolor="#FFFFFF">
		        <table height="25" width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#BABABA">
		            <tr align="center" bgcolor="<%= adminColor("topbar") %>">
                    	<td width="30" height="40">구분</td>
                    	<td width="50">진행상태</td>
                    	<td width="40">CODE</td>
                      	<td width="50">이미지</td>
                    	<td>Gift카드명<br><font color="blue">[옵션]</font></td>
                    	<td width="60">판매가</td>

                    	<td width="100">예약발송일</td>
                    	<td width="100">전송일</td>
                    	<td width="100">등록일</td>
                    	<td width="100">취소일</td>
                    </tr>
                    <tr>
                        <td height="1" colspan="10" bgcolor="#BABABA"></td>
                    </tr>
		            <tr align="center" bgcolor="<%= adminColor("topbar") %>">
                    	<td height="60"></td>
                    	<td><%= ogiftcardordermaster.FOneItem.GetCardStatusName %></td>
                    	<td><%= ogiftcardordermaster.FOneItem.FcardItemid %></td>
                    	<td><img src="<%= ogiftcardordermaster.FOneItem.FSmallimage %>"></td>
                    	<td>
                    		<%= ogiftcardordermaster.FOneItem.FCarditemname %><br><font color="blue">[<%= ogiftcardordermaster.FOneItem.FcardOptionName %>]</font>
                    	</td>
                    	<td><%= FormatNumber(ogiftcardordermaster.FOneItem.Fsubtotalprice, 0) %></td>
                    	<td><%= ogiftcardordermaster.FOneItem.FbookingDate %></td>
                    	<td><%= Left(ogiftcardordermaster.FOneItem.FsendDate, 10) %></td>
                    	<td><%= ogiftcardordermaster.FOneItem.FcardregDate %></td>
                    	<td><%= ogiftcardordermaster.FOneItem.Fcanceldate %></td>
                    </tr>
                 </table>
            </td>
		</tr>
		</table>
	</td>
</tr>
