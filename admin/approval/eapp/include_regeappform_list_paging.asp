<%

iPerCnt = 10

iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrPage mod iPerCnt) = 0 Then
	iEndPage = iCurrPage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
	<tr height="25" >
		<td colspan="15" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
			    <tr valign="bottom" height="25">
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrPage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			    </tr>
			</table>
		</td>
	</tr>