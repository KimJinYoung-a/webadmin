<%
'###########################################################
' Description : 바코드 출력 프린트 설정 매뉴
' Hieditor : 2016.12.15 한용민 생성
'###########################################################
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		※ 프린터 설정
		<br>
		* 용지규격 :
		<select name="printername" onchange="reg('');">
			<option value="FORMTEC_A4" <% if printername = "FORMTEC_A4" then response.write " selected" %>>폼텍 (규격A4)</option>
			<option value="TEC_B-FV4_45x22" <% if printername = "TEC_B-FV4_45x22" then response.write " selected" %>>TEC B-FV4 (규격45x22)</option>
			<option value="TEC_B-FV4_35x15" <% if printername = "TEC_B-FV4_35x15" then response.write " selected" %>>TEC B-FV4 (규격35x15)</option>

			<% if onoffgubun="OFFLINE" then %>
				<option value="TEC_B-FV4_45x45" <% if printername = "TEC_B-FV4_45x45" then response.write " selected" %>>TEC B-FV4 (규격45x45)</option>
			<% end if %>

			<option value="TEC_B-FV4_80x50" <% if printername = "TEC_B-FV4_80x50" then response.write " selected" %>>TEC B-FV4 (규격80x50)</option>
			<option value="TTP-243_45x22" <% if printername = "TTP-243_45x22" then response.write " selected" %>>TTP-243 (규격45x22)</option>
			<option value="TTP-243_35x15" <% if printername = "TTP-243_35x15" then response.write " selected" %>>TTP-243 (규격35x15)</option>

			<% if onoffgubun="OFFLINE" then %>
				<option value="TTP-243_45x45" <% if printername = "TTP-243_45x45" then response.write " selected" %>>TTP-243 (규격45x45)</option>
			<% end if %>

			<option value="TTP-243_80x50" <% if printername = "TTP-243_80x50" then response.write " selected" %>>TTP-243 (규격80x50)</option>
		</select>
		&nbsp;
		* 표시상품명 :
		<select name="isforeignprint" onchange="reg('');">
			<option value="N" <% if (isforeignprint = "N") then %>selected<% end if %>>국내상품명</option>
			<option value="Y" <% if (isforeignprint = "Y") then %>selected<% end if %>>해외상품명</option>
		</select>
		&nbsp;
		* 금액표시방식 :
		<select name="printpriceyn" onchange="reg('');">
			<option value="Y" <% if (printpriceyn = "Y") then %>selected<% end if %>>소비자가표시</option>
			<option value="C" <% if (printpriceyn = "C") then %>selected<% end if %>>할인가표시</option>
			<option value="R" <% if (printpriceyn = "R") then %>selected<% end if %>>판매가표시</option>
			<option value="S" <% if (printpriceyn = "S") then %>selected<% end if %>>심플금액표시</option>
			<option value="N" <% if (printpriceyn = "N") then %>selected<% end if %>>금액표시안함</option>
		</select>
		&nbsp;
		* 브랜드표시 :
		<select name="makeriddispyn" onchange="reg('');">
			<option value="Y" <% if (makeriddispyn = "Y") then %>selected<% end if %>>브랜드표시</option>
			<option value="N" <% if (makeriddispyn = "N") then %>selected<% end if %>>브랜드표시안함</option>
		</select>
		&nbsp;
		* 제목표시 :
		<select name="titledispyn" onchange="reg('');">
			<option value="Y" <% if (titledispyn = "Y") then %>selected<% end if %>>제목표시</option>
			<option value="N" <% if (titledispyn = "N") then %>selected<% end if %>>제목표시안함</option>
		</select>

        <!--* 바코드 용지 규격 --->
		<% if printername = "TTP-243_45x22" then %>
			<input type="hidden" name="paperwidth" value="45" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="22" size="4" maxlength=9>
		<% elseif printername = "TTP-243_35x15" then %>
			<input type="hidden" name="paperwidth" value="35" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="15" size="4" maxlength=9>
		<% elseif printername = "TTP-243_45x45" then %>
			<input type="hidden" name="paperwidth" value="45" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="45" size="4" maxlength=9>
		<% elseif printername = "TTP-243_80x50" then %>
			<input type="hidden" name="paperwidth" value="80" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="50" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_45x22" then %>
			<input type="hidden" name="paperwidth" value="450" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="220" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_35x15" then %>
			<input type="hidden" name="paperwidth" value="350" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="150" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_45x45" then %>
			<input type="hidden" name="paperwidth" value="450" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="450" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_80x50" then %>
			<input type="hidden" name="paperwidth" value="800" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="500" size="4" maxlength=9>
		<% end if %>

		<input type="hidden" name="papermargin" value="3" size="4" maxlength=9>

		<script type='text/javascript'>
			jsSetSelectBoxColor();
		</script>
	</td>
</tr>
</table>