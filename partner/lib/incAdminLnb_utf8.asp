<%
session.codePage = 65001
%>
<%
'###########################################################
' Description : 헤더 lnb utf8
' Hieditor : 2017.03.22 한용민 생성
'###########################################################
%>
 
						<div class="lnbSection">
    <% if (ssBctDiv_IsAcademy) then %>
   							<div class="homeArea"><a href="/diyadmin/index.asp"><span class="cBk1">ADMIN HOME</span></a></div>
    <% else %>
							<div class="homeArea"><a href="/partner/index.asp"><span class="cBk1">ADMIN HOME</span></a></div>
    <% end if %>
							<div class="lnbWrap scrl">
	<%
		Dim conMNum, comSMNum
		if menupos <> "" then  '-- 메뉴번호 있을때만 서브메뉴 보여준다. 
  
			if ubound(split(menupos,"^"))>0 then
				conMNum = split(menupos,"^")(0)
				comSMNum = split(menupos,"^")(1)
 	%>
								<div class="lnb">
									<dl class="current">
										<dt><span><%=conParentMenuName(conMNum)%></span></dt>
										<dd>
											<ul>
												<%For conLoop = 0 To conChildMenuSize(conMNum)%>
												<li <%IF Cstr(comSMNum) = Cstr(conLoop) THEN%>class="on"<%END IF%>>
													<span>
														<% if instr(Replace(conChildMenuLinkUrl(conMNum,conLoop),"/lectureadmin/","/diyadmin/"), "?") > 0 then %>
															<a href="<%=Replace(conChildMenuLinkUrl(conMNum,conLoop),"/lectureadmin/","/diyadmin/")%>&menupos=<%=conMNum%>^<%=conLoop%>">
														<% else %>
															<a href="<%=Replace(conChildMenuLinkUrl(conMNum,conLoop),"/lectureadmin/","/diyadmin/")%>?menupos=<%=conMNum%>^<%=conLoop%>">
														<% end if %>

														<%=conChildMenuName(conMNum,conLoop)%></a>
													</span>
												</li>
												<%Next%>
											</ul>
										</dd>
									</dl>
								</div>
		<% end if %>
<% end if %>
<!-- 추가-->					
							<div style="margin-top:25px; padding:12px 5px; text-align:center; background-color:#e8e8e8; font-family:'malgun Gothic','맑은고딕', Dotum, '돋움', sans-serif; border:1px dashed #ddd">
							<strong style="font-size:13px;">고객주문 관련 문의</strong><br /><span style="font-size:11px; color:#666;">(그외 파트별담당자문의)</span>
							<div style="background-color:#fff; padding:10px; margin-top:10px;">
								<strong style="font-family:'malgun Gothic','맑은고딕', Dotum, '돋움', sans-serif; font-size:18px; color:#00cccc; text-shadow:1px 1px rgba(0,51,51,0.4);">070-4868-1799</strong>
								<table style="width:97%; font-size:11px; color:#999; margin:10px auto 0 auto; line-height:13px;">
									<tr>
										<th style="text-align:left;">평일</th>
										<td style="text-align:right;">10:00 ~ 05:00</td>
									</tr>
									<tr>
										<th style="text-align:left;">점심시간</th>
										<td style="text-align:right;">12:00 ~ 01:00</td>
									</tr>
									<tr>
										<td colspan="2"  style="text-align:center;">토/일ㆍ공휴일 휴무</td>
									</tr>
								</table>
							</div>
						</div>
						<!--//-->
							</div>								
						</div>
					 
