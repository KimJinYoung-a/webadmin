<%
'###########################################################
' Description : ��� gnb
' Hieditor : 2016.11.24 ������ ����
'			 2016.12.27 �ѿ�� ����(menupos �Ķ��Ÿ ��ũ ��� ����)
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
		if menupos <> "" then  '-- �޴���ȣ �������� ����޴� �����ش�. 
  
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
<!-- �߰�-->					
							<div style="margin-top:25px; padding:12px 5px; text-align:center; background-color:#e8e8e8; font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; border:1px dashed #ddd">
							<strong style="font-size:13px;">���ֹ� ���� ����</strong><br /><span style="font-size:11px; color:#666;">(�׿� ��Ʈ������ڹ���)</span>
							<div style="background-color:#fff; padding:10px; margin-top:10px;">
								<strong style="font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; font-size:18px; color:#00cccc; text-shadow:1px 1px rgba(0,51,51,0.4);">070-4868-1799</strong>
								<table style="width:97%; font-size:11px; color:#999; margin:10px auto 0 auto; line-height:13px;">
									<tr>
										<th style="text-align:left;">����</th>
										<td style="text-align:right;">10:00 ~ 05:00</td>
									</tr>
									<tr>
										<th style="text-align:left;">���ɽð�</th>
										<td style="text-align:right;">12:00 ~ 01:00</td>
									</tr>
									<tr>
										<td colspan="2"  style="text-align:center;">��/�Ϥ������� �޹�</td>
									</tr>
								</table>
							</div>
						</div>
						<!--//-->
							</div>								
						</div>
					 
