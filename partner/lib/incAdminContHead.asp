
								<script type="text/javascript">
								function jsPopHelp(menupos){
									var winHelp = window.open("/partner/lib/popHelp.asp?menupos="+menupos,"popHelp","width=800,height=800,scrollbars=yes,resizable=yes");
									winHelp.focus();
								}
								</script>
								
									<div class="locate"><h2><%=conParentMenuName(conMNum)%> &gt; <strong><%=conChildMenuName(conMNum,comSMNum)%></strong></h2></div>
								<% if conNotice(conMNum,comSMNum) <> "" and conNotice(conMNum,comSMNum)<>"NULL" then %>
									<div class="simpleDesp">
										- <%=conNotice(conMNum,comSMNum)%>
									</div>
								<% end if %>
								<% if conHelp(conMNum,comSMNum) <> "" then %>
									<div class="helpBox" onclick="jsPopHelp('<%=menupos%>');" style="cursor:pointer">
										<dl>
											<dt>HELP</dt>
										</dl>
									</div>
								<% end if %>
							
