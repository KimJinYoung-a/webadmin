<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 메일진
' History : 최초생성자모름
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_mailzinecls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopmailzine_newitemcls.asp" -->
<%
dim idx ,i
idx = requestCheckVar(request("idx"),10)

dim mailmain
set mailmain = new CUploadMaster
mailmain.MailzineView idx

function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function
%>

<script type='text/javascript'>
<!--
function TnSendMail(){
	document.mailform.submit();
}

function TnPOPBig(v){
	  var p = (v);
	  w = window.open("/lib/showimage.asp?img=" + v, "imageView", "status=no,resizable=yes,scrollbars=yes");
}
//-->
</script>

<form method="post" action="/admin/offshop/lib/dooffshopmailzine.asp" name="mailform">
<input type="hidden" name="idx" value="<% = idx %>">
</form>
<html>
<head>
<title>[텐바이텐 오프라인샵 소식지]</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body style="margin: 0 0 0 0">
<table width="584" border=0 cellspacing=0 cellpadding=0 align="center">
<tr>
<td>
<table width="584" border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td>
			<table background="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshoptitle1.jpg" width=230 height=69>
				<tr>
					<td align="center" style="padding: 5 125 0 0" class="verdana-large"><font color="#FFFFFF"><% = left(mailmain.Fregdate,4) %><br><% = FormatDate(mailmain.Fregdate,"00.00") %></font>
					</td>
				</tr>
			</table>
		</td>
		<td style="padding-right:120"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshoptitle2.jpg" width=235 height=69></td>
	</tr>
</table>

<!-- 공지 -->

<table width="584" border=0 cellspacing=0 cellpadding=0 style="border:7 solid #A1A1A0">
	<tr>
		<td align=center>
			<table width="95%" height=34 background="http://imgstatic.10x10.co.kr/offshopmailzine/images/titlenews.jpg">
				<tr><td>&nbsp;</td></tr>
			</table>
			<table width="95%" background="http://imgstatic.10x10.co.kr/offshopmailzine/images/titlebg.jpg">
				<tr>
					<td class="a">
						<span style="line-height:230%;">
						<b><font color="#77BB66"><% = nl2br(mailmain.Fnews) %></font></b>
			 			</span>
 					</td>
 				</tr>
 			</table>
 		</td>
 		
	</tr>
</table>

<!-- 핑거스 -->
<table width="584">
	<tr>
		<td><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopfingers.jpg" width=584 height=36>
		</td>
	</tr>
	<tr>
		<td><a href="<% = mailmain.Furl1 %>" onfocus="this.blur();" target="_blank"><img src="<% = mailmain.Fimg1 %>" border="0"></a></td>
	</tr>
</table>

<%
dim mailnew
set mailnew = new COffshopMailzine
mailnew.FRectMasteridx=idx
mailnew.GetPreNewItem 
%>

<!-- 신규상품  -->
<table width="584">
	<tr>
		<td colspan=2><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopnewarrival.jpg" width=584 height=34></td>
	</tr>
	<tr>
		<td>
			<table  border=0 cellspacing=0 cellpadding=3>
				<% for i=0 to mailnew.FResultCount-1 %>
				<% if i=0 or i=5 or i=10 or i=15 or i=20 or i=25 or i=30 or i=35 then %>
				<tr>
					<td valign=top><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopnew<%= mailnew.FItemList(i).Fcd1 %>.jpg"></td>
				<% end if %>
					<td valign=top align=left>
						<table width="100" border=0 cellspacing=0 cellpadding=0>
							<tr>
								<td width=100 height=100 background="<%= mailnew.FItemList(i).FImageList %>" style="border:1 solid #CCCCCC">
									<a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailnew.FItemList(i).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width="100" height="100"></a>
								</td>
							</tr>
							<tr>
								<td valign=top class="verdana-small"><%= mailnew.FItemList(i).FItemName%><br><font color="#3366FF"><% = FormatNumber(mailnew.FItemList(i).FSellCash,0) %>won</font></td>
							</tr>
						</table>
					</td>
				<% if i=4 or i=9 or i=14 or i=19 or i=24 or i=29 or i=34 then %>
				</tr>
				<% end if %>
				<% next %>
				</tr>
			</table>
		</td>
	</tr>
</table>

<!-- #include virtual="/lib/classes/offshop/offshopmailzine_bestitemcls.asp" -->

<!-- Off line 베스트 -->
<% dim mailoffline 
set mailoffline = new COnOffShopMailzine
mailoffline.FRectMasteridx=idx
mailoffline.FRectgubun="02"
mailoffline.GetOnOffBest
%>


<table width="584">
	<tr>
		<td colspan=2><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopoffbest.jpg" width=584 height=36>
		</td>
	</tr>
	<tr>
		<td width=34 height=450 valign=top background="http://imgstatic.10x10.co.kr/offshopmailzine/images/leftoffline.jpg"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif""></td>
		<td valign=top>
			<table width="100%" border=0 cellspacing=0 cellpadding=2>
				<tr>
					<td topmargin=0 valign=top>
						<% if  ImageExists(mailoffline.FItemList(0).FImageMain) then %>
						<table width="285" height="275" border="0" cellpadding="0" cellspacing="0" background="<% = mailoffline.FItemList(0).FImageMain %>"  style="border:1 solid #CCCCCC;background-repeat: no-repeat;background-position:center;background-attachment:fixed">
						<% else %>
						<table width="285" height="275" border="0" cellpadding="0" cellspacing="0" background="<% = mailoffline.FItemList(0).FImageBasic %>"  style="border:1 solid #CCCCCC;background-repeat: no-repeat;background-position:center;background-attachment:fixed">
						<% end if %>
							<tr>
								<td width=40 height=49 style="padding:8 0 0 5"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopno1.gif" width=40 height=49></td>
								<td width=230></td>
							</tr>
							<tr>
								<td colspan=2><a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailoffline.FItemList(0).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=270 height=213></a></td>
							</tr>
						</table>
						
						<table>
							<tr>
								<td valign=top class="verdana-basic"><strong><%= mailoffline.FItemList(0).FItemName %></strong><br><font color="#3366FF" class="verdana-small"><% = FormatNumber(mailoffline.FItemList(0).FSellCash,0) %>won</font></td>
							</tr>
						</table>
					</td>
					<td topmargin=0 valign=top>
						<table  width=255 border=0 cellspacing=0 cellpadding=0>
							<tr>
							<% for i=1 to 4 %>
								<td valign=top>
									<table width=120 height=116 border=0 cellspacing=0 cellpadding=0 style="border:1 solid #CCCCCC">
										<tr>
											<td valign=top style="padding:5 0 0 0"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopno<%=i+1%>.jpg" width=23></td>
											<td valign=bottom align=right>
												<table width=100 height=100 border=0 cellspacing=0 cellpadding=0>
													<tr>
														<td background="<%= mailoffline.FItemList(i).FImageList %>"><a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailoffline.FItemList(i).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=100 height=100></a></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
									<table>
										<tr>
											<td valign=top width=100 height=32  class="verdana-small"><%= mailoffline.FItemList(i).FItemName %><br><font color="#3366FF"><% = FormatNumber(mailoffline.FItemList(i).FSellCash,0) %>won</font></td>
										</tr>
									</table>
								</td>
								<% if i mod 3=2 then %>
								</tr>
							<tr>
								<% end if %>
								
								<% next %>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table border=0 cellspacing=0 cellpadding=2>
				<tr>
				<% for i=5 to mailoffline.FResultcount -1 %>
					<td>
						<table width=50 height=50 border=0 cellspacing=0 cellpadding=1>
							<tr>
								<td background="<%= mailoffline.FItemList(i).FImageSmall %>" style="border:1 solid #CCCCCC"><a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailoffline.FItemList(i).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=50 height=50 alt="<%=mailoffline.FItemList(i).FItemname%>"></a></td>
							</tr>
						</table>
						<table border=0 cellspacing=0 cellpadding=0>
							<tr>
								<td class="verdana-small"><font color="#3366FF"><% = FormatNumber(mailoffline.FItemList(i).FSellCash,0) %>won</font></td>
							</tr>
						</table>
					</td>
				<% if i mod 14=13 then %>
				</tr>
				<% end if %>
				<% next %>
				</tr>
			</table>
		</td>
	</tr>
</table>

<!-- 온라인 베스트 -->

<% dim mailonline 
set mailonline = new COnOffShopMailzine
mailonline.FRectMasteridx=idx
mailonline.FRectgubun="01"
mailonline.GetOnOffBest
%>

<table width="584">
	<tr>
		<td colspan=2><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshoponbest.jpg" width=584 height=36>
		</td>
	</tr>
	<tr>
		<td width=34 height=450 valign=top background="http://imgstatic.10x10.co.kr/offshopmailzine/images/leftonline.jpg"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif""></td>
		<td valign=top>
			<table width="100%" border=0 cellspacing=0 cellpadding=2>
				<tr>
					<td topmargin=0 valign=top>
						<% if  ImageExists(mailoffline.FItemList(0).FImageMain) then %>
						<table width="285" height="275" border="0" cellpadding="0" cellspacing="0" background="<% = mailonline.FItemList(0).FImageMain %>"  style="border:1 solid #CCCCCC;background-repeat: no-repeat;background-position:center;background-attachment:fixed">
						<% else %>
						<table width="285" height="275" border="0" cellpadding="0" cellspacing="0" background="<% = mailonline.FItemList(0).FImageBasic %>"  style="border:1 solid #CCCCCC;background-repeat: no-repeat;background-position:center;background-attachment:fixed">
						<% end if %>
							<tr>
								<td width=40 height=49 style="padding:8 0 0 5"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopno1.gif" width=40 height=49></td>
								<td width=230></td>
							</tr>
							<tr>
								<td colspan=2><a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailonline.FItemList(0).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=270 height=213></a></td>
							</tr>
						</table>
						<table>
							<tr>
								<td valign=top class="verdana-basic"><strong><%= mailonline.FItemList(0).FItemName %></strong><br><font color="#3366FF" class="verdana-small"><% = FormatNumber(mailonline.FItemList(0).FSellCash,0) %>won</font></td>
							</tr>
						</table>
					</td>
					<td topmargin=0 valign=top>
						<table  width=255 border=0 cellspacing=0 cellpadding=0>
							<tr>
							<% for i=1 to 4 %>
								<td valign=top>
									<table width=120 height=116 border=0 cellspacing=0 cellpadding=0 style="border:1 solid #CCCCCC">
										<tr>
											<td valign=top style="padding:5 0 0 0"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopno<%=i+1%>.jpg" width=23></td>
											<td valign=bottom align=right>
												<table width=100 height=100 border=0 cellspacing=0 cellpadding=0>
													<tr>
														<td background="<%= mailonline.FItemList(i).FImageList %>" valign=bottom><a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailonline.FItemList(i).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=100 height=100></a></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
									<table>
										<tr>
											<td valign=top width=100 height=32  class="verdana-small"><%= mailonline.FItemList(i).FItemName %><br><font color="#3366FF"><% = FormatNumber(mailonline.FItemList(i).FSellCash,0) %>won</font></td>
										</tr>
									</table>
								</td>
								<% if i mod 3=2 then %>
								</tr>
							<tr>
								<% end if %>
								
								<% next %>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table border=0 cellspacing=0 cellpadding=2>
				<tr>
				<% for i=5 to mailonline.FResultcount -1 %>
					<td>
						<table width=50 height=50 border=0 cellspacing=0 cellpadding=0>
							<tr>
								<td background="<%= mailonline.FItemList(i).FImageSmall %>" style="border:1 solid #CCCCCC"><a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mailonline.FItemList(i).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=50 height=50 alt="<%=mailonline.FItemList(i).FItemname%>"></a></td>
							</tr>
						</table>
						<table border=0 cellspacing=0 cellpadding=0>
							<tr>
								<td class="verdana-small"><font color="#3366FF"><% = FormatNumber(mailonline.FItemList(i).FSellCash,0) %>won</font></td>
							</tr>
						</table>
					</td>
				<% if i mod 14=13 then %>
				</tr>
				<% end if %>
				<% next %>
				</tr>
			</table>
		</td>
	</tr>
</table>

<!-- MD추천 브랜드 -->
<%
dim mdbest
set mdbest = new COnOffShopMailzine
mdbest.GetMDitemList idx
%>
<table width="584" border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopmdbrand.jpg" width=584 height=36>
		</td>
	</tr>
	<tr>
		<td>
			<table width="100%" border=0 cellspacing=0 cellpadding=2>
				<tr>
					<td width=240 height=210 valign=top> <img src="<% = mailmain.Fimg2 %>" width=240 height=210 ></td>
					<td valign=top>
						<table border=0 cellspacing=0 cellpadding=3>
							<tr>
								<% for i=0 to mdbest.FResultcount-1 %>
								<td>
									<table border=0 cellspacing=0 cellpadding=0 style="border:1 solid #CCCCCC">
										<tr>
											<td background="<%= mdbest.FItemList(i).FImageList%>">
												<a href="http://10x10.co.kr/shopping/category_prd.asp?itemid=<%= mdbest.FItemList(i).FItemid %>" onfocus="this.blur();" target="_blank"><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/space.gif" border="0" width=100 height=100 alt="<%=mdbest.FItemList(i).FItemname%>"></a></td></a>
											</td>
										</tr>
										<tr>
											<td class="verdana-small"><font color="#3366FF"><% = FormatNumber(mdbest.FItemList(i).FSellCash,0) %>won</font></td>
										</tr>
									</table>
								</td>
							<% if i mod 3=2 then %>
							</tr>
							<tr>
							<% end if %>
							<% next %>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<% dim brand
 brand=split(mailmain.Fbrand,",")
%>
<table width="584" border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td>
			<table border=0 cellspacing=0 cellpadding=1>
				<tr>
				<% for i=0 to 5 %>
					<td><a href="http://10x10.co.kr/street/streetmain.asp?designeid=<%= brand(i) %>"><img src="http://imgstatic.10x10.co.kr/main/brand/best_brand_<%= brand(i) %>.gif" width="190" height="45" vspace="1" style="border:1 solid #CCCCCC"></a></td>
				<% if i mod 3=2 then %>
				</tr>
				<tr>
				<% end if %>
				<% next %>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<!-- 이번주 추천 이벤트  -->
<table width="584" border=0 cellspacing=0 cellpadding=0 >
	<tr>
		<td colspan=2><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopthisweekevent.jpg" width=584 height=37></td>
	</tr>
	<tr>
		<td valign=top style="padding-top:12" align="center">
			<table border=0 cellspacing=0 cellpadding=3>
				<tr><td><a href="<% = mailmain.Furl2 %>" target="_blank"><img src="<% = mailmain.Fimg3 %>" border="0"></a></td></tr>
				<tr><td><a href="<% = mailmain.Furl3 %>" target="_blank"><img src="<% = mailmain.Fimg4 %>" border="0"></a></td></tr>
				<tr><td><a href="<% = mailmain.Furl4 %>" target="_blank"><img src="<% = mailmain.Fimg5 %>" border="0"></a></td></tr>
				<tr><td><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopbotomevent.jpg" border="0"></td></tr>
				<tr><td height=50><br><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopline.jpg" border="0"><br></td></tr>
			</table>
		</td>
		<td align=left valign=top style="padding-top:14;padding-left:12">
			<table border=0 cellspacing=0 cellpadding=0>
				<tr><td><a href="javascript:TnPOPBig('<% = mailmain.Fimg7 %>');"><img src="<% = mailmain.Fimg6 %>" border="0"></a></td></tr>
				<tr><td height=5></td></tr>
				<tr><td><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopbottompop.jpg"></td></tr>
				<tr><td align=left><br><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopbottomsubmit.jpg"></td></tr>
			</table>
	</tr>
</table>
<!-- 하단 배너 -->
<br>
<table width="584">
	<tr>
		<td><img src="http://imgstatic.10x10.co.kr/offshopmailzine/images/offshopbottom.jpg" width="580" height="74">
		</td>
	</tr>
</table>
</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="584" height="50">
<tr>
	<td align="center"><input type="button" value="메일 보내기" onclick="TnSendMail(<% = mailmain.Fidx %>);"></td>
</tr>
</table>
</body>
</html>
<%
set mailmain= nothing
set mailnew=nothing
set mailonline =nothing
set mailoffline= nothing
set mdbest= nothing 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->