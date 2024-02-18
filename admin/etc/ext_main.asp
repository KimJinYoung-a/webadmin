<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 상담
' History : 2009.04.17 이상구 생성
'			2016.03.25 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/ext_mainlib.asp"-->
<%

dim i, j, k
dim row, arrRow

dim IsUpdateSellStatusNeed

IsUpdateSellStatusNeed = False

If Trim(application("extTimeSellStatusDiff")) = "" or Not IsArray(application("arrExtSellStatusDiff")) then
	IsUpdateSellStatusNeed = True
	application("extTimeSellStatusDiff") = Now()
end if


dim oExtSellDiff, arrExtSellDiff()
if IsUpdateSellStatusNeed = True then
	Set oExtSellDiff = new CExtMain
	oExtSellDiff.GetExtSellDiffList()
	redim arrExtSellDiff(oExtSellDiff.FResultCount)
	for i = 0 to oExtSellDiff.FResultCount - 1
		row = oExtSellDiff.FItemList(i).Fsellsite
		row = row & vbTab & oExtSellDiff.FItemList(i).Fcnt
		row = row & vbTab & oExtSellDiff.FItemList(i).FtotToBeNotSell
		row = row & vbTab & oExtSellDiff.FItemList(i).FtotToBeSell

		arrExtSellDiff(i) = row
	next
	application("arrExtSellStatusDiff") = arrExtSellDiff
end if

if IsArray(application("arrExtSellStatusDiff")) then
	redim arrExtSellDiff(UBound(application("arrExtSellStatusDiff")))
	for i = 0 to UBound(application("arrExtSellStatusDiff")) - 1
		arrExtSellDiff(i) = application("arrExtSellStatusDiff")(i)
	next
end if

%>
<script language="JavaScript" src="/cscenter/js/convert.date.js"></script>
<script language='javascript'>

var extTimeSellStatusDiff = new Date(getDateFromFormat("<%= application("extTimeSellStatusDiff") %>", "yyyy-MM-dd a h:mm:ss"));

function DisplayClock() {
	var v = new Date();

	var objSellStatusDiff = document.getElementById("objSellStatusDiff");

	objSellStatusDiff.innerHTML = GetDateDiffString(v.getTime() - extTimeSellStatusDiff.getTime());

	setTimeout('DisplayClock();','1000');
}

function GetDateDiffString(v) {
	var result = "";

	if (v < (60 * 1000)) {
		v = v / 1000;
		result = parseInt(v) + "초 전";
	} else if (v < (60 * 60 * 1000)) {
		v = v / (60 * 1000);
		result = parseInt(v) + "분 전";
	} else {
		result =  "1시간 전";
	}

	return result;
}

function RefreshData(v) {
	var frm = document.frm;

	frm.mode.value = "RefreshData";
	frm.extTime.value = v;
	frm.submit();
}

function jsPopStopSellList(sitename) {
	var popwin, menupos;

	if (sitename === "ssg") {
		menupos = <%= CHKIIF(application("Svr_Info")="Dev",3970,3990) %>;
		popwin = window.open("/admin/etc/ssg/ssgitem.asp?stopsell=Y&menupos=" + menupos,"jsPopStopSellList","width=1600 height=900 scrollbars=yes resizable=yes");
		popwin.focus();
	} else if (sitename === "gsshop") {
		menupos = <%= CHKIIF(application("Svr_Info")="Dev",1643,1643) %>;
		popwin = window.open("/admin/etc/gsshop/gsshopItem.asp?stopsell=Y&menupos=" + menupos,"jsPopStopSellList","width=1600 height=900 scrollbars=yes resizable=yes");
		popwin.focus();
	} else if (sitename === "auction") {
		menupos = <%= CHKIIF(application("Svr_Info")="Dev",1643,1643) %>;
		popwin = window.open("/admin/etc/auction/auctionitem.asp?stopsell=Y&menupos=" + menupos,"jsPopStopSellList","width=1600 height=900 scrollbars=yes resizable=yes");
		popwin.focus();
	} else {
		alert("미지정 제휴몰입니다.");
	}
}

function jsPopStartSellList(sitename) {
	var popwin, menupos;

	if (sitename === "ssg") {
		menupos = <%= CHKIIF(application("Svr_Info")="Dev",3970,3990) %>;
		popwin = window.open("/admin/etc/ssg/ssgitem.asp?startsell=Y&menupos=" + menupos,"jsPopStartSellList","width=1600 height=900 scrollbars=yes resizable=yes");
		popwin.focus();
	} else if (sitename === "gsshop") {
		menupos = <%= CHKIIF(application("Svr_Info")="Dev",1643,1643) %>;
		popwin = window.open("/admin/etc/gsshop/gsshopItem.asp?startsell=Y&menupos=" + menupos,"jsPopStopSellList","width=1600 height=900 scrollbars=yes resizable=yes");
		popwin.focus();
	} else if (sitename === "auction") {
		menupos = <%= CHKIIF(application("Svr_Info")="Dev",3743,1751) %>;
		popwin = window.open("/admin/etc/auction/auctionitem.asp?startsell=Y&menupos=" + menupos,"jsPopStartSellList","width=1600 height=900 scrollbars=yes resizable=yes");
		popwin.focus();
	} else {
		alert("미지정 제휴몰입니다.");
	}
}

function ssgEditProcess() {
	var ifr = document.getElementById('act_ssg');
	var params = "/admin/etc/ssg/ssgitem.asp?page=&research=on&makerid=&itemid=&itemname=&ssgGoodNo=&cdl=&ExtNotReg=D&bestOrd=on&sellyn=Y&limityn=&sailyn=&onlyValidMargin=Y&isMadeHand=&isOption=&infodiv=&notinmakerid=&notinitemid=&priceOption=&extsellyn=N&exctrans=N&failCntExists=N&deliverytype=&mwdiv=&MatchCate=&setMargin=&ssgNo10x10Yes=on";
	params = params + "&auto=Y";
	ifr.src = params;
}

function interparkEditProcess() {
	var ifr = document.getElementById('act_interpark');
	var params = "/admin/etc/interpark/interparkItem.asp?page=&research=on&makerid=&itemid=&itemname=&interparkPrdno=&cdl=&ExtNotReg=D&bestOrd=on&sellyn=Y&limityn=&sailyn=&onlyValidMargin=Y&isMadeHand=&isOption=&infodiv=&notinmakerid=&notinitemid=&priceOption=&extsellyn=N&exctrans=N&failCntExists=N&deliverytype=&mwdiv=&MatchCate=&interparkNo10x10Yes=on";
	params = params + "&auto=Y";
	ifr.src = params;
}

function EzwelSelectEditProcess() {
	var ifr = document.getElementById('act_ezwel');
	var params = "/admin/etc/ezwel/ezwelitem.asp?page=&menupos=1719&research=on&makerid=&itemid=&itemname=&EzwelGoodNo=&cdl=&ExtNotReg=D&bestOrd=on&sellyn=Y&limityn=&sailyn=&onlyValidMargin=Y&isMadeHand=&isOption=&infodiv=&notinmakerid=&notinitemid=&priceOption=&extsellyn=N&exctrans=N&failCntExists=N&deliverytype=&mwdiv=&MatchCate=&MatchPrddiv=&EzwelNo10x10Yes=on&getRegdate=";
	params = params + "&auto=Y";
	ifr.src = params;
}

function st11EditProcess() {
	var ifr = document.getElementById('act_st11');
	var params = "/admin/etc/11st/11stItem.asp?page=&menupos=3941&research=on&makerid=&itemid=&itemname=&st11GoodNo=&cdl=&ExtNotReg=D&bestOrd=on&sellyn=Y&limityn=&sailyn=&onlyValidMargin=Y&isMadeHand=&isOption=&infodiv=&notinmakerid=&notinitemid=&priceOption=&extsellyn=N&exctrans=N&failCntExists=N&deliverytype=&mwdiv=&MatchCate=&st11No10x10Yes=on";
	params = params + "&auto=Y";
	ifr.src = params;
}

window.onload = function() {
	DisplayClock();
}

</script>
<style type='text/css'>
form {
	margin:0px;
	padding: 0 0px;
}
</style>
<form name="frm" method="post" action="ext_main_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="extTime" value="">
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<!-- 왼쪽메뉴 시작 -->
		<td width="33%" valign="top">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr valign="top">
					<td>
        				<!-- 판매상태 상이 시작 -->
						<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        					<tr bgcolor="#B2CCFF">
        						<td>
        							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
										<tr height="25">
            								<td style="border-bottom:1px solid #BABABA">
            			    					<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>판매상태 상이 미처리</b>
												(<span id="objSellStatusDiff"></span>) <a href="javascript:RefreshData('extTimeSellStatusDiff')"><img src="/images/icon_reload.gif" border="0"></a>
            								</td>
            								<td align="right" style="border-bottom:1px solid #BABABA">
            			    					&nbsp;
            								</td>
            							</tr>
            							<tr height="25">
											<td>텐텐 판매중지 / 제휴 판매</td>
            								<td align="right">
            								</td>
            							</tr>
										<%
										for i = 0 to UBound(arrExtSellDiff) - 1
											arrRow = Split(arrExtSellDiff(i), vbTab)
										%>
            							<tr height="25">
            								<td>
												&nbsp;&nbsp;*
												<%= arrRow(0) %>
											</td>
            								<td align="right">
        										<%= arrRow(2) %>
												<a href="javascript:jsPopStopSellList('<%= arrRow(0) %>')">
        											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
												</a>
            								</td>
            							</tr>
										<% next %>
            							<tr height="25">
            								<td>텐텐 판매 / 제휴 판매중지</td>
            								<td align="right">
            								</td>
            							</tr>
										<%
										for i = 0 to UBound(arrExtSellDiff) - 1
											arrRow = Split(arrExtSellDiff(i), vbTab)
										%>
            							<tr height="25">
            								<td>
												&nbsp;&nbsp;*
												<%= arrRow(0) %>
											</td>
            								<td align="right">
        										<%= arrRow(3) %>
												<a href="javascript:jsPopStartSellList('<%= arrRow(0) %>')">
        											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
												</a>
            								</td>
            							</tr>
										<% next %>
            						</table>
            					</td>
            				</tr>
            			</table>
        				<!-- 판매상태 상이 시작 -->
            	    </td>
            	</tr>
				<tr valign="top">
					<td height="10"></td>
				</tr>
				<tr valign="top">
					<td>
        				<!-- 판매전환 시작 -->
						<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        					<tr bgcolor="#B2CCFF">
        						<td>
        							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
										<tr height="25">
            								<td style="border-bottom:1px solid #BABABA">
            			    					<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>제휴몰 판매전환</b>
            								</td>
            								<td align="right" style="border-bottom:1px solid #BABABA">
            			    					&nbsp;
            								</td>
            							</tr>
            							<tr height="25">
											<td>SSG <a href="javascript:ssgEditProcess()"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            								<td align="right">
            								</td>
										</tr>
            							<tr height="25">
            								<td>
												<iframe name="act_ssg" id="act_ssg" frameborder="0" width="100%" height="150" style="background: #FFFFFF;"></iframe>
											</td>
										</tr>
            							<tr height="25">
											<td>인터파크 <a href="javascript:interparkEditProcess()"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            								<td align="right">
            								</td>
										</tr>
            							<tr height="25">
            								<td>
												<iframe name="act_interpark" id="act_interpark" frameborder="0" width="100%" height="150" style="background: #FFFFFF;"></iframe>
											</td>
										</tr>
            							<tr height="25">
											<td>이지웰 <a href="javascript:EzwelSelectEditProcess()"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            								<td align="right">
            								</td>
										</tr>
            							<tr height="25">
            								<td>
												<iframe name="act_ezwel" id="act_ezwel" frameborder="0" width="100%" height="150" style="background: #FFFFFF;"></iframe>
											</td>
										</tr>
            							<tr height="25">
											<td>11번가 <a href="javascript:st11EditProcess()"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            								<td align="right">
            								</td>
										</tr>
            							<tr height="25">
            								<td>
												<iframe name="act_st11" id="act_st11" frameborder="0" width="100%" height="150" style="background: #FFFFFF;"></iframe>
											</td>
										</tr>
            						</table>
            					</td>
            				</tr>
            			</table>
        				<!-- 판매전환 종료 -->
            	    </td>
            	</tr>
				<!--
				<tr valign="top">
					<td height="10"></td>
				</tr>
				<tr valign="top">
					<td>

            	    </td>
            	</tr>
				<tr valign="top">
					<td height="10"></td>
				</tr>
				<tr valign="top">
					<td>

            	    </td>
            	</tr>
				<tr valign="top">
					<td height="10"></td>
				</tr>
				<tr valign="top">
					<td>

            	    </td>
            	</tr>
				-->
            	</table>
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
		</td>
		<!-- 왼쪽메뉴 끝 -->
		<td width="10"></td>
		<!-- 가운데메뉴 시작 -->
		<td width="33%" valign="top">

		</td>
		<!-- 가운데메뉴 끝 -->
		<td width="10"></td>
		<!-- 오픈쪽메뉴 시작 -->
		<td valign="top">

		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
