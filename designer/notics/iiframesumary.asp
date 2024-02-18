<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<SCRIPT LANGUAGE="JavaScript">
function loadImages() {
		if (document.getElementById) {  // DOM3 = IE5, NS6
		document.getElementById('hidepage').style.display = 'none';
		}
		else {
		if (document.layers) {  // Netscape 4
		document.hidepage.display = 'none';
		}
		else {  // IE 4
		document.all.hidepage.style.display = 'none';
      }
   }
}
</script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>

<body OnLoad="loadImages()">
<div id="hidepage">
<table width="50%" height="105" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="10" valign="bottom">
		<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr>
		<td background="/images/tbl_blue_round_04.gif"></td>
	    <td align="center">
	    	<b><font color="blue">
	    	금일 미처리 목록을 로딩하고 있습니다.<br>
	    	잠시만 기다려주세요....<br>
	    	</font></b>
	    </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>
</div>


<% 'response.flush %>
<%
dim sqlstr
dim ibalju
dim yyyy1,mm1,dd1,nowdate
dim mibaljuCount, mibeasongCount

nowdate = Left(CStr(now()),10)
yyyy1 = Left(nowdate,4)
mm1   = Mid(nowdate,6,2)
dd1   = Mid(nowdate,9,2)


''업체배송 상품이 있는경우만 검색하게 변경
dim upchebeasongExists
''sqlStr = "select count(itemid) as cnt "
''sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
''sqlStr = sqlStr + " where makerid='" + session("ssBctId") + "'"
''sqlStr = sqlStr + " and mwdiv='U'"
sqlStr = "select IsNULL(sum(smKeyValInt),0) as cnt from db_partner.dbo.tbl_partner_summaryInfo"
sqlStr = sqlStr + " where makerid='" + session("ssBctId") + "'"
sqlStr = sqlStr + " and smKeyName in ('UDTT','UDFX','UD0','UD2','UD3')"
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	upchebeasongExists = rsget("cnt")>0
rsget.Close







dim csnofincnt, itemqanotfinish, eventotfinish, tmpsoldoutItemCnt
csnofincnt          = 0
itemqanotfinish     = 0
eventotfinish       = 0
tmpsoldoutItemCnt   = 0

''CS 미처리 갯수
sqlStr = "exec [db_cs].[dbo].sp_Ten_upcheCsCount '" + CStr(session("ssBctID")) +  "'"  + vbcrlf
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
  csnofincnt = rsget("cnt")
rsget.Close

''상품문의 미처리
sqlstr = "select count(*) as cnt"
sqlstr = sqlstr + " from [db_cs].[dbo].tbl_my_item_qna"
sqlstr = sqlstr + " where makerid='" + session("ssBctID") + "'"
sqlstr = sqlstr + " and isusing='Y'"
sqlstr = sqlstr + " and replyuser=''"
if application("Svr_Info") <> "Dev" then
	sqlstr = sqlstr + " and id >= 400000 "
end if

rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	itemqanotfinish = rsget("cnt")
rsget.Close

''미처리사은품발송건 ''배송요청일 -2 추가.
sqlstr = "select count(*) as cnt"
sqlstr = sqlstr + " from [db_sitemaster].[dbo].tbl_etc_songjang w"
sqlstr = sqlstr + " where w.delivermakerid='" + session("ssBctID") + "' and w.deleteyn='N' and ((w.songjangno is NULL) or (w.songjangno='')) and w.isupchebeasong='Y' and datediff(d,reqdeliverdate,getdate())>=-2"
IF application("Svr_Info") <> "Dev" THEN
	sqlstr = sqlstr + " and w.id >= 120000 "
end if
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	eventotfinish = rsget("cnt")
rsget.Close

''주문서 미확인/미출고 갯수

dim logisnotconfirmcnt, logisnotsendcnt

sqlstr = "select  count(idx) as cnt" + VbCrlf
sqlstr = sqlstr + " from [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
sqlstr = sqlstr + " where targetid='" + session("ssBctID") + "'" + VbCrlf
sqlstr = sqlstr + " and baljuid='10x10'" + VbCrlf
sqlstr = sqlstr + " and statecd='0'" + VbCrlf
sqlstr = sqlstr + " and deldt is null" + VbCrlf
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	logisnotconfirmcnt = rsget("cnt")
rsget.Close


sqlstr = "select  count(idx) as cnt" + VbCrlf
sqlstr = sqlstr + " from [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
sqlstr = sqlstr + " where targetid='" + session("ssBctID") + "'" + VbCrlf
sqlstr = sqlstr + " and baljuid='10x10'" + VbCrlf
sqlstr = sqlstr + " and statecd='1'" + VbCrlf
sqlstr = sqlstr + " and deldt is null" + VbCrlf
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	logisnotsendcnt = rsget("cnt")
rsget.Close

''일시품절상품갯수
sqlstr = "select  count(i.itemid) as cnt" + VbCrlf
sqlstr = sqlstr + " from db_item.dbo.tbl_item i" + VbCrlf
sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid "
sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option_stock ot "
sqlstr = sqlstr + "         on ot.itemgubun='10'  "
sqlstr = sqlstr + "         and ot.itemid=i.itemid "
sqlstr = sqlstr + "         and ot.itemoption=IsNULL(v.itemoption,'0000') "
sqlstr = sqlstr + " where i.makerid='" + session("ssBctID") + "'" + VbCrlf
sqlstr = sqlstr + " and i.sellyn='S'" + VbCrlf
sqlstr = sqlstr + " and i.mwdiv in ('M','W')" + VbCrlf
sqlstr = sqlstr + " and i.isusing='Y'"                      ''확인
sqlstr = sqlstr + " and i.danjongyn in ('S','N')"
sqlstr = sqlstr + " and ot.stockreipgodate is NULL"
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	tmpsoldoutItemCnt = rsget("cnt")
rsget.Close

set ibalju = new CJumunMaster
''최근 2주간
ibalju.FRectRegStart = DateSerial(yyyy1,mm1, dd1-15)
ibalju.FRectRegEnd = DateSerial(yyyy1,mm1, dd1+1)
ibalju.FRectDesignerID = CStr(session("ssBctID"))

if (upchebeasongExists) then
	Call ibalju.DesignerDateMiBaljuMiBeasongCount(mibaljuCount, mibeasongCount)
	''수정중..
	''mibaljuCount = "..."
	''mibeasongCount = "..."
end if

'''핑거스 DIY 주문 관련
Dim IsFingersDIYOrderCNT : IsFingersDIYOrderCNT = 0
Dim diyuserid
'' 핑거스 아이디 존재할경우
sqlStr = "select top 1 p1.id from db_partner.dbo.tbl_partner p1"
sqlStr = sqlStr & "	 Join db_partner.dbo.tbl_partner p2"
sqlStr = sqlStr & " on p2.id='" + session("ssBctID") + "'"
sqlStr = sqlStr & " and p1.groupid=p2.groupid"
sqlStr = sqlStr & " and p1.id<>'" + session("ssBctID") + "'"
sqlStr = sqlStr & " Join db_user.dbo.tbl_user_c c"
sqlStr = sqlStr & " on p1.id=c.userid"
sqlStr = sqlStr & " and c.userdiv=14"
sqlStr = sqlStr & " and p1.isusing='Y' "

rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
If Not rsget.Eof then
    diyuserid = rsget("id")
End IF
rsget.Close

IF (diyuserid<>"") then
    ''sqlStr = "exec [db_academy].[dbo].sp_Academy_Upche_Mibalju_CNT '" + diyuserid +  "'"  + vbcrlf
    ''connection 오류로 바꿈. (디비 78로 이전후 : 원인불명..)
    sqlStr = "exec [db_partner].[dbo].sp_Academy_Upche_Mibalju_CNT '" + diyuserid +  "'"  + vbcrlf
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
      IsFingersDIYOrderCNT = rsget("cnt")
    rsget.Close
END IF


'''오프라인 매장배송 관련
dim ISOffDlvBrand
sqlStr = "select count(*) as CNT from db_shop.dbo.tbl_shop_designer where makerid='"&session("ssBctID")&"'"
sqlStr = sqlStr & " and defaultbeasongdiv=2"
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
If Not rsget.Eof then
    ISOffDlvBrand = rsget("CNT")>0
End IF
rsget.Close

Dim offmibaljuCount,offmibeaCount
offmibaljuCount =0
offmibeaCount   =0

IF (ISOffDlvBrand) then
    ''' 매장 업배
    sqlStr = "exec [db_shop].[dbo].[sp_Ten_Shop_Upche_MibaljuMibeasong_Count] '" + CStr(session("ssBctID")) +  "'"  + vbcrlf
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    If not rsget.Eof then
        offmibaljuCount = rsget("MiBaljuCnt")
        offmibeaCount   = rsget("MiBeasongCnt")
    end if
    rsget.Close
END IF

%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td width="32%">
			<!-- 업체배송관련창 시작 -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td>
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>업체배송관련 미처리목록</b>
				    </td>
				</tr>
				<tr height="80" bgcolor="#FFFFFF">
					<td>
						<% if upchebeasongExists then %>

						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
					    	<tr height="20">
							    <td>2주간 미확인 주문건</td>
							    <td align="right">
							    	<b><%= mibaljuCount %></b> 건
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/datebaljulist.asp?menupos=112">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>2주간 미출고 주문건(미출고사유 입력)</td>
							    <td align="right">
							    	<b><%= mibeasongCount %></b> 건
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/upchemibeasonglist.asp?menupos=969">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>

							<tr height="20">
								<td>2주간 미출고 주문건(송장 입력)</td>
							    <td align="right">
							    	<b><%= mibeasongCount %></b> 건
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/upchesendsongjanginput.asp?menupos=96" class="a">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<% if (diyuserid<>"") then %>
							<tr height="18">
							    <td>아카데미 미확인 주문건</td>
							    <td align="right">
							    	<b><%= IsFingersDIYOrderCNT %></b> 건

							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom" style="cursor:pointer" onClick="alert('상단 브랜드목록에서 로그인 변경후 \n\n [Fingers] DIY 상품주문관리>>업체배송주문확인 메뉴에서 확인 가능합니다.');">
							    </td>
							</tr>
							<% end if %>
						</table>

						<% else %> <!-- 업체배송없는경우 -->

						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
					    	<tr height="60">
							    <td align="center">
							    	<b><font color="blue"><img src="/images/exclam.gif" border="0" align="absbottom">진행하는 업체배송상품이 없습니다.</font></b>
							    </td>
							</tr>
						</table>

						<% end if %>

					</td>
				</tr>
			</table>
		</td>

		<td width="2%" bgcolor="#F4F4F4"></td>

		<td width="32%">
			<!-- 상품문의 및 CS접수 -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td>
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품문의 및 CS접수관련 미처리목록</b>
				    </td>
				</tr>
				<tr height="80" bgcolor="#FFFFFF">
					<td>
						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
							<tr height="20">
							    <td>미처리 상품문의건</td>
							    <td align="right">
							    	<b><%= itemqanotfinish %></b> 건
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/board/newitemqna_list.asp?menupos=295">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>미처리 CS접수건(장기간(3개월) 미처리건 제외)</td>
							    <td align="right">
							    	<b><%= csnofincnt %></b> 건
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/upchecslist.asp?menupos=249">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>미처리 사은품 발송건</td>
							    <td align="right">
							    	<b><%= eventotfinish %></b> 건
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/event/deliverinfolist.asp?menupos=980">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!-- 상품문의 및 CS접수 -->
		</td>

		<td width="2%" bgcolor="#F4F4F4"></td>

		<td width="32%">
			<!-- 물류창고 상품입관련 -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td>
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>물류센터 상품입고관련 미처리목록</b>
				    </td>
				</tr>
				<tr height="80" bgcolor="#FFFFFF">
					<td>
						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
							<tr height="20">
							    <td>미확인 발주건</td>
							    <td align="right">
							    	<b><%= logisnotconfirmcnt %></b> 건
							    	<a target=_parent href="/designer/storage/orderlist.asp?menupos=538&research=on&statecd=0">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>주문확인후 미발송건</td>
							    <td align="right">
							    	<b><%= logisnotsendcnt %></b> 건
							    	<a target=_parent href="/designer/storage/orderlist.asp?menupos=538&research=on&statecd=1">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>일시품절상품(텐바이텐배송)</td>
							    <td align="right">
							    	<b><%= tmpsoldoutItemCnt %></b> 건
							    	<a target=_parent href="/designer/itemmaster/upche_danjong_set.asp?menupos=1065">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!-- 물류창고 상품입관련 -->
		</td>

	</tr>
	<% if (ISOffDlvBrand) then %>
	<tr>
	    <td colspan="5" height="5" bgcolor="#F4F4F4"></td>
	</tr>
	<tr>
		<!-- 매장 직배송 관련 -->
		<td colspan="5" align="left" bgcolor="#F4F4F4">
		    <table width="32%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td colspan="3">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>매장주문 배송 미처리목록</b>
				    </td>
				</tr>
				<tr height="40" bgcolor="#FFFFFF">
					<td width="32%">
					    <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				    	<tr height="20">
						    <td> 미확인 주문건</td>
						    <td align="right">
						    	<b><%= offmibaljuCount %></b> 건
						    	<a target=_parent href="/common/offshop/beasong/upche_datebaljulist.asp?menupos=1301">
						    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
						    </td>
					    </tr>
					    <tr height="20">
						    <td> 미출고 주문건</td>
						    <td align="right">
						    	<b><%= offmibeaCount %></b> 건
						    	<a target=_parent href="/common/offshop/beasong/upche_sendsongjanginput.asp?menupos=1303">
						    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
						    </td>
					    </tr>
					    </table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
    <% end if %>
</table>
</body>
</html>

<%
set ibalju = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
