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
	    	���� ��ó�� ����� �ε��ϰ� �ֽ��ϴ�.<br>
	    	��ø� ��ٷ��ּ���....<br>
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


''��ü��� ��ǰ�� �ִ°�츸 �˻��ϰ� ����
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

''CS ��ó�� ����
sqlStr = "exec [db_cs].[dbo].sp_Ten_upcheCsCount '" + CStr(session("ssBctID")) +  "'"  + vbcrlf
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
  csnofincnt = rsget("cnt")
rsget.Close

''��ǰ���� ��ó��
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

''��ó������ǰ�߼۰� ''��ۿ�û�� -2 �߰�.
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

''�ֹ��� ��Ȯ��/����� ����

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

''�Ͻ�ǰ����ǰ����
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
sqlstr = sqlstr + " and i.isusing='Y'"                      ''Ȯ��
sqlstr = sqlstr + " and i.danjongyn in ('S','N')"
sqlstr = sqlstr + " and ot.stockreipgodate is NULL"
rsget.CursorLocation = adUseClient 
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly
	tmpsoldoutItemCnt = rsget("cnt")
rsget.Close

set ibalju = new CJumunMaster
''�ֱ� 2�ְ�
ibalju.FRectRegStart = DateSerial(yyyy1,mm1, dd1-15)
ibalju.FRectRegEnd = DateSerial(yyyy1,mm1, dd1+1)
ibalju.FRectDesignerID = CStr(session("ssBctID"))

if (upchebeasongExists) then
	Call ibalju.DesignerDateMiBaljuMiBeasongCount(mibaljuCount, mibeasongCount)
	''������..
	''mibaljuCount = "..."
	''mibeasongCount = "..."
end if

'''�ΰŽ� DIY �ֹ� ����
Dim IsFingersDIYOrderCNT : IsFingersDIYOrderCNT = 0
Dim diyuserid
'' �ΰŽ� ���̵� �����Ұ��
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
    ''connection ������ �ٲ�. (��� 78�� ������ : ���κҸ�..)
    sqlStr = "exec [db_partner].[dbo].sp_Academy_Upche_Mibalju_CNT '" + diyuserid +  "'"  + vbcrlf
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
      IsFingersDIYOrderCNT = rsget("cnt")
    rsget.Close
END IF


'''�������� ������ ����
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
    ''' ���� ����
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
			<!-- ��ü��۰���â ���� -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td>
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ü��۰��� ��ó�����</b>
				    </td>
				</tr>
				<tr height="80" bgcolor="#FFFFFF">
					<td>
						<% if upchebeasongExists then %>

						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
					    	<tr height="20">
							    <td>2�ְ� ��Ȯ�� �ֹ���</td>
							    <td align="right">
							    	<b><%= mibaljuCount %></b> ��
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/datebaljulist.asp?menupos=112">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>2�ְ� ����� �ֹ���(�������� �Է�)</td>
							    <td align="right">
							    	<b><%= mibeasongCount %></b> ��
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/upchemibeasonglist.asp?menupos=969">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>

							<tr height="20">
								<td>2�ְ� ����� �ֹ���(���� �Է�)</td>
							    <td align="right">
							    	<b><%= mibeasongCount %></b> ��
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/upchesendsongjanginput.asp?menupos=96" class="a">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<% if (diyuserid<>"") then %>
							<tr height="18">
							    <td>��ī���� ��Ȯ�� �ֹ���</td>
							    <td align="right">
							    	<b><%= IsFingersDIYOrderCNT %></b> ��

							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom" style="cursor:pointer" onClick="alert('��� �귣���Ͽ��� �α��� ������ \n\n [Fingers] DIY ��ǰ�ֹ�����>>��ü����ֹ�Ȯ�� �޴����� Ȯ�� �����մϴ�.');">
							    </td>
							</tr>
							<% end if %>
						</table>

						<% else %> <!-- ��ü��۾��°�� -->

						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
					    	<tr height="60">
							    <td align="center">
							    	<b><font color="blue"><img src="/images/exclam.gif" border="0" align="absbottom">�����ϴ� ��ü��ۻ�ǰ�� �����ϴ�.</font></b>
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
			<!-- ��ǰ���� �� CS���� -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td>
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ���� �� CS�������� ��ó�����</b>
				    </td>
				</tr>
				<tr height="80" bgcolor="#FFFFFF">
					<td>
						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
							<tr height="20">
							    <td>��ó�� ��ǰ���ǰ�</td>
							    <td align="right">
							    	<b><%= itemqanotfinish %></b> ��
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/board/newitemqna_list.asp?menupos=295">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>��ó�� CS������(��Ⱓ(3����) ��ó���� ����)</td>
							    <td align="right">
							    	<b><%= csnofincnt %></b> ��
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/jumunmaster/upchecslist.asp?menupos=249">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>��ó�� ����ǰ �߼۰�</td>
							    <td align="right">
							    	<b><%= eventotfinish %></b> ��
							    	<a target=_parent href="<%=getSCMSSLURL%>/designer/event/deliverinfolist.asp?menupos=980">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!-- ��ǰ���� �� CS���� -->
		</td>

		<td width="2%" bgcolor="#F4F4F4"></td>

		<td width="32%">
			<!-- ����â�� ��ǰ�԰��� -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td>
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������� ��ǰ�԰���� ��ó�����</b>
				    </td>
				</tr>
				<tr height="80" bgcolor="#FFFFFF">
					<td>
						<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
							<tr height="20">
							    <td>��Ȯ�� ���ְ�</td>
							    <td align="right">
							    	<b><%= logisnotconfirmcnt %></b> ��
							    	<a target=_parent href="/designer/storage/orderlist.asp?menupos=538&research=on&statecd=0">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>�ֹ�Ȯ���� �̹߼۰�</td>
							    <td align="right">
							    	<b><%= logisnotsendcnt %></b> ��
							    	<a target=_parent href="/designer/storage/orderlist.asp?menupos=538&research=on&statecd=1">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
							<tr height="20">
							    <td>�Ͻ�ǰ����ǰ(�ٹ����ٹ��)</td>
							    <td align="right">
							    	<b><%= tmpsoldoutItemCnt %></b> ��
							    	<a target=_parent href="/designer/itemmaster/upche_danjong_set.asp?menupos=1065">
							    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
							    </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<!-- ����â�� ��ǰ�԰��� -->
		</td>

	</tr>
	<% if (ISOffDlvBrand) then %>
	<tr>
	    <td colspan="5" height="5" bgcolor="#F4F4F4"></td>
	</tr>
	<tr>
		<!-- ���� ����� ���� -->
		<td colspan="5" align="left" bgcolor="#F4F4F4">
		    <table width="32%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
				    <td colspan="3">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�����ֹ� ��� ��ó�����</b>
				    </td>
				</tr>
				<tr height="40" bgcolor="#FFFFFF">
					<td width="32%">
					    <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				    	<tr height="20">
						    <td> ��Ȯ�� �ֹ���</td>
						    <td align="right">
						    	<b><%= offmibaljuCount %></b> ��
						    	<a target=_parent href="/common/offshop/beasong/upche_datebaljulist.asp?menupos=1301">
						    	<img src="/images/icon_arrow_link.gif" border="0" align="absbottom"></a>
						    </td>
					    </tr>
					    <tr height="20">
						    <td> ����� �ֹ���</td>
						    <td align="right">
						    	<b><%= offmibeaCount %></b> ��
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
