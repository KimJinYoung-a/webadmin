<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionReportCls.asp"-->
<%
dim SType '// �з�
dim makerid , mastercode , detailcode , Sdate , Edate
dim i , grpWidth

SType = requestCheckVar(request("SType"),10)
makerid = requestCheckVar(request("makerid"),32)
mastercode = requestCheckVar(request("mastercode"),10)
detailcode = requestCheckVar(request("detailcode"),10)
Sdate = requestCheckVar(request("SDate"),10)
Edate = requestCheckVar(request("EDate"),10)

IF Sdate="" THEN
	Sdate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

IF Edate="" THEN
	Edate= dateSerial(Year(now()),Month(now()),day(now())+1)
End IF

if mastercode = "" then mastercode = 0
if detailcode = "" then detailcode = 0

dim oReport  '// ��� ����Ÿ
set oReport = new ExhibitionReport
	oReport.FRectMakerid = makerid
	oReport.FRectStart = Sdate
	oReport.FRectEnd = dateSerial(year(Edate),month(EDate),Day(EDate))
	oReport.FrectMasterCode = mastercode
	oReport.FrectDetailCode = detailcode

dim t_TotalCost, t_FTotalNo
t_TotalCost = 0
t_FTotalNo  = 0
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link rel="stylesheet" type="text/css" href="/js/jqueryui/css/jquery-ui.css"/>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript' src="/js/jsCal/js/jscal2.js"></script>
<script type='text/javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript">
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function viewImage(div,itemid)
	{
		iframeDB1.location.href = "/admin/report/iframe_viewImage.asp?div="+div+"&itemid="+itemid+"";
	}
</script>

<div class="content scrl" style="top:40px;">
	<div class="pad20">
		<div>
			<h1>��ȹ����� ��¥ , ��ǰ , �귣�� �󼼺���</h1>
		<div>
		<!-- ��� �˻��� ���� -->
		<div class="tPad15">
			<form name="frm" method="get" action="">
			<table class="tbType1 listTb">
			<input type="hidden" name="page" value="1">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
				<tr>
					<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
					<td style="text-align:left">
					�˻��Ⱓ : <input id="SDate" name="SDate" value="<%=Sdate%>" class="text" size="10" maxlength="10"/><img src="http://scm.10x10.co.kr/images/calicon.gif" id="SDate_trigger" border="0" style="cursor:pointer;vertical-align:middle;"/>
                            ~ <input id="Edate" name="Edate" value="<%=Edate%>" class="text" size="10" maxlength="10"/><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="Edate_trigger" border="0" style="cursor:pointer;vertical-align:middle;"/>
                            <script type="text/javascript">
                                var CAL_Start = new Calendar({
                                    inputField : "SDate", trigger    : "SDate_trigger",
                                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                                });
                            </script>
                            <script type="text/javascript">
                                var CAL_End = new Calendar({
                                    inputField : "Edate", trigger    : "Edate_trigger",
                                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                                });
                            </script>
					<br/><br/>
					��� : <% DrawMainPosCodeCombo "mastercode", mastercode ,"" %>
					<% if mastercode > 0 then %>
						<% DrawDetailSelectBox "detailcode" , detailcode , mastercode %>
					<% end if %>
					<br/><br/>
					�з� :
						<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %>> ��¥��
						<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %>> ��ǰ��
						<input type="radio" name="SType" value="M" <% If SType = "M" Then response.write "checked" %>> �귣�庰
					</td>
					<td width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
					</td>
				</tr>
			</table>
			</form>
		</div>

		<div class="tPad15">
			<table class="tbType1 listTb">
				<%
				SELECT CASE SType
					CASE "D" '// ��¥�� �̺�Ʈ ���
						call oReport.GetExhibitionStatisticsByDateDataMart
				%>
				<tr bgcolor="#DDDDFF">
					<td width="90" align="center">������</td>
					<td width="70" align="center">�Ǹž�</td>
					<td width="70" align="center">�ǸŰ���</td>
					<td width="500" align="center">�׷���</td>
				</tr>
				<% if oReport.FResultCount > 0 then %>
				<% for i=0 to oReport.FResultCount-1 %>
				<%
				t_TotalCost = t_TotalCost + oReport.ExhibitionReportList(i).Fselltotal
				t_FTotalNo  = t_FTotalNo + oReport.ExhibitionReportList(i).Fsellcnt
				%>
				<tr bgcolor="#FFFFFF">
					<td align="center"><%= oReport.ExhibitionReportList(i).Fselldate %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></td>
					<td align="right"><%= oReport.ExhibitionReportList(i).Fsellcnt %></td>
					<td width="500" style="text-align:left;">
						<%
							'�׷��� ���� ��� (2008.07.08;������ ����)
							if oReport.maxc>0 then
								grpWidth = Clng(oReport.ExhibitionReportList(i).Fselltotal/oReport.maxc*400)
							else
							grpWidth = 0
							end if
						%>
						<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
					</td>
				</tr>
				<% next %>
				<% end if %>
				<%
					CASE "T"  '// ��ǰ�� �̺�Ʈ ���
						call oReport.GetExhibitionStatisticsByItemDataMart
				%>
				<tr bgcolor="#EDEDFF">
					<td width="150" align="center" rowspan="2">�귣��</td>
					<td width="90" align="center" rowspan="2">�����۹�ȣ</td>
					<td rowspan="2">�̹���</td>
					<td width="70" align="center" colspan="2">��</td>
					<td width="70" align="center" colspan="2">PC��</td>
					<td width="70" align="center" colspan="2">�������</td>
					<td width="70" align="center" colspan="2">APP</td>
					<td width="70" align="center" rowspan="2">Wish</td>
				</tr>
				<tr bgcolor="#EDEDFF">
					<td width="70" align="center">�Ǹž�</td>
					<td width="70" align="center">�ǸŰ���</td>
					<td width="70" align="center">�Ǹž�</td>
					<td width="70" align="center">�ǸŰ���</td>
					<td width="70" align="center">�Ǹž�</td>
					<td width="70" align="center">�ǸŰ���</td>
					<td width="70" align="center">�Ǹž�</td>
					<td width="70" align="center">�ǸŰ���</td>
				</tr>
				<% if oReport.FResultCount > 0 then %>
				<% for i=0 to oReport.FResultCount-1 %>
				<%
				t_TotalCost = t_TotalCost + oReport.ExhibitionReportList(i).Fselltotal
				t_FTotalNo  = t_FTotalNo + oReport.ExhibitionReportList(i).Fsellcnt
				%>
				<tr bgcolor="#FFFFFF">
					<td align="center"><%= oReport.ExhibitionReportList(i).Fmakerid %></td>
					<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oReport.ExhibitionReportList(i).FItemid %>" target="_blank" title="�̸�����"><%= oReport.ExhibitionReportList(i).FItemid %></a></td>
					<td><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(oReport.ExhibitionReportList(i).FItemid)%>/<%=oReport.ExhibitionReportList(i).Fsmallimage%>"></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_PC,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_PC,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_mobile,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_mobile,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_App,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_App,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).FwishCnt,0) %>��</td>
				</tr>
				<% next %>
				<% end if %>
				<%
					CASE "M"  '// �귣�庰 �̺�Ʈ ���
						call oReport.GetExhibitionStatisticsByMakerIDDataMart
				%>
				<tr bgcolor="#DDDDFF">
					<td width="150" align="center">�귣��</td>
					<td width="70" align="center">�Ǹž�</td>
					<td width="70" align="center">�ǸŰ���</td>
					<td width="500" align="center">�׷���</td>
				</tr>
				<% if oReport.FResultCount > 0 then %>
				<% for i=0 to oReport.FResultCount-1 %>
				<%
				t_TotalCost = t_TotalCost + oReport.ExhibitionReportList(i).Fselltotal
				t_FTotalNo  = t_FTotalNo + oReport.ExhibitionReportList(i).Fsellcnt
				%>
				<tr bgcolor="#FFFFFF">
					<td align="center"><%= oReport.ExhibitionReportList(i).Fmakerid %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></td>
					<td align="right"><%= oReport.ExhibitionReportList(i).Fsellcnt %></td>
					<td style="text-align:left;">
						<%
							'�׷��� ���� ��� (2008.07.08;������ ����)
							if oReport.maxc>0 then
								grpWidth = Clng(oReport.ExhibitionReportList(i).Fselltotal/oReport.maxc*400)
							else
								grpWidth = 0
							end if
						%>
						<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
					</td>
				</tr>
				<% next %>
				<% end if %>
				<%
					CASE ELSE
						response.write "�����߻�,�ٽ� �õ�"
					END SELECT
				%>
			</table>
			<div>
				<table class="tbType1 listTb">
					<tr>
						<td> ���ձݾ� <%= FormatNumber(t_TotalCost,0) %> / ���� <%= FormatNumber(t_FTotalNo,0) %></td>
					</tr>
				</table>
			</div>
 		</div>
	</div>
</div>
<%
set oReport = Nothing
%>
<iframe src="about:blank" name="iframeDB1" width="0" height="0">
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
