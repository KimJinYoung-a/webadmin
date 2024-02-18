<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2, rdsite
dim fromDate,toDate,totalmoney,totalea
dim minusTotal,minusCnt
dim channelDiv, chkOldJumun

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
rdsite = request("rdsite")
channelDiv  = NullFillWith(request("channelDiv"),"")
chkOldJumun = NullFillWith(request("chkOld"),"")


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
oreport.FRectRegStart = fromDate
oreport.FRectRegEnd = toDate
'oreport.FRectRdsite = rdsite
oreport.FRectOldJumun = chkOldJumun
oreport.FRectSellChannelDiv = channelDiv
oreport.SearchTimeSellrePort

dim i,p1,p2
%>

<script type='text/javascript'>
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function chkForm() {
	var frm = document.frm;
	var vDt1=new Date(frm.yyyy1.value,frm.mm1.value,frm.dd1.value).valueOf();
	var vDt2=new Date(frm.yyyy2.value,frm.mm2.value,frm.dd2.value).valueOf();
	var chkDateGap=(vDt2-vDt1)/(1000*60*60*24);		//�����̸� ���ѵ� �Ϸ翡 �ش��ϴ� ������ ���Ͽ�, �ʴ����� �ϴ����� ��ȯ

	if(chkDateGap<0) {
		alert("�˻� �Ⱓ�� Ȯ�����ּ���.");
		return;
	}

	if(frm.chkOld.checked && chkDateGap>92) {
		alert("6���� ���� �˻��� 3�����̳��θ� �˻��� �����մϴ�.");
		return;
	}

	frm.submit();
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="" onsubmit="chkForm(); return false;">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			�˻��Ⱓ :
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> &nbsp;&nbsp;
		<!--	<input type="checkbox" name="rdsite" <% if rdsite="on" then response.write "checked" %> >������ǸŸ�-->
			&nbsp; / &nbsp;
                	ä�α��� 
                	   <% drawSellChannelComboBoxGroup "channelDiv",channelDiv %>  
                <!--	<select name="channelDiv">
                	<option value="" <%=CHKIIF(channelDiv="","selected","") %> >��ü</option>
                	<option value="web" <%=CHKIIF(channelDiv="web","selected","") %> >��</option>
                	<option value="jaehu" <%=CHKIIF(channelDiv="jaehu","selected","") %> >����</option>
                	<option value="mjaehu" <%=CHKIIF(channelDiv="mjaehu","selected","") %> >���������</option>
                	<option value="mobile" <%=CHKIIF(channelDiv="mobile","selected","") %> >�����</option>
                	<option value="ipjum" <%=CHKIIF(channelDiv="ipjum","selected","") %> >��ǰ����</option>
                	</select>
                	<a href="javascript:image_view('http://webadmin.10x10.co.kr/admin/maechul/statistic/ch_gubun_exp.jpg');">[����]</a>-->
			&nbsp; / &nbsp; <input type="checkbox" name="chkOld" value="on" <%=chkIIF(chkOldJumun="on","checked","")%>> 6��������
		</td>
		<td class="a" align="right">
			<a href="javascript:chkForm();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">�Ⱓ</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="120"><font color="#FFFFFF">����</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
		%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
		  <td width="120" height="10">
          	<%= oreport.FMasterItemList(i).Fdpart %>��
          </td>
          <td  height="10" width="600">
			<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= p1 %>%"></div><br>
          	<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p2 %>%"></div>
          </td>
		  <td class="a" width="160" align="right">
		    <%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>�� <br>
          	<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>�� <br>
          	(<%= FormatNumber(oreport.FMasterItemList(i).Fminustotal,0) %>�� / <%= FormatNumber(oreport.FMasterItemList(i).Fminuscount,0) %> ��)
		  </td>
        </tr>
		<% 
		totalmoney = totalmoney + oreport.FMasterItemList(i).Fselltotal
		totalea = totalea + oreport.FMasterItemList(i).Fsellcnt
		'// ��Ұ� ����
		minusTotal = minusTotal + oreport.FMasterItemList(i).Fminustotal
		minusCnt = minusCnt + oreport.FMasterItemList(i).Fminuscount
		%>
        <% next %>
		<tr>
			<td colspan="3" align="right" bgcolor="#FFFFFF" class="a">
				�ѱݾ� : <font color="red"><% = FormatNumber(totalmoney,0) %></font> �Ѽ��� : <font color="#000099"><% = FormatNumber(totalea,0) %></font>&nbsp;/&nbsp;
				(��ǰ�Ѿ� : <% = FormatNumber(minusTotal,0) %>�� / �ѹ�ǰ�� : <% = FormatNumber(minusCnt,0) %>��)&nbsp;&nbsp;&nbsp;&nbsp;
			</td>
		</tr>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->