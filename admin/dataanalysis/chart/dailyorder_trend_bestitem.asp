<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim MaxPageSize : MaxPageSize = 500
Dim oChart, vArr1, vArr2, i, j
Dim vSDate, vChannel, vOrdType, rdsitegrp

vSDate = requestCheckvar(request("startdate"),10)
vChannel = requestCheckvar(request("channel"),10)
vOrdType = requestCheckvar(request("ordtype"),32)
rdsitegrp = requestCheckvar(request("rdsitegrp"),32)

if (vOrdType="") then vOrdType="C" ''�Ǽ�(C) , �ݾ�(S), ����(G)

If vSDate = "" Then
	vSDate = dateadd("d",-0,Date())
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectChannel = vChannel
	oChart.FRectOrderType = vOrdType
	oChart.FRectRdsiteGrp = rdsitegrp
	oChart.FPageSize = MaxPageSize
	
	
	vArr2 = oChart.fnDailyMeachul_bestitem_DW
SET oChart = nothing

Dim vSum1,vSum2,vSum3, vSum4,vSum5, vSum6,vSum7, vSum8
Dim isellStr, iLimitStr, priceStr, iSellyn, iLimityn, iLimitNo

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script>
$(function() {
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			//CAL_End.args.min = date;
			//CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	
});

function goSearch(){
	if($("#sdate").val() == ""){
		alert("�������� �Է��ϼ���");	
		return false;
	}
	if($("#edate").val()== ""){
		alert("�������� �Է��ϼ���");	
		return false;
	}
	document.frm1.submit();
}
</script>
<body>
<p>
<form name="frm1" method="get" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	��ȸ��¥ : 
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;
    
    ä�� : <% call drawConversionChannelSelectBox("channel",vChannel) %>
    &nbsp;&nbsp;
    
     |
    &nbsp;&nbsp;
    
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >�ֹ��Ǽ���
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >�����Ѿ׼�
    <!--
    <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >������ͼ�
    -->
    &nbsp;&nbsp;
    |
    &nbsp;&nbsp;
    rdsiteŸ�� : <% call drawConversionTypeGroupSelectBox2_DW("rdsitegrp",rdsitegrp,"rdsite",2,"") %>
   
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<br />
<%
if isArray(vArr2) then
    if (UBound(vArr2,2)>=MaxPageSize-1) then
      response.write "�ִ� "&MaxPageSize&"��"
    end if
end if
%>
<table width="100%" cellpadding="2" cellspacing="5" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr2) then %>
    <td valign="top">
        <table width="100%" cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td>Rank</td>
            <td>��ǰ�ڵ�</td>
            <td>�귣��ID</td>
            <td>��ǰ��</td>
            <td>�ֹ��Ǽ�</td>
            <td>��ǰ����</td>
            <td>��ǰ����</td>
            <td>�����Ѿ�</td>
            <td>�������</td>
            <td>���ʽ�����</td>
            <td>�����Ѿ�</td>
            <td>�������II</td>
            <td>�ǸŻ���</td>
            <!--
            <td align="center">- ����� -</td>
            <td>��ǰ�ڵ�</td>
            <td>�귣��ID</td>
            <td>��ǰ��</td>
            <td>�ֹ��Ǽ�</td>
            <td>��ǰ����</td>
            <td>�����Ѿ�</td>
            <td>�������</td>
            <td>�ǸŻ���</td>
            -->
        </tr>
        <% For i = 0 To UBound(vArr2,2) %>
        <%
        vSum1=vSum1+vArr2(1,i)
        vSum2=vSum2+vArr2(2,i)
        vSum3=vSum3+vArr2(3,i)
        vSum4=vSum4+vArr2(4,i)
        
        vSum5=vSum5+vArr2(5,i)
        vSum6=vSum6+vArr2(6,i)
        
        vSum7=vSum7+vArr2(7,i)
        vSum8=vSum8+vArr2(8,i)
        
        isellStr    =""
        iLimitStr   =""
        priceStr    = ""
        
        iSellyn = vArr2(8+4,i)
        iLimityn = vArr2(9+4,i)
        iLimitNo = vArr2(10+4,i)-vArr2(11+4,i)
        if (iLimitNo<1) then iLimitNo=0
            
        
        if (iSellyn<>"Y") then isellStr="<strong><font color='#FF0000'>ǰ��</font></strong>"
        if (iSellyn="S") then isellStr="<strong><font color='#CC3333'>�Ͻ�ǰ��</font></strong>"
        if (iLimityn="Y") then iLimitStr="<font color='#3333CC'>����("&iLimitNo&")</font>"
          
        ''----------------
'        vSum11=vSum11+vArr2(1+13,i)
'        vSum21=vSum21+vArr2(2+13,i)
'        vSum31=vSum31+vArr2(3+13,i)
'        vSum41=vSum41+vArr2(4+13,i)
'        
'        isellStr1    =""
'        iLimitStr1   =""
'        priceStr1    = ""
'        
'        iSellyn1 = vArr2(8+13,i)
'        iLimityn1 = vArr2(9+13,i)
'        iLimitNo1 = vArr2(10+13,i)-vArr2(11+13,i)
'        if (iLimitNo1<1) then iLimitNo1=0
'            
'        
'        if (iSellyn1<>"Y") then isellStr1="<strong><font color='#FF0000'>ǰ��</font></strong>"
'        if (iSellyn1="S") then isellStr1="<strong><font color='#CC3333'>�Ͻ�ǰ��</font></strong>"
'        if (iLimityn1="Y") then iLimitStr1="<font color='#3333CC'>����("&iLimitNo1&")</font>"
              
        %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="center"><%=vArr2(5+4,i)%></td>
            <td align="left"><%=vArr2(0,i)%></td>
            <td align="left"><%=vArr2(7+4,i)%></td>
            <td align="left"><%=vArr2(6+4,i)%></td>
            <td><%=FormatNumber(vArr2(1,i),0)%></td>
            <td><%=FormatNumber(vArr2(2,i),0)%></td>
            <td><%=FormatNumber(vArr2(7,i),0)%></td>
            <td><%=FormatNumber(vArr2(3,i),0)%></td>
            <td><%=FormatNumber(vArr2(4,i),0)%></td>
            <td><%=FormatNumber(vArr2(8,i),0)%></td>
            <td><%=FormatNumber(vArr2(5,i),0)%></td>
            <td><%=FormatNumber(vArr2(6,i),0)%></td>
            <td align="left"><%=isellStr%><%if(iLimitStr<>"")then response.write " "&iLimitStr%></td>
            
        </tr>
        <% next %>
        <tr bgcolor="#F4F4F4" align="right">
            <td align="left">�հ�</td>
            <td align="left"></td>
            <td align="left"></td>
            <td align="left"></td>
            <td><%=FormatNumber(vSum1,0)%></td>
            <td><%=FormatNumber(vSum2,0)%></td>
            <td><%=FormatNumber(vSum7,0)%></td>
            <td><%=FormatNumber(vSum3,0)%></td>
            <td><%=FormatNumber(vSum4,0)%></td>
            <td><%=FormatNumber(vSum8,0)%></td>
            <td><%=FormatNumber(vSum5,0)%></td>
            <td><%=FormatNumber(vSum6,0)%></td>
            <td></td>
        </tr>
        </table>
    </td>
    <% end if %>
	
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->