<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ְ��ڷ�ROW
' History : 2018.03.23 ������ ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/report/simplereportcls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim vReportType : vReportType = requestCheckvar(request("reporttype"),32)
Dim oSimpleReport, vArrData, vArrCols, i, j

Dim vSDate, vEDate, vChannel, vOrdType
Dim vDategbn, addparam1, addparam2, itemid

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vDategbn = requestCheckvar(request("dategbn"),10)
vChannel = requestCheckvar(request("channel"),10)
vOrdType = requestCheckvar(request("ordtype"),32)
addparam1= requestCheckvar(request("addparam1"),32)
addparam2= requestCheckvar(request("addparam2"),32)
itemid = requestCheckvar(request("itemid"),10)

if (vOrdType="") then vOrdType="S" ''�Ǽ�(C) , �ݾ�(S), ����(G)

'' �⺻�� '' �����ִ� datepart("ww",now()-1day) '' �츮�� ��~�Ͽ��ϱ����� ���ַ� �Ѵ�
dim defaultWW : defaultWW=DatePart("ww",dateadd("d",-1,now()))

'' �̹��� 
dim thisMon : thisMon = LEFT(dateadd("d",DatePart("w",now())*-1+2,now()),10)
dim thisSun : thisSun = dateadd("d",6,thisMon)

'' ���ÿ���
dim thisW : thisW = datepart("w",now())

if (vDategbn="") then vDategbn="O" ''�ֹ���

If vSDate = "" Then
    if (thisW=2 or thisW=3) then  ''��, ȭ������ ������ ����
        vSDate = dateadd("d",-7,thisMon)
	    vEDate = dateadd("d",-7,thisSun)
    else
	    vSDate = thisMon
	    vEDate = LEFT(date(),10)
    end if
End If

SET oSimpleReport = new CSimpleReport
	oSimpleReport.FRectSDate = vSDate
	oSimpleReport.FRectEDate = LEFT(DateAdd("d",1,vEDate),10)
	oSimpleReport.FRectDateGbn = vDategbn
	oSimpleReport.FRectReportType = vReportType
	oSimpleReport.FRectChannel = vChannel
	oSimpleReport.FRectAddParam1 = addparam1
	oSimpleReport.FRectAddParam2 = addparam2
	oSimpleReport.FPageSize = 30
	oSimpleReport.FRectOrderType = vOrdType
	oSimpleReport.FRectitemid = itemid
	vArrData = oSimpleReport.getSimpleReport(vArrCols)
SET oSimpleReport = nothing

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type="text/javascript">

$(function() {
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
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
	
    $("#btnSubmit").hide();
    $("#imgSubmit").fadeIn();
    document.frm1.submit();
}

function calcuDt(tp){
    var stval='';
    var edval='';
    
    if (tp=='tw'){ //�̹���
        stval="<%=thisMon%>";
        edval="<%=thisSun%>";
    }else if (tp=='pw'){ //������
        stval="<%=dateadd("d",-7,thisMon)%>";
        edval="<%=dateadd("d",-7,thisSun)%>";
    }else if (tp=='tpm'){ //�̹���-4Week
        stval="<%=dateadd("d",-7*4,thisMon)%>";
        edval="<%=dateadd("d",-7*4,thisSun)%>";    
    }else if (tp=='tpy'){ //�̹���-1Year
        stval="<%=dateadd("d",-7*52,thisMon)%>";
        edval="<%=dateadd("d",-7*52,thisSun)%>";    
    }else if (tp=='ppm'){ //������-4Week
        stval="<%=dateadd("d",-7*4-7,thisMon)%>";
        edval="<%=dateadd("d",-7*4-7,thisSun)%>";    
    }else if (tp=='ppy'){ //������-1Year
        stval="<%=dateadd("d",-7*52-7,thisMon)%>";
        edval="<%=dateadd("d",-7*52-7,thisSun)%>";    
    }else{
        
    }
    
    document.frm1.startdate.value=stval;
    document.frm1.enddate.value=edval;
}
</script>

<body>
<form name="frm1" method="get" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	<select name="dategbn">
	<option value='O' <%=CHKIIF(vDategbn="O","selected","")%> >�ֹ���</option>
	<% if (vReportType<>"dealsales") then %>
	<option value='P' <%=CHKIIF(vDategbn="P","selected","")%> >������</option>
    <% end if %>
	</select>
     : 
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;
    
    <input type="button" value="������(<%=defaultWW-1%>��)" onclick="calcuDt('pw')">
    <input type="button" value="������-1Year" onclick="calcuDt('ppy')">
    <input type="button" value="������-4Week" onclick="calcuDt('ppm')">
    
    <input type="button" value="�̹���(<%=defaultWW%>��)" onclick="calcuDt('tw')">
    <input type="button" value="�̹���-1Year" onclick="calcuDt('tpy')">
    <input type="button" value="�̹���-4Week" onclick="calcuDt('tpm')">
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
        <img id="imgSubmit" src="/images/loading.gif" style="width:45px; display:none;" />
	</td>
</tr>
<tr align="center" bgcolor="#F4F4F4">
    <td align="left">
    Report : <% call drawReportSelectBox("reporttype",vReportType) %>
    &nbsp;&nbsp;
    <%
    if (vReportType="bestitemcoupon") or (vReportType="salesitemcpnbyuserlevel") or (vReportType="itemcpnevalwithsales") then 
    %>
        ��ǰ������ȣ : <input type="text" name="addparam1" value="<%=addparam1%>" size="20" maxlength="32">
    <%
    end if
    %>
    <%
    if (vReportType="outmallsales") or (vReportType="aaaaaa") or (vReportType="ssssssss") then 
    %>
        ���޸�ID : <input type="text" name="addparam1" value="<%=addparam1%>" size="20" maxlength="32">
    <%
    end if
    %>
    <%
    if (vReportType="evtsubscript") or (vReportType="aaaaaa") or (vReportType="ssssssss") then 
    %>
        �̺�Ʈ�ڵ� : <input type="text" name="addparam1" value="<%=addparam1%>" size="20" maxlength="32">
    <%
    end if
    %>
    <%
    if (vReportType="dealsales") then 
    %>
        ���ڵ� : <input type="text" name="addparam1" value="<%=addparam1%>" size="10" maxlength="10">
        ����ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size="10" maxlength="10">
    <%
    end if
    %>
    ä�� : <% call drawConversionChannelSelectBox("channel",vChannel) %>
    &nbsp;&nbsp;   
    ���� : 
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >�ֹ��Ǽ���
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >�����Ѿ׼�
    <!-- input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >������ͼ� -->
    &nbsp;&nbsp;    
    <% if (vReportType="bestitemcoupon") then %>
    ���� ���� :
    <select name="addparam2">
        <option value="" <%=CHKIIF(addparam2="","selected","")%> >��ü</option>
	    <option value="V" <%=CHKIIF(addparam2="V","selected","")%> >���̹�</option>
	    <option value="C" <%=CHKIIF(addparam2="C","selected","")%> >�Ϲ�</option>
	</select>
    <% elseif (vReportType="newitembybrandcate") then %>
    ǥ�� ���� :
    <select name="addparam1">
	    <option value="" <%=CHKIIF(addparam1="","selected","")%> >���</option>
	    <option value="B" <%=CHKIIF(addparam1="B","selected","")%> >�귣���</option>
	</select>
    <% end if %>
    </td>
</tr>

</table>
</form>
<p>
    <% Call drawReportDescription(vReportType) %>
</p>
<% 
dim fld, vArr, rows, cols, col_name, col_wid, col_fmt, col_align, colsplited
%>
<table cellpadding="3" cellspacing="2" border="0" class="a" align="center" width="100%">
<tr bgcolor="#FFFFFF">
    <td width="500">
        
        <% If isArray(vArrCols) Then %>
        <table cellpadding="5" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4" align="center">
        <% For cols = 0 To UBound(vArrCols) %>
            <% 
                colsplited = split(vArrCols(cols),"|") 
                if isArray(colsplited) then
                    col_name = colsplited(0)
                else
                    col_name = colsplited
                end if
            %>
            <td >
                <%=col_name%>
            </td>
        <% Next %>
        </tr>
        <% end if %>
        <% if isArray(vArrData) then %>
        <% For i = 0 To UBound(vArrData,2) %>
        <tr bgcolor="#FFFFFF" align="center">
            <% for cols=0 To UBound(vArrCols) %>
            <%
                colsplited = split(vArrCols(cols),"|") 
                col_fmt = ""
                col_align = ""
                col_wid = ""
                if isArray(colsplited) then
                    if UBOUND(colsplited)>0 then col_fmt = colsplited(1)
                    if UBOUND(colsplited)>1 then col_align = colsplited(2)    
                    if UBOUND(colsplited)>2 then col_wid = colsplited(3) 
                else
                    col_fmt  = "S"
                    col_align = ""
                    col_wid  = ""
                end if
            %>
            <td <%= CHKIIF(col_wid<>"","width='"&col_wid&"'","") %> <%= CHKIIF(col_align="R","align='right'","") %> >
                <% if (col_fmt="N") then %>
                    <% if vArrData(cols,i)="" or isnull(vArrData(cols,i)) then %>
                        0
                    <% else %>
                        <%=FormatNumber(vArrData(cols,i),0)%>
                    <% end if %>
                <% else %>
                    <%=vArrData(cols,i)%>
                <% end if %>
            </td>
            <% next %>
        </tr>
        <% next %>
        </table>
        <% else %>
        No data
        <% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->