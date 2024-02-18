<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주간자료ROW
' History : 2018.03.23 서동석 생성
'			2016.06.30 한용민 수정
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

if (vOrdType="") then vOrdType="S" ''건수(C) , 금액(S), 수익(G)

'' 기본값 '' 현재주는 datepart("ww",now()-1day) '' 우리는 월~일요일까지를 한주로 한다
dim defaultWW : defaultWW=DatePart("ww",dateadd("d",-1,now()))

'' 이번주 
dim thisMon : thisMon = LEFT(dateadd("d",DatePart("w",now())*-1+2,now()),10)
dim thisSun : thisSun = dateadd("d",6,thisMon)

'' 오늘요일
dim thisW : thisW = datepart("w",now())

if (vDategbn="") then vDategbn="O" ''주문일

If vSDate = "" Then
    if (thisW=2 or thisW=3) then  ''월, 화요일은 일주일 전값
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
		alert("시작일을 입력하세요");	
		return false;
	}
	if($("#edate").val()== ""){
		alert("종료일을 입력하세요");	
		return false;
	}
	
    $("#btnSubmit").hide();
    $("#imgSubmit").fadeIn();
    document.frm1.submit();
}

function calcuDt(tp){
    var stval='';
    var edval='';
    
    if (tp=='tw'){ //이번주
        stval="<%=thisMon%>";
        edval="<%=thisSun%>";
    }else if (tp=='pw'){ //지난주
        stval="<%=dateadd("d",-7,thisMon)%>";
        edval="<%=dateadd("d",-7,thisSun)%>";
    }else if (tp=='tpm'){ //이번주-4Week
        stval="<%=dateadd("d",-7*4,thisMon)%>";
        edval="<%=dateadd("d",-7*4,thisSun)%>";    
    }else if (tp=='tpy'){ //이번주-1Year
        stval="<%=dateadd("d",-7*52,thisMon)%>";
        edval="<%=dateadd("d",-7*52,thisSun)%>";    
    }else if (tp=='ppm'){ //지난주-4Week
        stval="<%=dateadd("d",-7*4-7,thisMon)%>";
        edval="<%=dateadd("d",-7*4-7,thisSun)%>";    
    }else if (tp=='ppy'){ //지난주-1Year
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
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
	<select name="dategbn">
	<option value='O' <%=CHKIIF(vDategbn="O","selected","")%> >주문일</option>
	<% if (vReportType<>"dealsales") then %>
	<option value='P' <%=CHKIIF(vDategbn="P","selected","")%> >결제일</option>
    <% end if %>
	</select>
     : 
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;
    
    <input type="button" value="지난주(<%=defaultWW-1%>주)" onclick="calcuDt('pw')">
    <input type="button" value="지난주-1Year" onclick="calcuDt('ppy')">
    <input type="button" value="지난주-4Week" onclick="calcuDt('ppm')">
    
    <input type="button" value="이번주(<%=defaultWW%>주)" onclick="calcuDt('tw')">
    <input type="button" value="이번주-1Year" onclick="calcuDt('tpy')">
    <input type="button" value="이번주-4Week" onclick="calcuDt('tpm')">
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" id="btnSubmit" class="button_s" value="검색" onClick="goSearch(document.frm1);">
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
        상품쿠폰번호 : <input type="text" name="addparam1" value="<%=addparam1%>" size="20" maxlength="32">
    <%
    end if
    %>
    <%
    if (vReportType="outmallsales") or (vReportType="aaaaaa") or (vReportType="ssssssss") then 
    %>
        제휴몰ID : <input type="text" name="addparam1" value="<%=addparam1%>" size="20" maxlength="32">
    <%
    end if
    %>
    <%
    if (vReportType="evtsubscript") or (vReportType="aaaaaa") or (vReportType="ssssssss") then 
    %>
        이벤트코드 : <input type="text" name="addparam1" value="<%=addparam1%>" size="20" maxlength="32">
    <%
    end if
    %>
    <%
    if (vReportType="dealsales") then 
    %>
        딜코드 : <input type="text" name="addparam1" value="<%=addparam1%>" size="10" maxlength="10">
        딜상품코드 : <input type="text" name="itemid" value="<%= itemid %>" size="10" maxlength="10">
    <%
    end if
    %>
    채널 : <% call drawConversionChannelSelectBox("channel",vChannel) %>
    &nbsp;&nbsp;   
    정렬 : 
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >주문건수순
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >구매총액순
    <!-- input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >매출수익순 -->
    &nbsp;&nbsp;    
    <% if (vReportType="bestitemcoupon") then %>
    쿠폰 구분 :
    <select name="addparam2">
        <option value="" <%=CHKIIF(addparam2="","selected","")%> >전체</option>
	    <option value="V" <%=CHKIIF(addparam2="V","selected","")%> >네이버</option>
	    <option value="C" <%=CHKIIF(addparam2="C","selected","")%> >일반</option>
	</select>
    <% elseif (vReportType="newitembybrandcate") then %>
    표시 구분 :
    <select name="addparam1">
	    <option value="" <%=CHKIIF(addparam1="","selected","")%> >요약</option>
	    <option value="B" <%=CHKIIF(addparam1="B","selected","")%> >브랜드상세</option>
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