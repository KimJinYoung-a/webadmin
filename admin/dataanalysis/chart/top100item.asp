<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/report/simplereportcls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim vReportType : vReportType = "top100item"
Dim oSimpleReport, vArrData, vArrCols, i, j

Dim vSDate, vEDate, vChannel, vOrdType, chkvs
Dim vDategbn, addparam1

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vDategbn = requestCheckvar(request("dategbn"),10)
vChannel = requestCheckvar(request("channel"),10)
vOrdType = requestCheckvar(request("ordtype"),32)
addparam1= requestCheckvar(request("addparam1"),32)
chkvs= requestCheckvar(request("chkvs"),32)

if (vOrdType="") then vOrdType="S" ''건수(C) , 금액(S), 수익(G)


If vSDate = "" Then
    vSDate = LEFT(dateadd("d",-2,now()),10)
	vEDate = LEFT(dateadd("d",-1,now()),10)
End If


SET oSimpleReport = new CSimpleReport
	oSimpleReport.FRectSDate = vSDate
	oSimpleReport.FRectEDate = LEFT(DateAdd("d",1,vEDate),10)
	oSimpleReport.FRectDateGbn = vDategbn
	oSimpleReport.FRectReportType = vReportType
	oSimpleReport.FRectChannel = vChannel
	oSimpleReport.FRectAddParam1 = addparam1
	oSimpleReport.FPageSize = 100
	oSimpleReport.FRectOrderType = vOrdType
	oSimpleReport.FRectchkvs =chkvs

	vArrData = oSimpleReport.getSimpleReportStatistics(vArrCols)

SET oSimpleReport = nothing

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
	document.frm1.submit();
}
function popOutItem(iitemid, s1, s2, s3, e1, e2, e3){
	var popwin=window.open('/admin/upchejungsan/upcheselllist.asp?page=1&menupos=138&designer=&itemid='+iitemid+'&datetype=ipkumil&yyyy1='+s1+'&mm1='+s2+'&dd1='+s3+'&yyyy2='+e1+'&mm2='+e2+'&dd2='+e3+'&delivertype=all&purchasetype=&sitename=&disp1=&disp=&sellchnl=OUT','notin','width=900,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>


<body>
<form name="frm1" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
	결제일
     :
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;
    <% if (FALSE) then %>
    채널 : <% call drawConversionChannelSelectBox("channel",vChannel) %>
    &nbsp;&nbsp;
    <% end if %>
    정렬 :
    <label><input type="radio" name="channel" value="" <%=CHKIIF(vChannel="","checked","") %> ><span>전체</span></label>
    <label><input type="radio" name="channel" value="ten" <%=CHKIIF(vChannel="ten","checked","") %> ><span>자사몰</span></label>
    <label><input type="radio" name="channel" value="tenv" <%=CHKIIF(vChannel="tenv","checked","") %> ><span>자사몰(nv제외)</span></label>
    <label><input type="radio" name="channel" value="nv" <%=CHKIIF(vChannel="nv","checked","") %> ><span>NaverEP</span></label>
    <label><input type="radio" name="channel" value="out" <%=CHKIIF(vChannel="out","checked","") %> ><span>제휴몰</span></label>
    &nbsp;&nbsp;

    정렬 :
    <label><input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> ><span>주문건수순</span></label>
    <label><input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> ><span>구매총액순</span></label>
    <label><input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> ><span>매출수익순</span></label>
    &nbsp;&nbsp;

    <label><input type="checkbox" name="chkvs" <%= CHKIIF(chkvs<>"","checked","")%> >비교</label>
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<div>
    <p style="float:left;"><% Call drawReportDescription(vReportType) %></p>
    <p style="float:right; padding-right:5px;"><img id="btnExport" src="/images/btn_excel.gif" style="cursor:pointer;" /></p>
</div>
<%
dim fld, vArr, rows, cols, col_name, col_wid, col_fmt, col_align, colsplited
%>
<table cellpadding="3" cellspacing="2" border="0" class="a" align="center" width="100%">
<tr bgcolor="#FFFFFF">
    <td width="500">

        <% If isArray(vArrCols) Then %>
        <table cellpadding="5" cellspacing="1" class="a" align="center" bgcolor="#999999" id="tblExport">
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
                    <% if isNULL(vArrData(cols,i)) then %>

                    <% else %>
                    <%=FormatNumber(vArrData(cols,i),0)%>
                    <% end if %>
                <% else %>
                    <% 'vArrData(cols,i)%>
                    <%
                        If (cols = 0) AND vChannel = "out" Then
                            Response.Write "<span style='cursor:pointer;' onclick=""popOutItem('"& vArrData(cols,i) &"', '"& Split(vSDate, "-")(0) &"', '"& Split(vSDate, "-")(1) &"', '"& Split(vSDate, "-")(2) &"', '"& Split(vEDate, "-")(0) &"', '"& Split(vEDate, "-")(1) &"', '"& Split(vEDate, "-")(2) &"')"">"&vArrData(cols,i)&"</span>"
                        Else
                            response.write vArrData(cols,i)
                        End If
                    %>
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
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<script>
$(document).ready(function(){
    // 최소 2자리 문자열 변환
    function itoStr($num) {
        $num < 10 ? $num = '0'+$num : $num;
        return $num.toString();
    }

    // 선택 테이블 파일로 다운받기
    function exportTableToExcel(tableID, filename) {
        var downloadLink;
        var dataType = 'application/vnd.ms-excel';
        var tableSelect = document.getElementById(tableID);
        var tableHTML = tableSelect.outerHTML;

        // Specify file name
        filename = filename?filename+'.xls':'excel_data.xls';

        // Create download link element
        downloadLink = document.createElement("a");

        document.body.appendChild(downloadLink);

        if(navigator.msSaveOrOpenBlob){
            var blob = new Blob([tableHTML.replace(/<br>/g, '')], {
                type: dataType
            });
            navigator.msSaveOrOpenBlob( blob, filename);
        }else{
            // Create a link to the file
            downloadLink.href = 'data:' + dataType + ', ' + tableHTML.replace(/ /g, '%20').replace(/#/g, '%23').replace(/<br>/g, '');

            // Setting the file name
            downloadLink.download = filename;

            //triggering the function
            downloadLink.click();
        }
    }

    var btn = $('#btnExport'); // 버튼ID
    var tbl = 'tblExport';  // 테이블ID

    // 버튼액션 매핑
    btn.click(function(e){
        var chnNm = $('input:radio[name="channel"]:checked').next().text();
        var ordNm = $('input:radio[name="ordtype"]:checked').next().text();

        var dt = new Date();
        var year =  itoStr( dt.getFullYear() );
        var month = itoStr( dt.getMonth() + 1 );
        var day =   itoStr( dt.getDate() );
        var hour =  itoStr( dt.getHours() );
        var mins =  itoStr( dt.getMinutes() );

        var postfix = year + month + day + "_" + hour + "_" + mins;
        var fileName = "Top100상품_"+chnNm+"_"+ordNm+"_"+postfix;

        exportTableToExcel(tbl, fileName);

        e.preventDefault();
    });
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->