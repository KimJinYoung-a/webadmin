<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/report/simpleQryCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim menuGrpIdx : menuGrpIdx = fnGetMenuGrpIDx()
Dim qryidx : qryidx = requestCheckvar(request("qryidx"),10)
Dim oSimpleQuery, i, j
Dim MaxparamCNT : MaxparamCNT = 10
Dim iparamboxtype, iparamname, iparamBoxhtml, iparamval, iparammaxlen, iparamidx
Dim bparam1, bparam2

' ReDim paramArr(MaxparamCNT)
' for i=0 to MaxparamCNT-1
'     paramArr(i) = requestCheckvar(request("param"&(i+1)),32)
' Next
dim page : page = requestCheckvar(request("page"),10)
if (page="") then page=1

Dim oQryParam , retExplain, vArrCols, vArrData, retReturnValue
Dim retVal
if (qryidx<>"") then
    SET oQryParam = new CSimpleQuery
    oQryParam.FRectQryidx = qryidx
    oQryParam.getQueryParamArr

    for i=0 to oQryParam.FParamCount-1
        iparamname = oQryParam.FQryParamList(i).Fparamname
        oQryParam.FQryParamList(i).FStoredparamVal = oQryParam.FQryParamList(i).getRequestParam ''requestCheckVar(request(iparamname),oQryParam.FQryParamList(i).Fparamlength)

        if (oQryParam.FQryParamList(i).FStoredparamVal="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
            oQryParam.FQryParamList(i).FStoredparamVal = oQryParam.FQryParamList(i).getDefaultVAl
        end if
    next

    retVal = oQryParam.ExecSimpleQuery(retExplain, vArrCols, vArrData, retReturnValue)

end if



''-----------------------------------------
dim vReportType
Dim vSDate, vEDate, vChannel, vOrdType
Dim vDategbn, addparam1, addparam2, addOpt, addOptSub

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vDategbn = requestCheckvar(request("dategbn"),10)
vChannel = requestCheckvar(request("channel"),10)
vOrdType = requestCheckvar(request("ordtype"),32)
addparam1= requestCheckvar(request("addparam1"),32)
addparam2= requestCheckvar(request("cpngbn"),32)

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

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<style type="text/css">
.dx-widget {font-size:12px;}
</style>
<body>
<form name="frm1" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
    쿼리선택 : <% call drawSimpleQuerySelectBox("qryidx",qryidx,menuGrpIdx) %>
    &nbsp;&nbsp;



    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
        <div id="lyrProcess" style="position:absolute; background-color:#eee; width:50px;height:22px;display:none;"><img src="/images/loading.gif" style="withd:20px;height:20px;" /></div>
        <div>
		    <input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
        </div>
	</td>
</tr>
<tr align="center" bgcolor="#F4F4F4">
    <td align="left">
    <% if (qryidx<>"") then %>
    <% for i=0 to oQryParam.FParamCount-1 %>
        <%
        iparamboxtype = oQryParam.FQryParamList(i).Fparamboxtype
        iparamname    = oQryParam.FQryParamList(i).Fparamname
        iparammaxlen  = oQryParam.FQryParamList(i).Fparamlength
        iparamidx     = oQryParam.FQryParamList(i).Fparamidx
        iparamBoxhtml = ""
        iparamval     = ""
        bparam1       = ""
        bparam2       = ""
        SELECT CASE iparamboxtype
            CASE "yyyymmdd" :
                iparamval  = oQryParam.FQryParamList(i).getRequestParam 'requestCheckVar(request(iparamname),iparammaxlen)
                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).getDefaultVAL ''LEFT(dateadd("d",oQryParam.FQryParamList(i).Fdefaultval,now()),10)
                end if
                iparamBoxhtml= oQryParam.FQryParamList(i).FparamTitle & " : <input id='"&iparamname&"' name='"&iparamname&"' value='"&iparamval&"' class='text' size='10' maxlength='10' />"
                iparamBoxhtml = iparamBoxhtml & "<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='"&iparamname&"_trigger' border='0' style='cursor:pointer' align='absmiddle' />"
            CASE "yyyymm" :
                iparamval = oQryParam.FQryParamList(i).getRequestParam

                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).getDefaultVAL ''LEFT(dateadd("m",oQryParam.FQryParamList(i).Fdefaultval,now()),7)
                end if
                bparam1 = LEFT(iparamval,4)
                bparam2 = RIGHT(iparamval,2)

                iparamBoxhtml = ""
                response.write oQryParam.FQryParamList(i).FparamTitle & " : "
                DrawYMBoxIdx bparam1,bparam2,iparamidx
                response.write "&nbsp;&nbsp;"

            CASE "box" :
                iparamval  = oQryParam.FQryParamList(i).getRequestParam ''requestCheckVar(request(iparamname),iparammaxlen)

                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).Fdefaultval
                end if

                iparamBoxhtml = oQryParam.FQryParamList(i).FparamTitle & " : <input type='text' name='"&iparamname&"' value='"&iparamval&"' size='"&iparammaxlen&"' maxlength='"&iparammaxlen&"'>"

            CASE "select" :
                iparamval  = oQryParam.FQryParamList(i).getRequestParam

                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).Fdefaultval
                end if

                addOpt = split(oQryParam.FQryParamList(i).FparamSelectOpt,",")
                if isArray(addOpt) then
                    iparamBoxhtml = oQryParam.FQryParamList(i).FparamTitle & " : "
                    iparamBoxhtml = iparamBoxhtml & "<select name="""&iparamname&""" class=""select"">"

                    For j=0 to ubound(addOpt)
                        addOptSub = split(addOpt(j),"|")
                        iparamBoxhtml = iparamBoxhtml & "<option value=""" & addOptSub(1) & """ " & chkIIF(addOptSub(1)=iparamval,"selected","") & ">" & addOptSub(0) & "</option>"
                    Next

                    iparamBoxhtml = iparamBoxhtml & "</select"
                end if

            CASE else
                iparamBoxhtml = "?"&iparamboxtype&"?"

        END SELECT

        if (iparamBoxhtml<>"") then iparamBoxhtml=iparamBoxhtml&"&nbsp;&nbsp;"
        response.write iparamBoxhtml
        %>

    <% next %>
    <% end if %>
    </td>
</tr>

</table>
</form>
<!-- 설명필드 -->
<%=chkIIF(retExplain<>"","<p style=""text-align:right;"">※ "&retExplain&"</p>","")%>

<!-- 데이터 그리드 시작 -->
<div class="dx-viewport">
    <div class="demo-container">
        <div id="gridContainer"><center>No Data...</center></div>
    </div>
</div>
<!-- 데이터 그리드 끗 -->
<%
dim fld, vArr, rows, cols, col_name, col_wid, col_fmt, col_align, colsplited, col_data
%>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.common.css" />
<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.light.compact.css" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.2/jszip.min.js"></script>
<script src="https://cdn3.devexpress.com/jslib/19.1.4/js/dx.all.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" crossorigin="anonymous" referrerpolicy="no-referrer" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<style type="text/css">
	.select2-container .select2-selection--single {height:17px;}
	.select2-container--default .select2-selection--single .select2-selection__rendered {line-height:16px;}
	.select2-container--default .select2-selection--single .select2-selection__arrow {height: 15px;}
</style>
<script>
$(function() {
    <%
    if (qryidx<>"") then
    for i=0 to oQryParam.FParamCount-1
        iparamboxtype = oQryParam.FQryParamList(i).Fparamboxtype
        iparamname = oQryParam.FQryParamList(i).Fparamname
        iparamBoxhtml = ""
        SELECT CASE iparamboxtype
            CASE "yyyymmdd" :
                iparamBoxhtml = "var CAL_"&iparamname&" = new Calendar({"
                iparamBoxhtml = iparamBoxhtml & "inputField : '"&iparamname&"', trigger    : '"&iparamname&"_trigger',"
		        iparamBoxhtml = iparamBoxhtml & " onSelect: function() {"
			    iparamBoxhtml = iparamBoxhtml & "   var date = Calendar.intToDate(this.selection.get());"
			    'iparamBoxhtml = iparamBoxhtml & "   CAL_End.args.min = date;"
			    'iparamBoxhtml = iparamBoxhtml & "   CAL_End.redraw();"
			    iparamBoxhtml = iparamBoxhtml & "   this.hide();"
		        iparamBoxhtml = iparamBoxhtml & "   }, bottomBar: true, dateFormat: '%Y-%m-%d'"
	            iparamBoxhtml = iparamBoxhtml & "});"

            CASE else
                iparamBoxhtml =""
        END SELECT
        response.write iparamBoxhtml&vbCRLF

    Next
    end if
	%>
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
	$("#lyrProcess").show();
    document.frm1.submit();
}

$(function(){
    $("#gridContainer").dxDataGrid({
        showColumnLines: true, // 컬럼 라인
        showRowLines: true, // 로우 라인
        paging: {
            pageSize: 50
        },
        rowAlternationEnabled: true, // 로우별 회색 색상
        showBorders: true, // 전체 보더
        columnChooser: { // 화면에 보여주는 컬럼 선택
            enabled: true,
            mode: "select" // or "dragAndDrop"
        },
        "export": { // 엑셀 다운로드 관련
            enabled: true,
            fileName: "EmailCustomerList",
            allowExportSelectedData: true
        },
        headerFilter: { // 컬럼명 깔대기 검색 
            visible: true
        },
        columnAutoWidth: true,
        columns: [
        <%
            If isArray(vArrCols) Then
                '// 헤더출력
                For cols = 0 To UBound(vArrCols)
                    col_fmt  = "S"
                    col_align = ""
                    col_wid  = ""
                    colsplited = split(vArrCols(cols),"|")
                    if isArray(colsplited) then
                        col_name = colsplited(0)
                        if UBOUND(colsplited)>0 then col_fmt = colsplited(1)
                        if UBOUND(colsplited)>1 then col_align = colsplited(2)
                        if UBOUND(colsplited)>2 then col_wid = colsplited(3)
                    else
                        col_name = colsplited
                    end if

                    Response.Write "{"
                    Response.Write "dataField : """ & col_name & ""","
                    Response.Write "alignment : """ & CHKIIF(col_align="L","left",CHKIIF(col_align="R","right","center"))  & """, "
                    Response.Write "dataType: """ & CHKIIF(col_fmt="N","number","string") & ""","
                    if col_fmt="N" then Response.Write "format: ""fixedPoint"","
                    Response.Write "fixed: true"
                    Response.Write "},"
                Next
            end if
         %>
        ],
        <%
            If isArray(vArrData) Then
                '// 데이터출력
                Response.Write "dataSource : ["
                For i = 0 To UBound(vArrData,2)
                    Response.Write "{"
                    for cols=0 To UBound(vArrCols)
                        colsplited = split(vArrCols(cols),"|")
                        col_fmt  = "S"
                        if isArray(colsplited) then
                            col_name = colsplited(0)
                            if UBOUND(colsplited)>0 then col_fmt = colsplited(1)
                        else
                            col_name = colsplited
                            col_fmt  = "S"
                        end if

                        if (col_fmt="N") and (vArrData(cols,i)="" or isnull(vArrData(cols,i))) then
                            col_data = 0
                        elseif (vArrData(cols,i)="" or isnull(vArrData(cols,i))) then
                            col_data = ""
                        else
                            col_data = replace(vArrData(cols,i),"""","\""")
                        end if

                        Response.Write """" & col_name & """:"
                        Response.Write """" & col_data & ""","

                    Next
                    Response.Write "},"

                    if (i mod 500)=499 then
                        response.flush
                    end if

                Next
                Response.Write "]"
            end if
         %>
    });

    //쿼리선택
    $("select[name=qryidx]").select2();
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
