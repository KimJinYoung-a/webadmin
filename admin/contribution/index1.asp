<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/classes/contribution/contributionCls1.asp"--> 

<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
 
<% 
dim cateList, intC
dim dispCate
    dispCate =  requestCheckvar(request("blnDp"),10) 
    if dispCate = "Y" then
	cateList = fnGetCateList
    end if

Dim i,yyyy2,mm2,yyyy4,mm4, dd2, dd4, fromDate ,toDate,dategbn ,sDt,eDt
	
   	sDt     = requestcheckvar(request("sDt"),10)
	eDt     = requestcheckvar(request("eDt"),10)
    dategbn     = requestCheckvar(request("dategbn"),32)

if sDt ="" then
  yyyy2 = Cstr(Year( dateadd("m",-1,date()) ))
  mm2 = Cstr(Month( dateadd("m",-1,date()) ))
  dd2 = "01"
  sDt = DateSerial(yyyy2, mm2, dd2)
end if
if eDt="" then
  yyyy4 = Cstr(Year( dateadd("m",-1,date()) ))
  mm4 = Cstr(Month( dateadd("m",-1,date()) ))
  dd4 = Cstr(Day( dateadd("d",-1,DateSerial(Year(Date()), Month(Date()), 1)) ))
 eDt = DateSerial(yyyy4, mm4, dd4) 
 end if
if dategbn="" then dategbn="dtlActDate"


%>

 	<link rel="stylesheet" href="/css/datagrid/vertical-layout-light/style.css">
     <link rel="stylesheet" href="/css/reset.css">
 <script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script> 
<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.common.css" />
<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.light.compact.css" />

<!-- plugins:css -->
	<link rel="stylesheet" href="/css/datagrid/vendor/themify-icons.css">
	<link rel="stylesheet" href="/css/datagrid/vendor/vendor.bundle.base.css">
	<link rel="stylesheet" href="/css/datagrid/vendor/vendor.bundle.addons.css">
	<!-- End plugin css for this page -->
	<!-- inject:css -->

	<!-- endinject --> 
    
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.2/jszip.min.js"></script>
<script src="https://cdn3.devexpress.com/jslib/19.1.4/js/dx.all.js"></script>
  <style>
  #gridContainer {
    height: 800px;
}
  </style>
 <!-- 검색 필터 -->
    <div class="card">
        <div class="card-body">
            <h4 class="card-titile" style="padding:10px;border-bottom:1px solid #e4e4e4;">검색 조건</h4>
            <form name="frm" method="get" action="" class="mb-0">
                <div class="form-inline" style="width:500px; display:inline-block;">
                    <div class="form-inline mr-2">
                         <span style="font-size:12px;padding:10px;"> 날짜 : </span>
                        <select class="select" name="dategbn" onChange="jsh4SetYYYYMM4()">
                            <option value="ipkumdate" <%=CHKIIF(dategbn="ipkumdate","selected","")%> >원결제일자</option>
                            <option value="dtlActDate" <%=CHKIIF(dategbn="dtlActDate","selected","")%> >결제일자(처리일자)</option>
                            <option value="chulgoDate" <%=CHKIIF(dategbn="chulgoDate","selected","")%> >출고일자</option>
                            <option value="jFixedDt" <%=CHKIIF(dategbn="jFixedDt","selected","")%> >정산확정일자</option> 
                        </select>
                        &nbsp;
                        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" style="width:100px" placeholder="시작일"  />
                        <img src="http://webadmin.10x10.co.kr/images/admin_calendar.png" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 
                        ~
                        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" style="width:100px" placeholder="종료일" />
                        <img src="http://webadmin.10x10.co.kr/images/admin_calendar.png" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
                        <script language="javascript">
                            var CAL_Start = new Calendar({
                                inputField : "sDt", trigger    : "sDt_trigger",
                                onSelect: function() {
                                    var date = Calendar.intToDate(this.selection.get());
                                    CAL_End.args.min = date;
                                    CAL_End.redraw();
                                    this.hide();
                                }, bottomBar: true, dateFormat: "%Y-%m-%d"
                            });
                            var CAL_End = new Calendar({
                                inputField : "eDt", trigger    : "eDt_trigger",
                                onSelect: function() {
                                    var date = Calendar.intToDate(this.selection.get());
                                    CAL_Start.args.max = date;
                                    CAL_Start.redraw();
                                    this.hide();
                                }, bottomBar: true, dateFormat: "%Y-%m-%d"
                            });
                        </script> 
                    </div>
                </div>
                <div style="display:inline-block;">
                    <button type="button"  class="button_s" style="width:100px;height:30px;" onclick="SearchForm(this.form);">검색</button>
                </div>
            </form>
        </div>
        <div style="padding: 10px 0;border-bottom:1px solid #e4e4e4;width:100%" ></div>
     </div>
               
    <div class="col-md-12 grid-margin stretch-card" style="padding-top:10px">
        <div class="card">
            <div class="card-body" id="dxGridBody" >
                <!-- 데이터 그리드 시작 -->
                <div class="dx-viewport"  style="padding-top:10px">
                    <div id="gridContainer"></div>
                </div>
                <!-- 데이터 그리드 끗 -->
                </div>
            </div>
        </div>
    </div>
    <!-- 메인 끝 -->
     
<script type="text/javascript">
$(function(){ 
    var param = $(document.frm).serialize();
    
    callDataGrid(param);
});

// 데이터 그리드 호출
function callDataGrid(param) {
   var sourceUrl = "/admin/contribution/json_List1.asp"; 
    if(param) {
        sourceUrl = "/admin/contribution/json_List1.asp?" + param;
    } 
    $("#gridContainer").dxDataGrid({
        showColumnLines: true, // 컬럼 라인
        showRowLines: true, // 로우 라인
        rowAlternationEnabled: false, // 로우별 회색 색상
        showBorders: true, // 전체 보더
      
         //그룹핑
           grouping: {
            autoExpandAll: true,
        },

        groupPanel: {
            visible: true
        },
        paging: {
            pageSize: 30
        },
        pager: {
            showPageSizeSelector: true,
            allowedPageSizes: [30, 50, 100]
        },
        columnChooser: { // 화면에 보여주는 컬럼 선택
            enabled: true,
            mode: "select" // or "dragAndDrop"
        },
        searchPanel: {
            visible: true,
            width: 240,
            placeholder: "키워드 검색"
        },
     
        "export": { // 엑셀 다운로드 관련
            enabled: true,
            fileName: "EmailCustomerList",
            allowExportSelectedData: true
        },
        headerFilter: { // 컬럼명 깔대기 검색 
            visible: false
        },
        columnAutoWidth: true,
        columns: [   
            {dataField:"구분1", dataType:"string"},  
            {dataField:"구분2", dataType:"string"},  
            {dataField:"날짜", dataType:"string", alignment: "center" },
            {
                caption: "전체",
                columns: [{
                    caption: "금액",
                    dataField: "전체_금액",
                    alignment: "right", 
                    dataType:"number",
                    format:  "fixedPoint"     
                }, {
                    caption: "율",
                    dataField: "전체_율",
                    alignment: "right",
                   format: {
    					type: "percent",
    					precision: 1
    				}
                }]
            },

            <%IF isArray(cateList) THEN %>
              {
                caption: "NULL",
                columns: [{
                    caption: "금액",
                    dataField: "NULL_금액",
                    dataType:"number",
                    format:  "fixedPoint"    
                }, {
                    caption: "율",
                    dataField: "NULL_율",
                    valueFormat: "percent"
                }]
            }, 
            <%
                For intC = 0 To UBound(cateList,2)
            %> 
             {
                caption: "<%=cateList(1,intC)%>",
                columns: [{
                    caption: "금액",
                    dataField: "<%=cateList(1,intC)%>_금액",
                     dataType:"number",
                    format:  "fixedPoint"    
                }, {
                    caption: "율",
                    dataField: "<%=cateList(1,intC)%>_율",
                    valueFormat: "percent"
                }]
            }, 
            <%  Next
               END IF%> 
             { 
                dataField: "매출구분", 
                groupIndex: 0
            }, 
             { 
                dataField: "제휴구분", 
                groupIndex: 1
            }, 
             
        ],
 
         summary: {
            groupItems: [  {   
                column: "전체_금액",
                summaryType: "sum",
                valueFormat: "fixedPoint",
                displayFormat: "Total: {0}",
                showInGroupFooter: false,
                alignByColumn: true
            }, {
                column: "전체_율",
                summaryType: "sum",
                valueFormat: "percent",
                showInGroupFooter: false,
                alignByColumn: true 
             }]
        },
        
        dataSource : sourceUrl
    });
}


// 검색 실행
function SearchForm(frm) {
    frm.submit();
} 

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->