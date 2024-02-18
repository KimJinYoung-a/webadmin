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
<!-- #include virtual="/lib/classes/contribution/contributionCls.asp"--> 

<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
 
<% 
dim clsCMeachulLog
dim cateList, intC
dim dispCate,catekind
Dim i,syear,smonth, sday, eday,dategbn,dstdate,deddate 

    dispCate =  requestCheckvar(request("blnDp"),10) 
    catekind =  requestCheckvar(request("rdoCK"),1) 
 
    dategbn     = requestCheckvar(request("dategbn"),32)
   	syear     = requestcheckvar(request("sy"),4)
	smonth     = requestcheckvar(request("sM"),2)
    sday     = requestcheckvar(request("sd"),2)
    eday     = requestCheckvar(request("ed"),2) 

if dispCate ="" then dispCate = "N" 
if catekind ="" then catekind = "D"
if dategbn="" then dategbn="ipkumdate"
if syear ="" then  syear = Cstr(Year( dateadd("m",-1,date()) ))
if smonth ="" then smonth = Cstr(Month( dateadd("m",-1,date()) ))
if sday ="" then sday =  "01" 
  dstdate = DateSerial(syear, smonth, sday) 
if eday ="" then  eday =  Cstr(Day( dateadd("d",-1,DateSerial(Year(Date()), Month(Date()), 1)) )) 
 deddate = DateSerial(syear, smonth, eday)   
 if dateadd("m",1,dstdate) <= deddate then deddate = dateadd("d",dateadd("m",1,dstdate),-1)   
 
   if dispCate = "Y" then
    set clsCMeachulLog = new CMeachulLog
    clsCMeachulLog.FdateType =dategbn
	clsCMeachulLog.FstDate =dstdate
	clsCMeachulLog.FedDate =deddate
    clsCMeachulLog.Fcatekind = catekind
	cateList = clsCMeachulLog.fnGetCateList
     set clsCMeachulLog = nothing
    end if


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
   var sourceUrl = "/admin/contribution/json_List.asp"; 
    if(param) {
        sourceUrl = "/admin/contribution/json_List.asp?" + param;
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
       
        "export": { // 엑셀 다운로드 관련
            enabled: true,
            fileName: "공헌이익",
            allowExportSelectedData: true
        },
        headerFilter: { // 컬럼명 깔대기 검색 
            visible: false
        },
        columnAutoWidth: true,
         columns: [{
            dataField: "구분",
            caption: "구분"   
        },   
         {
                caption: "전체",
                alignment: "center",
                columns: [{
                    caption: "주문건수",
                    dataField: "전체_주문건수", 
                    dataType:"number", 
                    alignment: "right",
                    format:  "fixedPoint"     
                },{
                    caption: "상품수량",
                    dataField: "전체_상품수량", 
                    dataType:"number", 
                    alignment: "right",
                    format:  "fixedPoint"     
                },
                {
                    caption: "금액",
                    dataField: "전체_금액", 
                    dataType:"number", 
                    alignment: "right",
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
             
            <%
                For intC = 0 To UBound(cateList,2)
            %> 
             {
                caption: "<%=cateList(1,intC)%>",
                alignment: "center",
                columns: [{
                    caption: "상품수량",
                    dataField: "<%=cateList(1,intC)%>_상품수량", 
                    dataType:"number",
                    format:  "fixedPoint"    
                },{
                    caption: "금액",
                    dataField: "<%=cateList(1,intC)%>_금액", 
                    dataType:"number",
                    format:  "fixedPoint"    
                }, {
                    caption: "율",
                    dataField: "<%=cateList(1,intC)%>_율", 
                   format: {
    					type: "percent",
    					precision: 1
    				}
                }]
            }, 
            <%  Next
               END IF%> 
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