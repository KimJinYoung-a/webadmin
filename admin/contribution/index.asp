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
Dim i,syear,smonth, sday, eday,dategbn,dstdate,deddate, catecodesearch

    dispCate =  requestCheckvar(request("blnDp"),10) 
    catekind =  requestCheckvar(request("rdoCK"),1) 
 
    dategbn     = requestCheckvar(request("dategbn"),32)
   	syear     = requestcheckvar(request("sy"),4)
	smonth     = requestcheckvar(request("sM"),2)
    sday     = requestcheckvar(request("sd"),2)
    eday     = requestCheckvar(request("ed"),2)
    catecodesearch = requestCheckvar(request("catecode"),3)

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
    If Trim(catecodesearch) <> "" Then
        clsCMeachulLog.FcatecodeSearch = catecodesearch
    End If
	cateList = clsCMeachulLog.fnGetCateList
     set clsCMeachulLog = nothing
    end if
%>
<link rel="stylesheet" href="/css/reset.css">
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
<link rel="stylesheet" href="https://cdn3.devexpress.com/jslib/20.1.7/css/dx.common.css">
<link rel="stylesheet" href="https://cdn3.devexpress.com/jslib/20.1.7/css/dx.light.css">
<!-- DevExtreme library -->
<script type="text/javascript" src="https://cdn3.devexpress.com/jslib/20.1.7/js/dx.all.js"></script>
<script type="text/javascript" src="https://cdn3.devexpress.com/jslib/20.1.7/js/dx.web.js"></script>
<script type="text/javascript" src="https://cdn3.devexpress.com/jslib/20.1.7/js/dx.viz.js"></script> 
<script type="text/javascript" src="https://cdn3.devexpress.com/jslib/20.1.7/js/dx.viz-web.js"></script>
<style>
#gridContainer {
    max-height: 800px; 
}
</style>
<script type="text/javascript">
function jsChangeCk(){ 
  if ($("input:checkbox[id='blnDp']").is(":checked") ){  
      $("#rdoCK1").prop("checked", true);  
      $("#rdoCK1").prop("disabled", false);  
      $("#rdoCK2").prop("disabled", false);  
      $("#catecode").show();
  } else {
      $("#rdoCK1").prop("checked", false);  
      $("#rdoCK1").prop("disabled", true);  
      $("#rdoCK2").prop("disabled", true);  
      $("#catecode").hide();
  }
}

function jsSetDB(smode){
    if (confirm("등록하시겠습니까?")){
        document.frmDB.hidM.value =  smode;
        document.frmDB.yyyymm.value =  "<%=left(dstdate,7)%>"; 
        $(".procBtn").prop("disabled", true);
        document.frmDB.submit(); 
    }
}

function jsCsv(){
      var param = $(document.frm).serialize();
         hidfrm.location.href= "/admin/contribution/csv_list.asp?" + param;
}
</script>
<iframe name="hidfrm" width="0" height="0" ></iframe>
 <div class="card">
        <div class="card-body">
            <h4 class="card-titile" style="padding:10px;border-bottom:1px solid #e4e4e4;">검색 조건</h4>
            <form name="frm" method="get" action="" class="mb-0">
                <div  style="width:400px; display:inline-block;">
                    <div  >
                         <span style="font-size:12px;padding:10px;"> 날짜 : </span>
                        <select class="select" name="dategbn" onChange="jsh4SetYYYYMM4()"> 
                        	<option value="regdate" <%=CHKIIF(dategbn="regdate","selected","")%>>주문일</option>
                            <option value="ipkumdate" <%=CHKIIF(dategbn="ipkumdate","selected","")%> >결제일자</option>
                            <option value="beasongdate" <%=CHKIIF(dategbn="beasongdate","selected","")%> >출고일자</option>
                            <option value="jFixedDt" <%=CHKIIF(dategbn="jFixedDt","selected","")%> >정산확정일자</option> 
                        </select>
                        &nbsp;
                        <select name="sY">
                        <%for i=year(now()) to 2002 step -1%>
                        <option value="<%=i%>" <%if Cint(sYear) = cint(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                        </select>
                         <select name="sM">
                        <%for i=1 to 12%>
                        <option value="<%=i%>" <%if cInt(sMonth) = cInt(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                        </select>
                         <select name="sD">
                        <%for i=1 to 31%>
                        <option value="<%=i%>" <%if cInt(sDay) = cInt(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                        </select>
                        ~
                         <select name="eD">
                        <%for i=1 to 31%>
                        <option value="<%=i%>" <%if cInt(eDay) = cInt(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                        </select> 
                    </div> 
                </div>
                 <div style="display:inline-block;width:300px">
                    <input type="checkbox" name="blnDp" id="blnDp" value="Y" <%if dispCate="Y" then%>checked<%end if%> onclick="jsChangeCk();"> 카테고리 보기
                    <input type="hidden" name="rdoCK" id="rdoCK1" value="D" <%if dispCate<>"Y" then%>disabled<%else%><%if catekind ="D" then%>checked<%end if%><%end if%>>
                    <select name="catecode" id="catecode" class="select" <%if dispCate<>"Y" then%>style="display:none;"<% End If %>>
                        <option value="" <% If catecodesearch="" Then %>selected<% End If %>>전체</option>
                        <option value="101" <% If catecodesearch="101" Then %>selected<% End If %>>디자인문구</option>
                        <option value="102" <% If catecodesearch="102" Then %>selected<% End If %>>디지털/핸드폰</option>
                        <option value="104" <% If catecodesearch="104" Then %>selected<% End If %>>토이/취미</option>
                        <option value="124" <% If catecodesearch="124" Then %>selected<% End If %>>디자인가전</option>
                        <option value="121" <% If catecodesearch="121" Then %>selected<% End If %>>가구/수납</option>
                        <option value="122" <% If catecodesearch="122" Then %>selected<% End If %>>데코/조명</option>
                        <option value="120" <% If catecodesearch="120" Then %>selected<% End If %>>패브릭/생활</option>
                        <option value="112" <% If catecodesearch="112" Then %>selected<% End If %>>키친</option>
                        <option value="119" <% If catecodesearch="119" Then %>selected<% End If %>>푸드</option>
                        <option value="117" <% If catecodesearch="117" Then %>selected<% End If %>>패션의류</option>
                        <option value="116" <% If catecodesearch="116" Then %>selected<% End If %>>패션잡화</option>
                        <option value="125" <% If catecodesearch="125" Then %>selected<% End If %>>주얼리/시계</option>
                        <option value="118" <% If catecodesearch="118" Then %>selected<% End If %>>뷰티</option>
                        <option value="115" <% If catecodesearch="115" Then %>selected<% End If %>>베이비/키즈</option>
                        <option value="110" <% If catecodesearch="110" Then %>selected<% End If %>>Cat & Dog</option>
                    </select>
                    <!--( <input type="radio" name="rdoCK" id="rdoCK1" value="D"  <%if dispCate<>"Y" then%>disabled<%else%><%if catekind ="D" then%>checked<%end if%><%end if%>> 전시카테고리
                    input type="radio" name="rdoCK"  id="rdoCK2" value="M" <%if dispCate<>"Y" then%>disabled<%else%> <%if catekind ="M" then%>checked<%end if%> <%end if%>> 관리카테고리 )-->
                 </div>
                <div style="display:inline-block;">
                    <button type="button"  class="button_s" style="width:100px;height:30px;" onclick="SearchForm(this.form);">검색</button>
                </div>
            </form>
        </div>
        <div style="padding: 10px 0;border-bottom:1px solid #e4e4e4;width:100%" ></div>
     </div> 
      <div  style="padding-top:10px;">
        <a href="javascript:jsCsv();"><div style="height:20px;width:100px; border:1px solid #e4e4e4;text-align:center;background-color:#ffffff;">▶ excel down </div></a>
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
    <%if C_ADMIN_AUTH Or C_MngPart or C_PSMngPart then%>
      <div style="padding:20px 0">
        <div> <span style="font-size:12px;padding:10px;"> DB 등록 : <%=left(dstdate,7)%> </span>
        <input type="button"  value="1. PG 수수료" class="button procBtn" onClick="jsSetDB('c1');" style="cursor:hand;">
        <input type="button"  value="2. 신용카드 수수료" class="button procBtn" onClick="jsSetDB('c2');" style="cursor:hand;">
        <input type="button"   value="3. CPS 수수료" class="button procBtn" onClick="jsSetDB('c3');" style="cursor:hand;">
        <input type="button"  value="4. 제휴 수수료" class="button procBtn" onClick="jsSetDB('c4');" style="cursor:hand;">
        <input type="button"   value="5. 라이센스 수수료" class="button procBtn" onClick="jsSetDB('c5');" style="cursor:hand;">
        <input type="button"   value="6. 물류비" class="button procBtn" onClick="jsSetDB('c6');" style="cursor:hand;">
        <input type="button"   value="7. 판촉/광고비" class="button procBtn" onClick="jsSetDB('c7');" style="cursor:hand;">
        </div>
        <div style="padding:10px 10px  0px 10px;">- 등록버튼을 누른 후 [ 등록되었습니다 ] 창이 뜰때까지 기다려주세요! </div>
        <form name="frmDB" method="post" action="procPL.asp">
        <input type="hidden" name="hidM" value="">
        <input type="hidden" name="yyyymm" value="">
        </form>
      <div style="padding: 10px 0;border-bottom:1px solid #e4e4e4;width:100%" ></div>
      </div>    
      <%end if%>  
    <!-- 메인 끝 -->
<script type="text/javascript">
$(function(){ 
    var param = $(document.frm).serialize();
    callDataGrid(param);
});

function callDataGrid(param) {
   var sourceUrl = "/admin/contribution/json_List.asp"; 
    if(param) {
        sourceUrl = "/admin/contribution/json_List.asp?" + param;
    } 
  
    $("#gridContainer").dxTreeList({ 
        rootValue: -1,
        keyExpr: "ID",
        parentIdExpr: "Head_ID",
        //expandedRowKeys: [0,13,26,39,52,65,66,67,72,75,76,77,78,92],
        columnAutoWidth: true,
        autoExpandAll : false,
        showBorders: true,
        showRowLines: true, 
        selection: {
            mode: "single"
        },
        scrolling: {
            mode: "standard"
        },
        paging: {
            enabled: true,
            pageSize: 100
        },
       
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