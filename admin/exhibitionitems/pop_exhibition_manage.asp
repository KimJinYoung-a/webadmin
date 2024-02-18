<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<%
'###############################################
' PageName : pop_exhibition_manage.asp
' Discription : 기획전 상품 관리 팝업
' History : 2018.11.06 이종화
'###############################################

dim gidx, page , mode, searchMastercode, searchisUsing

gidx = requestCheckVar(request("gidx"),8)
page = requestCheckVar(request("page"),8)
searchMastercode = requestCheckVar(request("mastercode"),8)
searchisUsing = requestCheckVar(request("isusing"),1)

if gidx="" then gidx=0
if page="" then page=1
if searchisUsing<>"a" and searchisUsing="" then searchisUsing="1"

if gidx = 0 then 
    mode = "gubunAdd"
else
    mode = "gubunModify"
end if 

dim oExhibition,oExhibitionList

set oExhibition = new ExhibitionCls
oExhibition.FrectGcode = gidx
oExhibition.getOneGroupItem

set oExhibitionList = new ExhibitionCls
oExhibitionList.FPageSize=20
oExhibitionList.FCurrPage= page
oExhibitionList.FrectMasterCode = searchMastercode
oExhibitionList.FrectIsusing = searchisUsing
oExhibitionList.getGroupList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function SavePosCode(frm){
    <% if mode = "gubunAdd" then %>
    if((!frm.gubuncode[0].checked)&&(!frm.gubuncode[1].checked)){
        alert('기획전 구분을 선택 해주세요');
        frm.gubuncode.focus;
        return;
    }
    <% end if %>
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
    
}

function selectExhibition(v){
    if (v == 1) {
        $("#mastercode").hide();
        $("#detailcode").hide();
        $("#titlename").text("기획전이름");
    } else {
        $("#mastercode").show();
        $("#detailcode").show();
        $("#titlename").text("카테고리이름");
    }
}

function goPage(pg,gidx) {
    document.frmSearch.page.value=pg;
    if(gidx!=undefined){
        document.frmSearch.gidx.value=gidx;
    }
    document.frmSearch.submit();
}
</script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style>
.maintitle {color:red}
</style>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<div>
            <table width="660" cellpadding="2" cellspacing="1" class="tbType1 listTb">
            <form name="frmposcode" method="post" action="/admin/exhibitionitems/lib/exhibition_proc.asp" >
            <input type="hidden" name="mode" value="<%=mode%>" />
            <input type="hidden" name="gidx" value="<%=oExhibition.FOneItem.Fgidx%>" />
            <th colspan="2" style="padding:10px;">기획전상품관리 - 그룹관리</th>
            <% if oExhibition.FOneItem.Fgidx = "" or oExhibition.FOneItem.Fgidx = 0 then %>
            <tr>
                <th width="200px;">기획전 구분 선택</th>
                <td style="text-align:left;">
                    <input type="radio" name="gubuncode" value="1" id="gubun1" onclick="selectExhibition(1);"><label for="gubun1" onclick="selectExhibition(1);">기획전</label>
                    <input type="radio" name="gubuncode" value="2" id="gubun2" onclick="selectExhibition(2);"><label for="gubun2" onclick="selectExhibition(2);">상세구분</label>
                </td>
            </tr>
            <% end if %>
            <%  %>
            <%'// 기획전 selectBox %>
            <tr id="mastercode" style="display:<%=chkiif(oExhibition.FOneItem.Fgidx = 0 or oExhibition.FOneItem.Fgubuncode = 1,"none","")%>;">
                <th>기획전 목록</th>
                <td style="text-align:left;"><%=DrawSelectAllView("mastercode",oExhibition.FOneItem.Fmastercode,"")%></td>
            </tr>
            <%'// 기획전 selectBox %>
            <tr>
                <th id="titlename"><%=chkiif(oExhibition.FOneItem.Fgubuncode = "" or oExhibition.FOneItem.Fgubuncode = 1 ,"기획전이름","카테고리이름")%></th>
                <td style="text-align:left;"><input type="text" name="title" value="<%=oExhibition.FOneItem.Ftitle%>"/></td>
            </tr>
            <tr id="detailcode" style="display:<%=chkiif(oExhibition.FOneItem.Fgidx = 0,"none","")%>;">
                <th>카테고리코드</th>
                <td style="text-align:left;"><input type="text" name="detailcode" value="<%=oExhibition.FOneItem.Fdetailcode%>" size="5"/>
                <br/>
                ex) 기획전 메인 - 0
                <br/>
                ex) 카테고리 - 10 , 20 , 30 (숫자만)
                </td>    
            </tr>
            <tr>
                <th>사용여부</th>
                <td style="text-align:left;">
                    <input type="radio" name="isusing" value="1" id="usey" <%=chkiif(oExhibition.FOneItem.Fisusing = "" or oExhibition.FOneItem.Fisusing = "1","checked","")%>><label for="usey">사용함</label>
                    <input type="radio" name="isusing" value="0" id="usen" <%=chkiif(oExhibition.FOneItem.Fisusing = "0","checked","")%>><label for="usen">사용안함</div>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SavePosCode(frmposcode);"></td>
            </tr>
            </form>
            </table>
            <%
            set oExhibition = Nothing
            %>
            <br>
        </div>
        <div class="tPad15">
            <table width="660" cellpadding="2" cellspacing="1" class="tbType1 listTb">
            <tr>
                <td colspan="5">
                    <form name="frmSearch" method="GET" action="">
                    <input type="hidden" name="page" value="<%=page%>" />
                    <input type="hidden" name="mode" value="<%=mode%>" />
                    <input type="hidden" name="gidx" value="<%=gidx%>" />
                    <div style="float:left;text-align:legt">
                        <label>기확전 : <%=DrawSelectAllView("mastercode",searchMastercode,"onchange=""goPage(1)""")%></label> /
                        <label>사용여부 :
                            <select name="isusing" onchange="goPage(1)">
                            <option value="a" <%=chkIIF(searchisUsing="a","selected","")%>>::전체::</option>
                            <option value="1" <%=chkIIF(searchisUsing="1","selected","")%>>사용</option>
                            <option value="0" <%=chkIIF(searchisUsing="0","selected","")%>>사용안함</option>
                            </select>
                        </label>
                    </div>
                    <div style="float:right;text-align:right">
                        <a href="#" onclick="goPage(1,'')"><img src="/images/icon_new_registration.gif" border="0"></a>
                    </div>
                    </form>
                </td>
            </tr>
            <tr>
                <th width="100">idx</th>
                <th>기획전</th>
                <th>카테고리명</th>
                <th>연결번호</th>
                <th>사용여부</th>
            </tr>
            <% for i=0 to oExhibitionList.FResultCount-1 %>
            <tr>
                <td><%= oExhibitionList.FItemList(i).Fgidx %></td>
                <% if oExhibitionList.FItemList(i).Fdetailcode = 0 then  %>
                <td style="text-align:left"><a href="#" onclick="goPage(<%=page%>,<%= oExhibitionList.FItemList(i).Fgidx %>)"><span class="maintitle"><%=oExhibitionList.FItemList(i).Ftitle%></span></a></td>
                <td style="text-align:left"><a href="#" onclick="goPage(<%=page%>,<%= oExhibitionList.FItemList(i).Fgidx %>)"><span class="maintitle">기획전 메인</span></a></td>
                <td><span class="maintitle">MasterCode : <%=oExhibitionList.FItemList(i).Fmastercode%></span></td>
                <% else %>
                <td style="text-align:left"><a href="#" onclick="goPage(<%=page%>,<%= oExhibitionList.FItemList(i).Fgidx %>)">&nbsp;ㄴ<%=getMasterCodeName(oExhibitionList.FItemList(i).Fmastercode)%></a></td>
                <td style="text-align:left"><a href="#" onclick="goPage(<%=page%>,<%= oExhibitionList.FItemList(i).Fgidx %>)">&nbsp;ㄴ<%=oExhibitionList.FItemList(i).Ftitle%></a></td>
                <td>DetailCode : <%=oExhibitionList.FItemList(i).Fdetailcode%></td>
                <% end if %>
                <td><%= chkiif(oExhibitionList.FItemList(i).Fisusing,"사용","사용안함") %></td>    
            </tr>
            <% next %>
            <tr>
                <td colspan="5" align="center">
                <% if oExhibitionList.HasPreScroll then %>
                    <a href="#" onclick="goPage(<%= oExhibitionList.StartScrollPage-1 %>)">[pre]</a>
                <% else %>
                    [pre]
                <% end if %>

                <% for i=0 + oExhibitionList.StartScrollPage to oExhibitionList.FScrollCount + oExhibitionList.StartScrollPage - 1 %>
                    <% if i>oExhibitionList.FTotalpage then Exit for %>
                    <% if CStr(page)=CStr(i) then %>
                    <font color="red">[<%= i %>]</font>
                    <% else %>
                    <a href="#" onclick="goPage(<%= i %>)">[<%= i %>]</a>
                    <% end if %>
                <% next %>

                <% if oExhibitionList.HasNextScroll then %>
                    <a href="#" onclick="goPage(<%= i %>)">[next]</a>
                <% else %>
                    [next]
                <% end if %>
                </td>
            </tr>
            </table>
        </div>
    </div>
</div>
<%
    set oExhibitionList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->