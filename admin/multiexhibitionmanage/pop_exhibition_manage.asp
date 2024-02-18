<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/multiexhibitionmanage/lib/classes/itemsCls.asp"-->
<%
'###############################################
' PageName : pop_exhibition_manage.asp
' Discription : ��ȹ�� ��ǰ ���� �˾�
' History : 2018.11.06 ����ȭ
'###############################################

dim gidx, page , mode , mcode

gidx = request("gidx")
mcode = request("mcode")
page = request("page")

if gidx="" then gidx=0
if page="" then page=1

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
oExhibitionList.getGroupList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function SavePosCode(frm){
    <% if mode = "gubunAdd" then %>
    if((!frm.gubuncode[0].checked)&&(!frm.gubuncode[1].checked)){
        alert('��ȹ�� ������ ���� ���ּ���');
        frm.gubuncode.focus;
        return;
    }
    <% end if %>

    if(!frm.typename.value){
        alert('��ȹ������ �Է� ���ּ���');
        frm.typename.focus;
        return;
    }

    if(!frm.type.value){
        alert('���и��� �Է� ���ּ���');
        frm.type.focus;
        return;
    }
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function selectExhibition(v){
    if (v == 1) {
        $("#mastercode").hide();
        $("#detailcode").hide();
        document.frmposcode.type.value = "��ȹ��";
        $("#gubunname").text("��ȹ��");
        $("#titlename").text("��ȹ���̸�");
    } else {
        $("#mastercode").show();
        $("#detailcode").show();
        document.frmposcode.type.value = "";
        $("#gubunname").text("�Ӽ�");
        $("#titlename").text("�����̸�");
    }
}

function mkbutton(mastercode) {
    var filtercode = 1;
    var targetform = "frmposcode";
    var targetname = "type";
    $.ajax({
        method : "get",
        url: "/admin/multiexhibitionmanage/lib/ajax_function.asp",
        data : "mastercode="+mastercode+"&filtercode="+filtercode+"&targetform="+targetform+"&targetname="+targetname,
        cache: false,
        async: false,
        success: function(message) {
            $("#submenu").empty().html(message).css("padding-top","10px");
        }
    });
}

$(function(){
    // init select
    <% if mode = "gubunModify" then %>
    mkbutton(<%=oExhibition.FOneItem.Fmastercode%>);
    <% end if %>
});
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
            <form name="frmposcode" method="post" action="/admin/multiexhibitionmanage/lib/manage_proc.asp" >
            <input type="hidden" name="mode" value="<%=mode%>">
            <input type="hidden" name="gidx" value="<%=oExhibition.FOneItem.Fgidx%>">
            <th colspan="2" style="padding:10px;">��ȹ����ǰ���� - �׷�/�ɼ� ����</th>
            <% if oExhibition.FOneItem.Fgidx = "" or oExhibition.FOneItem.Fgidx = 0 then %>
            <tr>
                <th width="200px;">��ȹ�� ���� ����</th>
                <td style="text-align:left;">
                    <input type="radio" name="gubuncode" value="1" id="gubun1" onclick="selectExhibition(1);"><label for="gubun1" onclick="selectExhibition(1);">��ȹ��</label>
                    <input type="radio" name="gubuncode" value="2" id="gubun2" onclick="selectExhibition(2);"><label for="gubun2" onclick="selectExhibition(2);">�󼼱���</label>
                </td>
            </tr>
            <% end if %>
            <%  %>
            <%'// ��ȹ�� selectBox %>
            <tr id="mastercode" style="display:<%=chkiif(oExhibition.FOneItem.Fgidx = 0 or oExhibition.FOneItem.Fgubuncode = 1,"none","")%>;">
                <th>��ȹ�� ���</th>
                <td style="text-align:left;">
                    <%=DrawSelectAllView("mastercode",oExhibition.FOneItem.Fmastercode,"mkbutton")%>
                    <div id="submenu"></div>
                </td>
            </tr>
            <%'// ��ȹ�� selectBox %>
            <tr>
                <th id="titlename"><%=chkiif(oExhibition.FOneItem.Fgubuncode = "" or oExhibition.FOneItem.Fgubuncode = 1 ,"��ȹ���̸�","�����̸�")%></th>
                <td style="text-align:left;">
                    <span id="gubunname">��ȹ��</span>�� : 
                    <input type="text" name="typename" value="<%=oExhibition.FOneItem.Ftypename%>" autocomplete="off"/>
                    / 
                    ���и� : 
                    <input type="text" name="type" value="<%=oExhibition.FOneItem.Ftype%>" autocomplete="off"/>
                </td>
            </tr>
            <tr id="detailcode" style="display:<%=chkiif(oExhibition.FOneItem.Fgidx = 0,"none","")%>;">
                <th>���ڵ�</th>
                <td style="text-align:left;">IDX : <%=oExhibition.FOneItem.Fdetailcode%>
                <input type="hidden" name="detailcode" value="<%=oExhibition.FOneItem.Fdetailcode%>"/>
                <br/>
                ex) ��ȹ�� ���� - 0
                <br/>
                ex) �Ӽ��ڵ� - 1~N (�ڵ�����)
                </td>    
            </tr>
            <tr>
                <th>��뿩��</th>
                <td style="text-align:left;">
                    <input type="radio" name="isusing" value="1" id="usey" <%=chkiif(oExhibition.FOneItem.Fisusing = "" or oExhibition.FOneItem.Fisusing = "1","checked","")%>><label for="usey">�����</label>
                    <input type="radio" name="isusing" value="0" id="usen" <%=chkiif(oExhibition.FOneItem.Fisusing = "0","checked","")%>><label for="usen">������</div>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SavePosCode(frmposcode);"></td>
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
                <td colspan="5" style="text-align:right"><a href="?gidx="><img src="/images/icon_new_registration.gif" border="0"></a></td>
            </tr>
            <tr>
                <th width="100">�ɼ��ڵ�</th>
                <th>��ȹ��</th>
                <th>�ɼǸ�</th>
                <th>���и�</th>
                <th>��뿩��</th>
            </tr>
            <% for i=0 to oExhibitionList.FResultCount-1 %>
            <tr>
                <td><%= oExhibitionList.FItemList(i).Fdetailcode %></td>
                <% if oExhibitionList.FItemList(i).Fdetailcode = 0 then  %>
                <td style="text-align:left"><a href="?gidx=<%= oExhibitionList.FItemList(i).Fgidx %>&page=<%= page %>"><span class="maintitle"><%=oExhibitionList.FItemList(i).Ftypename%></span></a></td>
                <td style="text-align:left"><a href="?gidx=<%= oExhibitionList.FItemList(i).Fgidx %>&page=<%= page %>"><span class="maintitle">��ȹ�� ����</span></a></td>
                <td style="text-align:left"><a href="?gidx=<%= oExhibitionList.FItemList(i).Fgidx %>&page=<%= page %>"><span class="maintitle"><%=oExhibitionList.FItemList(i).Ftype%></span></a></td>
                <% else %>
                <td style="text-align:left"><a href="?gidx=<%= oExhibitionList.FItemList(i).Fgidx %>&page=<%= page %>">&nbsp;��<%=getMasterCodeName(oExhibitionList.FItemList(i).Fmastercode)%></a></td>
                <td style="text-align:left"><a href="?gidx=<%= oExhibitionList.FItemList(i).Fgidx %>&page=<%= page %>">&nbsp;��<%=oExhibitionList.FItemList(i).Ftypename%></a></td>
                <td style="text-align:left"><a href="?gidx=<%= oExhibitionList.FItemList(i).Fgidx %>&page=<%= page %>">&nbsp;��<%=oExhibitionList.FItemList(i).Ftype%></a></td>
                <% end if %>
                <td><%= chkiif(oExhibitionList.FItemList(i).Fisusing,"���","������") %></td>    
            </tr>
            <% next %>
            <tr>
                <td colspan="5" align="center">
                <% if oExhibitionList.HasPreScroll then %>
                    <a href="?page=<%= oExhibitionList.StartScrollPage-1 %>">[pre]</a>
                <% else %>
                    [pre]
                <% end if %>

                <% for i=0 + oExhibitionList.StartScrollPage to oExhibitionList.FScrollCount + oExhibitionList.StartScrollPage - 1 %>
                    <% if i>oExhibitionList.FTotalpage then Exit for %>
                    <% if CStr(page)=CStr(i) then %>
                    <font color="red">[<%= i %>]</font>
                    <% else %>
                    <a href="?page=<%= i %>">[<%= i %>]</a>
                    <% end if %>
                <% next %>

                <% if oExhibitionList.HasNextScroll then %>
                    <a href="?page=<%= i %>">[next]</a>
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