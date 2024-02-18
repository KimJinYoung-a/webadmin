<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ȹ�� �̺�Ʈ ���� �˾�
' History : 2018-11-07 ����ȭ
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%
    dim mastercode : mastercode = request("mastercode")
    dim idx : idx = request("idx")
    dim oExhibition
    dim mode, bannerImg

    if idx = 0 then 
        mode = "evtadd"
    else
        mode = "evtmodify"
    end if 

    '// �̹��� ���ε��
    Dim userid, encUsrId, tmpTx, tmpRn
        userid = session("ssBctId")

        Randomize()
        tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
        tmpRn = tmpTx(int(Rnd*26))
        tmpRn = tmpRn & tmpTx(int(Rnd*26))
        encUsrId = tenEnc(tmpRn & userid)
    '// �̹��� ���ε��

    set oExhibition = new ExhibitionCls
        oExhibition.Frectidx = idx
        oExhibition.getOneEventContents()
    bannerImg = oExhibition.FOneItem.FbannerImage
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript">
function fnEventSave(frm){
    if (!frm.evt_code.value) {
        alert("�̺�Ʈ �ڵ带 �Է� ���ּ���.");
        frm.evt_code.focus;
    }

    if (!frm.StartDate.value) {
        alert("��� ���� �������� �Է� ���ּ���.");
        frm.StartDate.focus;
    }

    if (!frm.EndDate.value) {
        alert("��� ���� �������� �Է� ���ּ���.");
        frm.EndDate.focus;
    }
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function jsLastEvent(num){
    winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
    winLast.focus();
}

// ���ε� ���� Ȯ�� �� ó��
function jsCheckUpload() {
    if($("#fileupload").val()!="") {
        $("#fileupmode").val("upload");

        $('#ajaxform').ajaxSubmit({
            //�������� validation check�� �ʿ��Ұ��
            beforeSubmit: function (data, frm, opt) {
                if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].upfile.value)) {
                    alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
                    $("#fileupload").val("");
                    return false;
                }
                $("#lyrPrgs").show();
            },
            //submit������ ó��
            success: function(responseText, statusText){
                var resultObj = JSON.parse(responseText)

                if(resultObj.response=="fail") {
                    alert(resultObj.faildesc);
                } else if(resultObj.response=="ok") {
                    document.frmreg.bannerImg.value=resultObj.fileurl;
                    $("#filepre").val(resultObj.fileurl);
                    $("#lyrBnrImg").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
                    $("#lyrImgUpBtn").hide();
                    $("#lyrImgDelBtn").show();
                } else {
                    alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
                }
                $("#fileupload").val("");
                $("#lyrPrgs").hide();
            },
            //ajax error
            error: function(err){
                alert("ERR: " + err.responseText);
                $("#fileupload").val("");
                $("#lyrPrgs").hide();
            }
        });
    }
}

// �������� ���� ���� ó��
function jsDelImg(){
    if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n�� ������ ������ �����Ǹ� ���� �� �� �����ϴ�.")){
        if($("#filepre").val()!="") {
            $("#fileupmode").val("delete");

            $('#ajaxform').ajaxSubmit({
                //��������
                beforeSubmit: function (data, frm, opt) {
                    $("#lyrPrgs").show();
                },
                //submit������ ó��
                success: function(responseText, statusText){
                    var resultObj = JSON.parse(responseText)

                    if(resultObj.response=="fail") {
                        alert(resultObj.faildesc);
                    } else if(resultObj.response=="ok") {
                        document.frmreg.bannerImg.value="";
                        $("#lyrBnrImg").hide().attr("src","/images/admin_login_logo2.png").fadeIn("fast");
                        $("#filepre").val("");
                        $("#lyrImgUpBtn").show();
                        $("#lyrImgDelBtn").hide();
                    } else {
                        alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
                    }
                    $("#lyrPrgs").hide();
                },
                //ajax error
                error: function(err){
                    alert("ERR: " + err.responseText);
                    $("#lyrPrgs").hide();
                }
            });
        }
    }
}
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<form name="frmreg" method="post" action="/admin/exhibitionitems/lib/exhibition_proc.asp">
        <input type="hidden" name="mode" value="<%=mode%>"/>
        <input type="hidden" name="eidx" value="<%=idx%>"/>
        <input type="hidden" name="mastercode" value="<%=mastercode%>">
		<table class="tbType1 listTb">
			<tr>
				<td>
					<table class="tbType1 listTb">
						<tr bgcolor="#FFFFFF" height="25">
							<td colspan="2" ><b>�̺�Ʈ ���</b></td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<th> �̺�Ʈ �ڵ�</th>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" size="10" name="evt_code" value="<%=oExhibition.FOneItem.Fevt_code%>"/> <input type="button" value="�̺�Ʈ �ҷ�����" onclick="jsLastEvent(1);"/>
                                <div class="tPad15" id="infomenu" style="display:<%=chkiif(idx > 0 , "", "none")%>;">
                                    <div>�̺�Ʈ�� : <span id="evt_name"><%=oExhibition.FOneItem.Fevt_name%></span></div>
                                    <div>������ : <span id="evt_startdate"><%=oExhibition.FOneItem.Fevt_startdate%></span></div>
                                    <div>������ : <span id="evt_enddate"><%=oExhibition.FOneItem.Fevt_enddate%></span></div>
                                    <div>������ : <span id="evt_saleper" style='color:red'><%=chkiif(oExhibition.FOneItem.Fsaleper <> "",oExhibition.FOneItem.Fsaleper,"���� ������ �����ϴ�.")%></span></div>
                                    <div>�������� : <span id="evt_salecoupon" style='color:green'><%=chkiif(oExhibition.FOneItem.Fsalecper <> "",oExhibition.FOneItem.Fsalecper,"�������� ������ �����ϴ�.")%></span></div>
                                </div>
							</td>
						</tr>
                        <tr>
                            <th>������</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="StartDate" id="startdate" value="<%=oExhibition.FOneItem.Fstartdate%>">
                                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" style="vertical-align:middle;"/>
                                <script type="text/javascript">
                                var CAL_Start = new Calendar({
                                    inputField : "startdate",
                                    trigger    : "startdate_trigger",
                                    onSelect: function() {
                                        var date = Calendar.intToDate(this.selection.get());
                                        CAL_End.args.min = date;
                                        CAL_End.redraw();
                                        this.hide();
                                    },
                                    bottomBar: true,
                                    dateFormat: "%Y-%m-%d"
                                });
                                </script>
                            </td>
                        </tr>
                        <tr>
                            <th>������</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="EndDate" id="enddate" value="<%=oExhibition.FOneItem.Fenddate%>">
                                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" style="vertical-align:middle;"/>
                                <script type="text/javascript">
                                var CAL_End = new Calendar({
                                    inputField : "enddate",
                                    trigger    : "enddate_trigger",
                                    onSelect: function() {
                                        var date = Calendar.intToDate(this.selection.get());
                                        CAL_Start.args.max = date;
                                        CAL_Start.redraw();
                                        this.hide();
                                    },
                                    bottomBar: true,
                                    dateFormat: "%Y-%m-%d"
                                });
                                </script>
                            </td>
                        </tr>
                        <tr>
                            <th>�̹��� ���</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="hidden" name="bannerImg" value="<%=bannerImg%>" />
                                <div style="width:220px; height:220px;">
                                    <div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                                    <img id="lyrBnrImg" src="<%=chkIIF(bannerImg="" or isNull(bannerImg),"/images/admin_login_logo2.png",bannerImg)%>" style="height:218px; border:1px solid #EEE;"/>
                                </div>
                                <div id="lyrImgDelBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="" or IsNull(bannerImg),"display:none;","")%>" onclick="jsDelImg();">�̹��� ����</button></div>
                                <div id="lyrImgUpBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="" or IsNull(bannerImg),"","display:none;")%>"><label for="fileupload">�̹��� ���ε�</label></div>
                            </td>
                        </tr>
                        <tr>
                            <th>�켱����</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="evtsorting" value="<%=chkiif(oExhibition.FOneItem.Fevtsorting = "","99",oExhibition.FOneItem.Fevtsorting)%>">
                            </td>
                        </tr>
                        <tr>
                            <th>��뿩��</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="radio" name="evtisusing" value="1" id="usey" <%=chkiif(oExhibition.FOneItem.Fisusing = ""  or oExhibition.FOneItem.Fisusing = "1" , "checked" , "")%>> <label for="usey">�����</label>
                                <input type="radio" name="evtisusing" value="0" id="usen" <%=chkiif(oExhibition.FOneItem.Fisusing = "0" , "checked" , "")%>> <label for="usen">������</label>
                            </td>
                        </tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2">
					<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="fnEventSave(frmreg);" style="cursor:pointer">
					<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="window.close();" style="cursor:pointer">
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
<%'// �̹��� ���ε� %>
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
    <input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
    <input type="hidden" name="mode" id="fileupmode" value="upload">
    <input type="hidden" name="div" value="SB">
    <input type="hidden" name="upPath" value="/event/swipeimage/">
    <input type="hidden" name="tuid" value="<%=encUsrId%>">
    <input type="hidden" name="prefile" id="filepre" value="<%=bannerImg%>">
</form>
<%
    set oExhibition = nothing
%>
<!-- ����Ʈ �� -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->