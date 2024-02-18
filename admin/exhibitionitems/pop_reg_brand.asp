<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기획전 브랜드 관리 팝업
' History : 2022-12-22 허진원
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
        mode = "brandAdd"
    else
        mode = "brandModify"
    end if 

    '// 이미지 업로드용
    Dim userid, encUsrId, tmpTx, tmpRn
        userid = session("ssBctId")

        Randomize()
        tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
        tmpRn = tmpTx(int(Rnd*26))
        tmpRn = tmpRn & tmpTx(int(Rnd*26))
        encUsrId = tenEnc(tmpRn & userid)
    '// 이미지 업로드용

    set oExhibition = new ExhibitionCls
        oExhibition.Frectidx = idx
        oExhibition.getOneBrandContents()
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
function fnBrandSave(frm){
    if (!frm.makerid.value) {
        alert("브랜드ID를 입력 해주세요.");
        frm.makerid.focus;
        return false;
    }

    if (!frm.StartDate.value) {
        alert("배너 노출 시작일을 입력 해주세요.");
        frm.StartDate.focus;
        return false;
    }

    if (!frm.EndDate.value) {
        alert("배너 노출 종료일을 입력 해주세요.");
        frm.EndDate.focus;
        return false;
    }
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

// 업로드 파일 확인 및 처리
function jsCheckUpload() {
    if($("#fileupload").val()!="") {
        $("#fileupmode").val("upload");

        $('#ajaxform').ajaxSubmit({
            //보내기전 validation check가 필요할경우
            beforeSubmit: function (data, frm, opt) {
                if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].upfile.value)) {
                    alert("JPG,PNG 이미지파일만 업로드 하실 수 있습니다.");
                    $("#fileupload").val("");
                    return false;
                }
                $("#lyrPrgs").show();
            },
            //submit이후의 처리
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
                    alert("처리중 오류가 발생했습니다.\n" + responseText);
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

// 물리적인 파일 삭제 처리
function jsDelImg(){
    if(confirm("이미지를 삭제하시겠습니까?\n\n※ 파일이 완전히 삭제되며 복구 할 수 없습니다.")){
        if($("#filepre").val()!="") {
            $("#fileupmode").val("delete");

            $('#ajaxform').ajaxSubmit({
                //보내기전
                beforeSubmit: function (data, frm, opt) {
                    $("#lyrPrgs").show();
                },
                //submit이후의 처리
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
                        alert("처리중 오류가 발생했습니다.\n" + responseText);
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
							<td colspan="2" ><b>브랜드 등록</b></td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<th> 브랜드ID </th>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<% drawSelectBoxDesignerwithName "makerid",oExhibition.FOneItem.Fmakerid %>
                                <div class="tPad15" id="infomenu" style="display:<%=chkiif(idx > 0 , "", "none")%>;">
                                    <div>브랜드명 : <span id="socname"><%=oExhibition.FOneItem.FsocName%></span></div>
                                </div>
							</td>
						</tr>
                        <tr>
                            <th>시작일</th>
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
                            <th>종료일</th>
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
                            <th>이미지 등록</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="hidden" name="bannerImg" value="<%=bannerImg%>" />
                                <div style="width:220px; height:220px;">
                                    <div id="lyrPrgs" style="display:none; position:absolute;padding:101px; background-color:rgba(0,0,0,0.2);"><img src="http://fiximage.10x10.co.kr/web2015/giftcard/ajax_loader.gif" alt="progress" /></div>
                                    <img id="lyrBnrImg" src="<%=chkIIF(bannerImg="" or isNull(bannerImg),"/images/admin_login_logo2.png",bannerImg)%>" style="height:218px; border:1px solid #EEE;"/>
                                </div>
                                <div id="lyrImgDelBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="" or IsNull(bannerImg),"display:none;","")%>" onclick="jsDelImg();">이미지 삭제</button></div>
                                <div id="lyrImgUpBtn" class="btn" style="<%=chkIIF(idx = 0 or bannerImg="" or IsNull(bannerImg),"","display:none;")%>"><label for="fileupload">이미지 업로드</label></div>
                            </td>
                        </tr>
                        <tr>
                            <th>우선순위</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="sortNo" value="<%=chkiif(oExhibition.FOneItem.FsortNo = "","99",oExhibition.FOneItem.FsortNo)%>">
                            </td>
                        </tr>
                        <tr>
                            <th>사용여부</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="radio" name="evtisusing" value="1" id="usey" <%=chkiif(oExhibition.FOneItem.Fisusing = ""  or oExhibition.FOneItem.Fisusing = "1" , "checked" , "")%>> <label for="usey">사용함</label>
                                <input type="radio" name="evtisusing" value="0" id="usen" <%=chkiif(oExhibition.FOneItem.Fisusing = "0" , "checked" , "")%>> <label for="usen">사용안함</label>
                            </td>
                        </tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2">
					<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="fnBrandSave(frmreg);" style="cursor:pointer">
					<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="window.close();" style="cursor:pointer">
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
<%'// 이미지 업로드 %>
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
<!-- 리스트 끝 -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->