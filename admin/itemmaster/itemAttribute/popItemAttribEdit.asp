<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : 상품 속성 등록/수정
' History : 2013.08.02 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim attribCd, i
Dim oAttrib

Dim attribDiv,attribDivName,attribName,attribNameAdd,attribUsing,attribSortNo
Dim mobile_image1, mobile_image2, mobile_image3, mobile_image4, mobile_image5, mobile_image6
Dim pc_image1, pc_image2, pc_image3, pc_image4, pc_image5, pc_image6


'// 파라메터 접수
attribCd = request("attribCd")

'// 수정할 속성내용 접수
if attribCd<>"" then
	set oAttrib = new CAttrib
		oAttrib.FRectattribCd = attribCd
		oAttrib.GetOneAttrib
		if oAttrib.FResultCount>0 then
			attribDiv		= oAttrib.FOneItem.FattribDiv
			attribDivName	= oAttrib.FOneItem.FattribDivName
			attribName		= oAttrib.FOneItem.FattribName
			attribNameAdd	= oAttrib.FOneItem.FattribNameAdd
			attribUsing		= oAttrib.FOneItem.FattribUsing
			attribSortNo	= oAttrib.FOneItem.FattribSortNo
			mobile_image1	= oAttrib.FOneItem.Fmobile_image1
			mobile_image2	= oAttrib.FOneItem.Fmobile_image2
			mobile_image3	= oAttrib.FOneItem.Fmobile_image3
			mobile_image4	= oAttrib.FOneItem.Fmobile_image4
			mobile_image5	= oAttrib.FOneItem.Fmobile_image5
			mobile_image6	= oAttrib.FOneItem.Fmobile_image6
			pc_image1	= oAttrib.FOneItem.Fpc_image1
            pc_image2	= oAttrib.FOneItem.Fpc_image2
            pc_image3	= oAttrib.FOneItem.Fpc_image3
            pc_image4	= oAttrib.FOneItem.Fpc_image4
            pc_image5	= oAttrib.FOneItem.Fpc_image5
            pc_image6	= oAttrib.FOneItem.Fpc_image6
		end if
	set oAttrib = Nothing
end if

if attribSortNo="" or isNull(attribSortNo) then attribSortNo="0"
if attribUsing="" or isNull(attribUsing) then attribUsing="Y"
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css">
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css">
<style type="text/css">
html {overflow-y:auto;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function(){
	//라디오버튼
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

// 폼검사
function SaveForm(frm) {
	<% if attribCd="" then %>
	if(frm.newDiv.checked) {
		if(frm.newAttrDivName.value=="") {
			alert("새로 추가할 속성구분명을 입력해주세요.");
			frm.newAttrDivName.focus();
			return;
		} else {
			frm.attribDivName.value=frm.newAttrDivName.value;
		}
	} else {
		if(frm.attribDiv.value=="") {
			alert("상품속성 구분을 선택해주세요.\n속성 구분이 없으면 구분생성을 선택 후 등록해주세요.");
			frm.attribDiv.focus();
			return;
		}
	}
	<% end if %>

	if(frm.attribName.value=="") {
		alert("상품속성명을 입력해주세요.");
		frm.attribName.focus();
		return;
	}

	if(frm.attribSortNo.value=="") {
		alert("상품속성 정렬순서를 숫자로 입력해주세요.");
		frm.attribSortNo.focus();
		return;
	}

	if(confirm("입력하신 내용으로 <%=chkIIF(attribCd="","등록","수정")%>하시겠습니까?")) {
	    <% IF attribCd <> "" THEN %>
            save_image().then(function(){
                frm.submit();
            });
        <% ELSE %>
            frm.submit();
        <% END IF %>
	} else {
		return;
	}
}

// 신규속성구분 추가
function fnNewAttrDiv(chk) {
	if(chk) {
		$("#lyrAttrDiv").show();
	} else {
		$("#lyrAttrDiv").hide();
	}
}

function change_image(name, file){
    let reader = new FileReader();
    reader.readAsDataURL(file.files[0]);

    switch (name){
        case "mobile_image1" :
            reader.onload = function(e){
                $("#showMobileImage1").attr("src", e.target.result);
            }
            break;
        case "mobile_image2" :
            reader.onload = function(e){
                $("#showMobileImage2").attr("src", e.target.result);
            }
            break;
        case "mobile_image3" :
            reader.onload = function(e){
                $("#showMobileImage3").attr("src", e.target.result);
            }
            break;
        case "mobile_image4" :
            reader.onload = function(e){
                $("#showMobileImage4").attr("src", e.target.result);
            }
            break;
        case "mobile_image5" :
            reader.onload = function(e){
                $("#showMobileImage5").attr("src", e.target.result);
            }
            break;
        case "mobile_image6" :
            reader.onload = function(e){
                $("#showMobileImage6").attr("src", e.target.result);
            }
            break;
        case "pc_image1" :
        	reader.onload = function(e){
        		$("#showPcImage1").attr("src", e.target.result);
        	}
        	break;
        case "pc_image2" :
        	reader.onload = function(e){
        		$("#showPcImage2").attr("src", e.target.result);
        	}
        	break;
        case "pc_image3" :
        	reader.onload = function(e){
        		$("#showPcImage3").attr("src", e.target.result);
        	}
        	break;
        case "pc_image4" :
        	reader.onload = function(e){
        		$("#showPcImage4").attr("src", e.target.result);
        	}
        	break;
        case "pc_image5" :
        	reader.onload = function(e){
        		$("#showPcImage5").attr("src", e.target.result);
        	}
        	break;
        case "pc_image6" :
        	reader.onload = function(e){
        		$("#showPcImage6").attr("src", e.target.result);
        	}
        	break;
    }
}

function save_image(){
    return new Promise(function(resolve, reject){
        let imgData = new FormData();
        imgData.append("attribCd", "<%=attribCd%>");
        if(document.getElementById("mobile_image1").files[0]){
            imgData.append("mobile_image1", document.getElementById("mobile_image1").files[0]);
        }
        if(document.getElementById("mobile_image2").files[0]){
            imgData.append("mobile_image2", document.getElementById("mobile_image2").files[0]);
        }
        if(document.getElementById("mobile_image3").files[0]){
            imgData.append("mobile_image3", document.getElementById("mobile_image3").files[0]);
        }
        if(document.getElementById("mobile_image4").files[0]){
            imgData.append("mobile_image4", document.getElementById("mobile_image4").files[0]);
        }
        if(document.getElementById("mobile_image5").files[0]){
            imgData.append("mobile_image5", document.getElementById("mobile_image5").files[0]);
        }
        if(document.getElementById("mobile_image6").files[0]){
            imgData.append("mobile_image6", document.getElementById("mobile_image6").files[0]);
        }
        if(document.getElementById("pc_image1").files[0]){
        	imgData.append("pc_image1", document.getElementById("pc_image1").files[0]);
        }
        if(document.getElementById("pc_image2").files[0]){
        	imgData.append("pc_image2", document.getElementById("pc_image2").files[0]);
        }
        if(document.getElementById("pc_image3").files[0]){
        	imgData.append("pc_image3", document.getElementById("pc_image3").files[0]);
        }
        if(document.getElementById("pc_image4").files[0]){
        	imgData.append("pc_image4", document.getElementById("pc_image4").files[0]);
        }
        if(document.getElementById("pc_image5").files[0]){
        	imgData.append("pc_image5", document.getElementById("pc_image5").files[0]);
        }
        if(document.getElementById("pc_image6").files[0]){
        	imgData.append("pc_image6", document.getElementById("pc_image6").files[0]);
        }

        let api_url;
        if (location.hostname.startsWith('webadmin')) {
            api_url = 'https://upload.10x10.co.kr';
        } else {
            api_url = 'http://testupload.10x10.co.kr';
        }
        $.ajax({
            url: api_url + "/linkweb/item_attribute/itemAttributeDetail_admin_imgreg_json.asp"
            , type: "POST"
            , processData: false
            , contentType: false
            , data: imgData
            , crossDomain: true
            , success: function (data) {
                const response = JSON.parse(data);
                console.log(data);

                if(response.mobile_image1){
                    $("input[name=mobile_image1]").val(response.mobile_image1);
                }
                if(response.mobile_image2){
                    $("input[name=mobile_image2]").val(response.mobile_image2);
                }
                if(response.mobile_image3){
                    $("input[name=mobile_image3]").val(response.mobile_image3);
                }
                if(response.mobile_image4){
                    $("input[name=mobile_image4]").val(response.mobile_image4);
                }
                if(response.mobile_image5){
                    $("input[name=mobile_image5]").val(response.mobile_image5);
                }
                if(response.mobile_image6){
                    $("input[name=mobile_image6]").val(response.mobile_image6);
                }
                if(response.pc_image1){
                	$("input[name=pc_image1]").val(response.pc_image1);
                }
                if(response.pc_image2){
                	$("input[name=pc_image2]").val(response.pc_image2);
                }
                if(response.pc_image3){
                	$("input[name=pc_image3]").val(response.pc_image3);
                }
                if(response.pc_image4){
                	$("input[name=pc_image4]").val(response.pc_image4);
                }
                if(response.pc_image5){
                	$("input[name=pc_image5]").val(response.pc_image5);
                }
                if(response.pc_image6){
                	$("input[name=pc_image6]").val(response.pc_image6);
                }

                return resolve();
            }
            , error : function (request,status,error){
                console.log("code", request.status);
                console.log("message", request.responseText);
                console.log("error", error);

                return reject();
            }
        });
    });
}
</script>
<div class="pad20">
<form name="frmSub" method="post" action="doItemAttrModify.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=chkIIF(attribCd="","attrNew","attrModi")%>">
<input type="hidden" name="attribCd" value="<%=attribCd%>" />
<input type="hidden" name="attribDivName" value="<%=attribDivName%>" />
<h3 class="bMar05">상품속성 정보 <%=chkIIF(attribCd="","등록","수정")%></h3>
<table class="tbType1 listTb">
    <colgroup>
        <col width="90" />
        <col width="*" />
        <col width="90" />
        <col width="*" />
    </colgroup>
    <% if attribCd<>"" then %>
    <tr height="26">
        <td bgcolor="#EEEEEE">상품속성코드</td>
        <td colspan="3" class="lt">
            <%=attribCd %>
        </td>
    </tr>
    <% end if %>
    <tr>
        <td bgcolor="#EEEEEE">상품속성구분</td>
        <td colspan="3" class="lt">
        <% if attribCd<>"" then %>
        <%=attribDivName%>
        <% else %>
            <%=getAttribDivSelectbox("attribDiv",attribDiv,"","onchange='document.frmSub.attribDivName.value=this.options[this.selectedIndex].text'")%>
            <input type="checkbox" id="newDiv" name="newDiv" value="Y" onclick="fnNewAttrDiv(this.checked)"><label for="newDiv">새구분 생성</label>
            <p id="lyrAttrDiv" style="display:none;">
                속성구분명: <input type="text" name="newAttrDivName" value="" size="20" maxlength="16" class="text" />
            </p>
        <% end if %>
        </td>
    </tr>
    <tr>
        <td bgcolor="#EEEEEE">상품속성명</td>
        <td colspan="3" class="lt">
            <p>기본 <input type="text" name="attribName" class="text" size="20" maxlength="20" value="<%=attribName%>" /></p>
            <p class="tMar05">추가 <input type="text" name="attribNameAdd" class="text" size="30" maxlength="40" value="<%=attribNameAdd%>" /></p>
        </td>
    </tr>
    <tr>
        <td bgcolor="#EEEEEE">정렬순서</td>
        <td class="lt">
            <input type="text" name="attribSortNo" class="text" size="4" value="<%=attribSortNo%>" />
        </td>
        <td bgcolor="#EEEEEE">사용여부</td>
        <td class="lt">
            <span id="rdoUsing">
            <input type="radio" name="attribUsing" id="rdoUsing1" value="Y" <%=chkIIF(attribUsing="Y","checked","")%>/><label for="rdoUsing1">사용</label><input type="radio" name="attribUsing" id="rdoUsing2" value="N" <%=chkIIF(attribUsing="N","checked","")%>/><label for="rdoUsing2">삭제</label>
            </span>
        </td>
    </tr>
    <% IF attribCd <>"" THEN %>
        <th colspan="4">MOBLIE</th>
        <tr>
            <td colspan="2" bgcolor="#EEEEEE">모바일 이미지1</td>
            <td colspan="2" class="lt">
                <img id="showMobileImage1" src="<%=mobile_image1%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                <input type="file" id="mobile_image1" value="" onchange="change_image('mobile_image1', this)"/>
                <input type="text" name="mobile_image1" value="<%=mobile_image1%>"/>
            </td>
        </tr>
        <tr>
            <td colspan="2" bgcolor="#EEEEEE">모바일 이미지2</td>
            <td colspan="2" class="lt">
                <img id="showMobileImage2" src="<%=mobile_image2%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                <input type="file" id="mobile_image2" value="" onchange="change_image('mobile_image2', this)"/>
                <input type="text" name="mobile_image2" value="<%=mobile_image2%>"/>
            </td>
        </tr>
        <tr>
            <td colspan="2" bgcolor="#EEEEEE">모바일 이미지3</td>
            <td colspan="2" class="lt">
                <img id="showMobileImage3" src="<%=mobile_image3%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                <input type="file" id="mobile_image3" value="" onchange="change_image('mobile_image3', this)"/>
                <input type="text" name="mobile_image3" value="<%=mobile_image3%>"/>
            </td>
        </tr>
        <tr>
            <td colspan="2" bgcolor="#EEEEEE">모바일 이미지4</td>
            <td colspan="2" class="lt">
                <img id="showMobileImage4" src="<%=mobile_image4%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                <input type="file" id="mobile_image4" value="" onchange="change_image('mobile_image4', this)"/>
                <input type="text" name="mobile_image4" value="<%=mobile_image4%>"/>
            </td>
        </tr>
        <tr>
            <td colspan="2" bgcolor="#EEEEEE">모바일 이미지5</td>
            <td colspan="2" class="lt">
                <img id="showMobileImage5" src="<%=mobile_image5%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                <input type="file" id="mobile_image5" value="" onchange="change_image('mobile_image5', this)"/>
                <input type="text" name="mobile_image5" value="<%=mobile_image5%>"/>
            </td>
        </tr>
        <tr>
            <td colspan="2" bgcolor="#EEEEEE">모바일 이미지6</td>
            <td colspan="2" class="lt">
                <img id="showMobileImage6" src="<%=mobile_image6%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
                <input type="file" id="mobile_image6" value="" onchange="change_image('mobile_image6', this)"/>
                <input type="text" name="mobile_image6" value="<%=mobile_image6%>"/>
            </td>
        </tr>
        <th colspan="4">PC</th>
        	<tr>
        		<td colspan="2" bgcolor="#EEEEEE">PC 이미지1</td>
        		<td colspan="2" class="lt">
        			<img id="showPcImage1" src="<%=pc_image1%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
        			<input type="file" id="pc_image1" value="" onchange="change_image('pc_image1', this)"/>
        			<input type="text" name="pc_image1" value="<%=pc_image1%>"/>
        		</td>
        	</tr>
        	<tr>
        		<td colspan="2" bgcolor="#EEEEEE">PC 이미지2</td>
        		<td colspan="2" class="lt">
        			<img id="showPcImage2" src="<%=pc_image2%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
        			<input type="file" id="pc_image2" value="" onchange="change_image('pc_image2', this)"/>
        			<input type="text" name="pc_image2" value="<%=pc_image2%>"/>
        		</td>
        	</tr>
        	<tr>
        		<td colspan="2" bgcolor="#EEEEEE">PC 이미지3</td>
        		<td colspan="2" class="lt">
        			<img id="showPcImage3" src="<%=pc_image3%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
        			<input type="file" id="pc_image3" value="" onchange="change_image('pc_image3', this)"/>
        			<input type="text" name="pc_image3" value="<%=pc_image3%>"/>
        		</td>
        	</tr>
        	<tr>
        		<td colspan="2" bgcolor="#EEEEEE">PC 이미지4</td>
        		<td colspan="2" class="lt">
        			<img id="showPcImage4" src="<%=pc_image4%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
        			<input type="file" id="pc_image4" value="" onchange="change_image('pc_image4', this)"/>
        			<input type="text" name="pc_image4" value="<%=pc_image4%>"/>
        		</td>
        	</tr>
        	<tr>
        		<td colspan="2" bgcolor="#EEEEEE">PC 이미지5</td>
        		<td colspan="2" class="lt">
        			<img id="showPcImage5" src="<%=pc_image5%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
        			<input type="file" id="pc_image5" value="" onchange="change_image('pc_image5', this)"/>
        			<input type="text" name="pc_image5" value="<%=pc_image5%>"/>
        		</td>
        	</tr>
        	<tr>
        		<td colspan="2" bgcolor="#EEEEEE">PC 이미지6</td>
        		<td colspan="2" class="lt">
        			<img id="showPcImage6" src="<%=pc_image6%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
        			<input type="file" id="pc_image6" value="" onchange="change_image('pc_image6', this)"/>
        			<input type="text" name="pc_image6" value="<%=pc_image6%>"/>
        		</td>
        	</tr>
    <% END IF %>
    <tr>
        <td colspan="4" align="center">
            <input type="button" value=" 취 소 " onClick="self.close()" class="ui-button" style="font-size:11px;"> &nbsp;
            <input type="button" value=" 저 장 " onClick="SaveForm(this.form);" class="ui-button" style="font-size:11px;">
        </td>
    </tr>
</table>
</form>
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->