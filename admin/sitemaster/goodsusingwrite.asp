<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
response.write "사용중지메뉴" ''사장님 지시 2013/01/31 사용을 원할경우 사장님께 문의
dbget.Close() : response.end
%>
<script language="javascript">
function SwapCate(cdl){
	document.evaluateFrm.cdL.value=cdl;
	document.evaluateFrm.submit();

}
function changefile(iimage,comp){
	var tmpD 	= document.getElementById(comp);
	tmpD.checked=false;
	
}

function delimage(ifile,iimage,comp){
	var tmpT		= document.getElementById(ifile);
	var tmpV 	= document.getElementById(iimage);

	if (comp.checked){
		tmpV.style.display="none";
	}
}

function checkImageSuffix (fileInput) {
   var suffixPattern = /(gif|jpg|jpeg|png)$/i;
   if (!suffixPattern.test(fileInput.value)) {
     alert('GIF,JEPG,PNG 파일만 가능합니다.');
     fileInput.focus();
     fileInput.select();
     return false;
   }
   return true;
}


function SubmitForm(frm)
{
    if (frm.usedcontents.value == "") {
            alert("상품평을 적어주세요.");
            frm.usedcontents.focus();
            return;
    }

    if ((frm.totPnt[0].checked)||(frm.totPnt[1].checked)||(frm.totPnt[2].checked)||(frm.totPnt[3].checked)){

    }else{
    	alert("총평을 선택해 주세요.");
    	frm.totPnt[3].focus();
    	return;
    };

    if ((frm.funPnt[0].checked)||(frm.funPnt[1].checked)||(frm.funPnt[2].checked)||(frm.funPnt[3].checked)){

    }else{
    	alert("기능평을 선택해 주세요.");
    	frm.funPnt[3].focus();
    	return;
    };

    if ((frm.dgnPnt[0].checked)||(frm.dgnPnt[1].checked)||(frm.dgnPnt[2].checked)||(frm.dgnPnt[3].checked)){

    }else{
    	alert("디자인평을 선택해 주세요.");
    	frm.dgnPnt[3].focus();
    	return;
    };

    if ((frm.prcPnt[0].checked)||(frm.prcPnt[1].checked)||(frm.prcPnt[2].checked)||(frm.prcPnt[3].checked)){

    }else{
    	alert("가격평을 선택해 주세요.");
    	frm.prcPnt[3].focus();
    	return;
    };

    if ((frm.stfPnt[0].checked)||(frm.stfPnt[1].checked)||(frm.stfPnt[2].checked)||(frm.stfPnt[3].checked)){

    }else{
    	alert("만족도평을 선택해 주세요.");
    	frm.stfPnt[3].focus();
    	return;
    };

	if ((frm.file1.value.length>0)&&(!checkImageSuffix(frm.file1))){
		return;
	};

	if ((frm.file2.value.length>0)&&(!checkImageSuffix(frm.file2))){
		return;
	};

	if (frm.file1.value.length>0){
		if ((frm.file1.fileSize>1024000)||(frm.file1.fileSize<1)){
			alert('파일사이즈가 너무 크거나 사용할 수 없습니다. 최대 1M까지 가능');
			frm.file1.focus();
			frm.file1.select();
			return;
		}

		if (frm.file1_image.width>600){
			alert('이미지 사이즈는 가로 600까지 가능합니다.');
			frm.file1.focus();
			frm.file1.select();
			return;
		}
	}

	if (frm.file2.value.length>0){
		if ((frm.file2.fileSize>1024000)||(frm.file2.fileSize<1)){
			alert('파일사이즈가 너무 크거나 사용할 수 없습니다. 최대 1M까지 가능');
			frm.file2.focus();
			frm.file2.select();
			return;
		}

		if (frm.file2_image.width>600){
			alert('이미지 사이즈는 가로 600까지 가능합니다.');
			frm.file2.focus();
			frm.file2.select();
			return;
		}
	}

    if (confirm("입력사항이 정확합니까?") == true) { frm.submit(); }
}
</script>
<link rel=stylesheet type="text/css" href="http://www.10x10.co.kr/lib/css/tenten.css">


<table width="750" border="0" cellspacing="0" cellpadding="0">
<form name="FrmGoodusing" method="post" action="<%=uploadUrl%>/linkweb/doevaluatewithimage2.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="itemoption" value="0000" />
<input type="hidden" name="orderserial" value="000000000" />
  <tr>
    <td valign="top" align="center">
<table width="700" border="0" cellpadding="0" cellspacing="1" bgcolor="9E9E9E" height="200">
  <tr>
    <td>
      <table width="700" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td bgcolor="#FFFFFF">
            <table width="700" border="0" cellpadding="0" cellspacing="0" height="30" align="center">
              <tr>
                <td width="100" bgcolor="F9F9F9" height="30" align="center">아이템번호</td>
					 <td bgcolor="F9F9F9" height="30"><input type="text" name="itemid" size="15" class="input_01"></td>
              </tr>
              <tr>
                <td width="100" bgcolor="F9F9F9" height="30" align="center">사용자 아이디</td>
					 <td bgcolor="F9F9F9" height="30"><input type="text" name="userid" size="30" maxlength="32" class="input_01">
					 (정확히 입력!)
					 </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td bgcolor="D1D1D1"></td>
        </tr>
        <input type="hidden" name="dummi" value="">
        <input type="hidden" name="mode" value="write">
		<input type="hidden" name="backpath" value="http://webadmin.10x10.co.kr/admin/sitemaster/goodsusingwrite.asp">
        <tr>
          <td bgcolor="#FFFFFF">
            <table width="700" border="0" cellpadding="0" cellspacing="0" height="30" align="center">
              <tr>
                <td width="100" bgcolor="F9F9F9" height="30">
                  <div align="center">상품평</div>
                </td>
                <td bgcolor="#eeeeee" width="1" height="30"></td>
                <td style="padding:3 6 3 6" height="30">
                  <textarea name="usedcontents" cols="84" rows="10" class="input_01"></textarea>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td bgcolor="D1D1D1"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <table width="700" border="0" cellpadding="0" cellspacing="0" height="30" align="center">
              <tr>
                <td width="100" bgcolor="F9F9F9" height="30">
                  <div align="center">첨부 이미지1</div>
                </td>
                <td bgcolor="#eeeeee" width="1" height="30"></td>
                <td style="padding:3 6 3 6" height="30">
                  <font color="#666666">파일 크기는1MB이하, JPG 또는 GIF형식의 파일만 가능합니다.<br>
					  사이즈는 width 600 이하로 설정해 주시기 바랍니다.</font><br>
                  <input type="file" name="file1" size="30" class="input_01" >
                  <input type=hidden name=file1_del>
                  <br>
                  <img name=file1_image src="" style="display:none">
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td bgcolor="D1D1D1"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <table width="700" border="0" cellpadding="0" cellspacing="0" height="30" align="center">
              <tr>
                <td width="100" bgcolor="F9F9F9" height="30">
                  <div align="center">첨부 이미지2</div>
                </td>
                <td bgcolor="#eeeeee" width="1" height="30"></td>
                <td style="padding:3 6 3 6" height="30">
                  <font color="#666666">파일 크기는1MB이하, JPG 또는 GIF형식의 파일만 가능합니다.<br>
					  사이즈는 width 600 이하로 설정해 주시기 바랍니다.</font><br>
                  <input type="file" name="file2" size="30" class="input_01" >
                  <input type=hidden name=file2_del>
                  <br>
                  <img name=file2_image src="" style="display:none">
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td bgcolor="D1D1D1"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <table width="700" border="0" cellpadding="0" cellspacing="0" height="30" align="center">
              <tr>
                <td width="100" bgcolor="F9F9F9" height="30">
                  <div align="center">만족도평가</div>
                </td>
                <td bgcolor="#eeeeee" width="1" height="30"></td>
                <td style="padding:3 6 3 6" height="30">
                  <table width="500" border="0" cellspacing="3" cellpadding="0">
                    <tr>
                      <td >
                        <table border="0" cellpadding="0" cellspacing="0" width="331">
                          <tr>
                            <td width="50" height="11">&nbsp;</td>
                            <td width="65" height="11">
                              <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"></div>
                            </td>
                            <td width="65" height="11">
                              <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"></div>
                            </td>
                            <td width="65" height="11">
                              <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"></div>
                            </td>
                            <td width="65" height="11">
                              <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"><img src="http://fiximage.10x10.co.kr/web2007/my1010/step.gif" width="9" height="9"></div>
                            </td>
                          </tr>
                          <tr>
                            <td width="50" height="20"><b>총평</b></td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="totPnt" value="1">1</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="totPnt" value="2">2</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="totPnt" value="3">3</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="totPnt" value="4">4</div>
                            </td>
                          </tr>
                          <tr>
                            <td colspan="6" bgcolor="#CCCCCC">
                              <div align="center"><img src="http://www.10x10.co.kr/images/my10x10/spacer.gif" width="100%" height="1"></div>
                            </td>
                          </tr>
                          <tr>
                            <td width="50" height="20">기능</td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="funPnt" value="1">1</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="funPnt" value="2">2</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="funPnt" value="3">3</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="funPnt" value="4">4</div>
                            </td>
                          </tr>
                          <tr>
                            <td colspan="6" bgcolor="#CCCCCC">
                              <div align="center"><img src="http://www.10x10.co.kr/images/my10x10/spacer.gif" width="100%" height="1"></div>
                            </td>
                          </tr>
                          <tr>
                            <td width="50" height="20">디자인</td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="dgnPnt" value="1">1</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="dgnPnt" value="2">2</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="dgnPnt" value="3">3</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"> <input type="radio" name="dgnPnt" value="4">4</div>
                            </td>
                          </tr>
                          <tr>
                            <td colspan="6" bgcolor="#CCCCCC"><img src="http://www.10x10.co.kr/images/my10x10/spacer.gif" width="100%" height="1"></td>
                          </tr>
                          <tr>
                            <td width="50" height="20">가격</td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="prcPnt" value="1">1</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="prcPnt" value="2">2</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="prcPnt" value="3">3</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="prcPnt" value="4">4</div>
                            </td>
                          </tr>
                          <tr>
                            <td colspan="6" bgcolor="#CCCCCC"><img src="http://www.10x10.co.kr/images/my10x10/spacer.gif" width="100%" height="1"></td>
                          </tr>
                          <tr>
                            <td width="50" height="20">만족도</td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="stfPnt" value="1">1</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="stfPnt" value="2">2</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="stfPnt" value="3">3</div>
                            </td>
                            <td width="65" height="20" class="verdana-basic">
                              <div align="center"><input type="radio" name="stfPnt" value="4">4</div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="700" border="0" align="left" height="55" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <div align="center">
      <a href="javascript:SubmitForm(FrmGoodusing)"><font color="#000000"><img src="http://fiximage.10x10.co.kr/web2007/my1010/s6_pwbtn.gif"></font></a>
      <a href="javascript:history.back();"><font color="#000000"><img src="http://fiximage.10x10.co.kr/web2007/my1010/s6_pwcan_btn.gif" border="0"></font></a>
      </div>
    </td>
  </tr>
</table>
	</td>
  </tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
