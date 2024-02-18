<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
'###########################################################
' Description : 카드 할인 정보 편집
' History : 2023.01.30 원승현 생성
'###########################################################

session.codePage = 65001		'세션코드 UTF-8 강제 설정
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/pgPromotionCls.asp"-->
<%
Dim idx : idx = requestCheckVar(getNumeric(request("idx")),10)
Dim sDt, eDt
Dim pgprogbn, cardcd, cimage, isusing, conts, regdate, contlink

Dim oCardPromo
SET oCardPromo= new CCardPromotion
oCardPromo.FRectIdx=idx
if (idx<>"") then
oCardPromo.getCardPromotionOne
end if

if oCardPromo.FResultCount>0 then
    cimage = oCardPromo.FOneItem.Fcimage
    pgprogbn = oCardPromo.FOneItem.Fpgprogbn
    cardcd = oCardPromo.FOneItem.FCardCd
    sDt = Left(oCardPromo.FOneItem.FSDt,10)
    eDt = Left(oCardPromo.FOneItem.FEDt,10)
    conts = db2html(oCardPromo.FOneItem.Fconts)
    contlink = ReplaceBracket(oCardPromo.FOneItem.Fcontlink)
    isusing = oCardPromo.FOneItem.FIsUsing
    regdate = oCardPromo.FOneItem.FRegDate
end if
SET oCardPromo= Nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
    function jsConfirmSm(){
        var frm = document.cardSalereg;

        if (frm.sDt.value.length<1){
            alert('시작일을 입력 하세요.');
            frm.sDt.focus();
            return false;
        }

        if (frm.eDt.value.length<1){
            alert('종료일을 입력 하세요.');
            frm.eDt.focus();
            return false;
        }

        if (frm.conts.value.length<1){
            alert('내용을 입력 하세요.');
            frm.conts.focus();
            return false;
        }

        if (confirm('저장 하시겠습니까?')){
            return true;
        }
    }
</script>
</head>
<body>
<form name="cardSalereg" method="post" action="/admin/sitemaster/pg/pop_RegPgPromotion_process.asp" onSubmit="return jsConfirmSm();" >
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="usinghtml" value="Y" />
    <table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
        <tr>
            <td colspan="2"><!--//코드 등록 및 수정-->
                <table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
                    <tr>
                        <td colspan="4"><strong>카드 할인정보 등록/편집</strong></td>
                    </tr>        
                    <% IF idx <> "" THEN%>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">코드번호</td>
                        <td bgcolor="#FFFFFF"><%=idx%></td>
                    </tr>
                    <% end if %>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">기간</td>
                        <td bgcolor="#FFFFFF">
                        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
                        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
                        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
                        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
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
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">내용</td>
                        <td bgcolor="#FFFFFF">
                            <textarea name="conts" rows="15" class="textarea" style="width:100%"><%= conts %></textarea>
                            <script>
                            //
                            window.onload = new function(){
                                var cardSaleContEditor = CKEDITOR.replace('conts',{
                                    height : 600,
                                    // 업로드된 파일 목록
                                    //filebrowserBrowseUrl : '/browser/browse.asp',
                                    // 파일 업로드 처리 페이지
                                    //filebrowserUploadUrl : '파일업로드'
                                    filebrowserImageUploadUrl : '',
                                    customConfig: '/js/ckeditor/config_cardInfo.js'
                                });
                                cardSaleContEditor.on( 'change', function( evt ) {
                                    // 입력할 때 textarea 정보 갱신
                                    //document.cardSalereg.conts.value = evt.editor.getData();
                                });
                            }
                            </script>
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">사용여부</td>
                        <td bgcolor="#FFFFFF">
                        <input type="radio" name="isusing" value="Y" <%=CHKIIF(isusing="Y" or isusing="","checked","")%> >사용
                        <input type="radio" name="isusing" value="N" <%=CHKIIF(isusing="N" ,"checked","")%> >사용안함
                        </td>
                    </tr>            
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
                <tr>
                    <td align="left"><a href="javascript:self.close()"><img src="/images/icon_cancel.gif" border="0"></a></td>
                    <td align="right"><input type="image" src="/images/icon_save.gif"></td>
                </tr>
                </table>
            </td>
        </tr>    
    </table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<% 
session.codePage = 949		'세션코드 EUC-KR 원복
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->