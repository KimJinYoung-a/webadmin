
function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}


function modiContract(ctrkey){
    var popwin = window.open('editContract.asp?ctrkey=' + ctrkey,'editContract','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function dnWebAdm(ctrkey){
    var popwin = window.open('viewContractWeb.asp?ctrkey=' + ctrkey,'preViewContractWeb','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function dnWebAdmDocu(ctrkey){
    var popwin = window.open('viewContractWebDocu.asp?ctrkey=' + ctrkey,'preViewContractWeb','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function dnPdfAdm(iUri){
    var popwin = window.open(iUri,'dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function popOpenContract(groupid){
    var popwin = window.open('openContract.asp?groupid='+groupid,'openContract','width=1300,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function preViewSendContract(groupid,signtype){ 
    var popwin = window.open('preViewSendContract.asp?groupid=' + groupid+'&signtype='+signtype,'preViewSendContract','width=840,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();

}