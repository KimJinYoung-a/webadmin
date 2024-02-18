// front-scm-js

// 슬라이드 관리
function popSlideManage(m,menutitle) {
	var popSlideManage = window.open('/admin/common/lib/pop_slide/pop_slide_manage_list.asp?mastercode='+ m +'&menu='+menutitle,'popSlideManage','width=1024,height=768,resizable=yes,scrollbars=yes')
	popSlideManage.focus();
}

// 슬라이드 미리보기
function popSlideView(m,d,menutitle) {
	var popSlideView = window.open('/admin/common/lib/pop_slide/pop_slide_preview.asp?mastercode='+ m +'&detailcode='+ d +'&menu='+menutitle,'popSlideView','width=1024,height=600,resizable=yes,scrollbars=yes')
	popSlideView.focus();
}