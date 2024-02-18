Vue.component('DATE-PICKER', {
    template : `<input :id="id" type="text" :value="date" class="date" readonly>`,
    props : {
        id : { type:String, default:'datepicker' }, // ID
        date : { type:String, default:'' }, // 일자 초기값
    },
    mounted() {
        const _this = this;
        $(this.$el).datepicker({
            dateFormat: "yy-mm-dd",
            monthNames: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
            monthNamesShort: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
            dayNames: ['일', '월', '화', '수', '목', '금', '토'],
            dayNamesShort: ['일', '월', '화', '수', '목', '금', '토'],
            dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
            changeMonth: true, // 월선택 select box 표시 (기본은 false)
            changeYear: true, // 년선택 selectbox 표시 (기본은 false)
            showMonthAfterYear: true, // 다음년도 월 보이기
            showOtherMonths: true, // 다른 월 달력에 보이기
            selectOtherMonths: true, // 다른 월 달력에 보이는거 클릭 가능하게 하기
            onSelect(date){
                _this.$emit('updateDate',date);
            }
        });
    },
    beforeDestroy(){
        $(this.$el).datepicker('hide').datepicker('destroy')
    },
});