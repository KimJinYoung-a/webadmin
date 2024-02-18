Vue.component('DATE-PICKER', {
    template : `<input :id="id" type="text" :value="date" class="date" readonly>`,
    props : {
        id : { type:String, default:'datepicker' }, // ID
        date : { type:String, default:'' }, // ���� �ʱⰪ
    },
    mounted() {
        const _this = this;
        $(this.$el).datepicker({
            dateFormat: "yy-mm-dd",
            monthNames: ['1��', '2��', '3��', '4��', '5��', '6��', '7��', '8��', '9��', '10��', '11��', '12��'],
            monthNamesShort: ['1��', '2��', '3��', '4��', '5��', '6��', '7��', '8��', '9��', '10��', '11��', '12��'],
            dayNames: ['��', '��', 'ȭ', '��', '��', '��', '��'],
            dayNamesShort: ['��', '��', 'ȭ', '��', '��', '��', '��'],
            dayNamesMin: ['��', '��', 'ȭ', '��', '��', '��', '��'],
            changeMonth: true, // ������ select box ǥ�� (�⺻�� false)
            changeYear: true, // �⼱�� selectbox ǥ�� (�⺻�� false)
            showMonthAfterYear: true, // �����⵵ �� ���̱�
            showOtherMonths: true, // �ٸ� �� �޷¿� ���̱�
            selectOtherMonths: true, // �ٸ� �� �޷¿� ���̴°� Ŭ�� �����ϰ� �ϱ�
            onSelect(date){
                _this.$emit('updateDate',date);
            }
        });
    },
    beforeDestroy(){
        $(this.$el).datepicker('hide').datepicker('destroy')
    },
});