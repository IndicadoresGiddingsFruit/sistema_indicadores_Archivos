$('#ModalLiberar').on('shown.bs.modal', function () {
    $('#myInput').focus()
})

function getDataAjax(id, action) {
    $.ajax({
        type: "POST",
        url: action,
        data: { id },
        success: function (response) {
            console.log(response);
        }
    });
}