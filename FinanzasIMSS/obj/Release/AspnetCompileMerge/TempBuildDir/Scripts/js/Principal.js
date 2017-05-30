$(document).ready(
    function () {
        $("#clsForm").on('click',
            function () {
                $("#fileInputCedLbl").get(0).firstChild.nodeValue = "Selecciona Formato...";
                $("#fileInputEnvLbl").get(0).firstChild.nodeValue = "Selecciona Parámetros...";
                $("#fileInputPdfLbl").get(0).firstChild.nodeValue = "Selecciona Comprobante de Pago...";
                $("#divLoad").html("");
            }
        );

        $("#fileInputCed").on('change',
            function () {
                var data = new FormData();
                var fileInput = $(this)[0];
                if (fileInput.files.length > 0) {
                    for (i = 0; i < fileInput.files.length; i++) {
                        $("#" + $(this).prop('id') + "Lbl").get(0).firstChild.nodeValue = String(fileInput.files[i].name);
                    }
                } else {
                    $("#fileInputCedLbl").get(0).firstChild.nodeValue = "Selecciona Formato...";
                }
            }
        );

        $("#fileInputEnv").on('change',
            function () {
                var data = new FormData();
                var fileInput = $(this)[0];
                if (fileInput.files.length > 0) {
                    for (i = 0; i < fileInput.files.length; i++) {
                        $("#" + $(this).prop('id') + "Lbl").get(0).firstChild.nodeValue = String(fileInput.files[i].name);
                    }
                } else {
                    $("#fileInputEnvLbl").get(0).firstChild.nodeValue = "Selecciona Parámetros...";
                }
            }
        );

        $("#fileInputPdf").on('change',
            function () {
                var data = new FormData();
                var fileInput = $(this)[0];
                if (fileInput.files.length > 0) {
                    for (i = 0; i < fileInput.files.length; i++) {
                        $("#" + $(this).prop('id') + "Lbl").get(0).firstChild.nodeValue = String(fileInput.files[i].name);
                    }
                } else {
                    $("#fileInputPdfLbl").get(0).firstChild.nodeValue = "Selecciona Comprobante de Pago...";
                }
            }
        );

        $('#uploader').on('submit',
            function (e) {
                e.preventDefault();
                AbrirModal(
                    {
                        Mensaje: "<p><h3>Cargando...</h3></p>",
                        AlineacionMensaje: "center"
                    }
                );
                var contador = 0;
                var data = new FormData(); //FormData object
                var fileInput = $('#fileInputCed')[0];
                //Iterating through each files selected in fileInput
                for (i = 0; i < fileInput.files.length; i++) {
                    //Appending each file to FormData object
                    //data.append(fileInput.files[i].name, fileInput.files[i]);
                    data.append("fileInputCed", fileInput.files[i]);
                    contador++;
                }
                fileInput = null;
                fileInput = $('#fileInputEnv')[0];
                for (i = 0; i < fileInput.files.length; i++) {
                    //Appending each file to FormData object
                    //data.append(fileInput.files[i].name, fileInput.files[i]);
                    data.append("fileInputEnv", fileInput.files[i]);
                    contador++;
                }
                fileInput = null;
                fileInput = $('#fileInputPdf')[0];
                for (i = 0; i < fileInput.files.length; i++) {
                    //Appending each file to FormData object
                    //data.append(fileInput.files[i].name, fileInput.files[i]);
                    data.append("fileInputPdf", fileInput.files[i]);
                    contador++;
                }
                if (contador > 0) {
                    //Creating an XMLHttpRequest and sending
                    $.ajax(
                        {
                            dataType: "html",
                            type: "POST",
                            url: url_Upload,
                            contentType: false,
                            processData: false,
                            data: data,
                            success: function (data, textStatus, jqXHR) {
                                if (data != null) {
                                    if (typeof (data) === "string") {
                                        var obj = null;
                                        try {
                                            obj = JSON.parse(data);
                                        } catch (ex) {
                                            obj = null;
                                        } finally {
                                            if (obj == null) {
                                                $("#divLoad").html(data);
                                            } else {
                                                $.toast(
                                                    {
                                                        text: obj.Message, // Text that is to be shown in the toast
                                                        heading: "Error", // Optional heading to be shown on the toast
                                                        icon: 'error', // Type of toast icon
                                                        showHideTransition: 'fade', // fade, slide or plain
                                                        allowToastClose: true, // Boolean value true or false
                                                        hideAfter: 6000, // false to make it sticky or number representing the miliseconds as time after which toast needs to be hidden
                                                        stack: 5, // false if there should be only one toast at a time or a number representing the maximum number of toasts to be shown at a time
                                                        position: 'bottom-right', // bottom-left or bottom-right or bottom-center or top-left or top-right or top-center or mid-center or an object representing the left, right, top, bottom values
                                                        textAlign: 'center',  // Text alignment i.e. left, right or center
                                                        loader: false // Whether to show loader or not. True by default
                                                    }
                                                );
                                            }
                                        }
                                    } else if (typeof (data) === "object") {
                                        $.toast(
                                            {
                                                text: data.Message, // Text that is to be shown in the toast
                                                heading: "Error", // Optional heading to be shown on the toast
                                                icon: 'error', // Type of toast icon
                                                showHideTransition: 'fade', // fade, slide or plain
                                                allowToastClose: true, // Boolean value true or false
                                                hideAfter: 6000, // false to make it sticky or number representing the miliseconds as time after which toast needs to be hidden
                                                stack: 5, // false if there should be only one toast at a time or a number representing the maximum number of toasts to be shown at a time
                                                position: 'bottom-right', // bottom-left or bottom-right or bottom-center or top-left or top-right or top-center or mid-center or an object representing the left, right, top, bottom values
                                                textAlign: 'center',  // Text alignment i.e. left, right or center
                                                loader: false // Whether to show loader or not. True by default
                                            }
                                        );
                                    } else {
                                        $.toast(
                                            {
                                                text: "Ocurrió un error al intentar obtener la información.", // Text that is to be shown in the toast
                                                heading: "Error", // Optional heading to be shown on the toast
                                                icon: 'error', // Type of toast icon
                                                showHideTransition: 'fade', // fade, slide or plain
                                                allowToastClose: true, // Boolean value true or false
                                                hideAfter: 6000, // false to make it sticky or number representing the miliseconds as time after which toast needs to be hidden
                                                stack: 5, // false if there should be only one toast at a time or a number representing the maximum number of toasts to be shown at a time
                                                position: 'bottom-right', // bottom-left or bottom-right or bottom-center or top-left or top-right or top-center or mid-center or an object representing the left, right, top, bottom values
                                                textAlign: 'center',  // Text alignment i.e. left, right or center
                                                loader: false // Whether to show loader or not. True by default
                                            }
                                        );
                                    }
                                } else {
                                    $.toast(
                                        {
                                            text: "Ocurrió un error al intentar obtener la información.", // Text that is to be shown in the toast
                                            heading: "Error", // Optional heading to be shown on the toast
                                            icon: 'error', // Type of toast icon
                                            showHideTransition: 'fade', // fade, slide or plain
                                            allowToastClose: true, // Boolean value true or false
                                            hideAfter: 6000, // false to make it sticky or number representing the miliseconds as time after which toast needs to be hidden
                                            stack: 5, // false if there should be only one toast at a time or a number representing the maximum number of toasts to be shown at a time
                                            position: 'bottom-right', // bottom-left or bottom-right or bottom-center or top-left or top-right or top-center or mid-center or an object representing the left, right, top, bottom values
                                            textAlign: 'center',  // Text alignment i.e. left, right or center
                                            loader: false // Whether to show loader or not. True by default
                                        }
                                    );
                                }
                                CerrarModal();
                            },
                            error: function (jqXHR, textStatus, errorThrown) {
                                $.toast(
                                    {
                                        text: errorThrown + "\n(" + JSON.stringify(jqXHR, null, 4) + ")", // Text that is to be shown in the toast
                                        heading: textStatus, // Optional heading to be shown on the toast
                                        icon: 'error', // Type of toast icon
                                        showHideTransition: 'fade', // fade, slide or plain
                                        allowToastClose: true, // Boolean value true or false
                                        hideAfter: 6000, // false to make it sticky or number representing the miliseconds as time after which toast needs to be hidden
                                        stack: 5, // false if there should be only one toast at a time or a number representing the maximum number of toasts to be shown at a time
                                        position: 'bottom-right', // bottom-left or bottom-right or bottom-center or top-left or top-right or top-center or mid-center or an object representing the left, right, top, bottom values
                                        textAlign: 'center',  // Text alignment i.e. left, right or center
                                        loader: false // Whether to show loader or not. True by default
                                    }
                                );
                                CerrarModal();
                            }
                        }
                    );
                } else {
                    $.toast(
                        {
                            text: "Favor de Seleccionar un Archivo.", // Text that is to be shown in the toast
                            heading: "Advertencia", // Optional heading to be shown on the toast
                            icon: 'warning', // Type of toast icon
                            showHideTransition: 'fade', // fade, slide or plain
                            allowToastClose: true, // Boolean value true or false
                            hideAfter: 6000, // false to make it sticky or number representing the miliseconds as time after which toast needs to be hidden
                            stack: 5, // false if there should be only one toast at a time or a number representing the maximum number of toasts to be shown at a time
                            position: 'bottom-right', // bottom-left or bottom-right or bottom-center or top-left or top-right or top-center or mid-center or an object representing the left, right, top, bottom values
                            textAlign: 'center',  // Text alignment i.e. left, right or center
                            loader: false // Whether to show loader or not. True by default
                        }
                    );
                    CerrarModal();
                }
            }
        );
    }
);