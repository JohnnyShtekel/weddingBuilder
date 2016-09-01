export function uploadFile(files) {
    if (document.getElementById("autocomplete-input").value != "") {
        if (files.length != 0) {
            var workerName = document.getElementById("autocomplete-input").value;
            var file = files[0];
            var form_data = new FormData(); // Creating object of FormData class
            form_data.append("file", file); // Appending parameter named file with properties of file_field to form_data
            form_data.append("worker", workerName);
            document.getElementsByClassName('card-action')[0].style.display = "block";

            $.ajax({
                type: 'POST',
                url: '/api/v1/gvia-yadim-report/upload/',
                data: form_data,
                contentType: false,
                cache: false,
                processData: false,
                success: function (data) {
                    document.getElementsByClassName('card-action')[0].style.display = "none";
                    Materialize.toast('עדכון קובץ בוצע בהצלחה', 4000);
                },
                error :function (data) {
                    document.getElementsByClassName('card-action')[0].style.display = "none";
                    Materialize.toast('פעולה נכשלה נא לבדוק את הקובץ שהכונס למערכת', 4000);
                }
            });
        }
        else {
            Materialize.toast('נא לטעון קובץ', 4000);
        }
    }
    else {
        Materialize.toast('נא לבחור שם עובד', 4000);
    }

}