import React from 'react';
import { PropTypes } from 'react'


class UploadButtons extends React.Component {
    constructor(props) {
        super(props);

    }


    handleSubmit(e) {
        e.preventDefault();
        if (document.getElementById("autocomplete-input").value != "") {
            if (document.getElementById('fb').files.length != 0) {
                var workerName = document.getElementById("autocomplete-input").value;
                var file = document.getElementById("fb").files[0];
                var form_data = new FormData(); // Creating object of FormData class
                form_data.append("file", file); // Appending parameter named file with properties of file_field to form_data
                form_data.append("worker", workerName);
                document.getElementById("loader").style.display = "block";
                document.getElementById("loadertext").style.display = "block";


                $.ajax({
                    type: 'POST',
                    url: '/upload',
                    data: form_data,
                    contentType: false,
                    cache: false,
                    processData: false,
                    async: false,
                    success: function (data) {
                        document.getElementById("loader").style.display = "none";
                        document.getElementById("loadertext").style.display = "none";
                        Materialize.toast('עדכון קובץ בוצע בהצלחה', 4000);
                    },
                     error :function (data) {
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

    render() {
        return (
            <div className="row">
                <form id="form" onSubmit={(e)=>this.handleSubmit(e)}>
                    <div className="animated fadeInDownBig">
                        <div className="file-field input-field" dir="rtl">
                            <div className="row">
                                <div className="col s12 ">
                                    <div className="btn col s12">
                                        <span>בחר קובץ</span>
                                        <input type="file" id="fb" multiple/>
                                    </div>
                                </div>


                                <div className="col s12" onClick={(e)=>this.handleSubmit(e)}>
                                    <a className="waves-effect waves-light btn-large red right col s12"
                                    ><i
                                        className="material-icons right " onClick={(e)=>this.handleSubmit(e)}>cloud</i>העלה</a>
                                </div>
                            </div>
                        </div>
                    </div>
                </form>
            </div>

        );
    }
}


export default UploadButtons