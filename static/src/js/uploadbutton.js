import React from 'react';
class UploadButtons extends React.Component {
    constructor(props) {
        super(props);

    }


    handleSubmit(e) {
         e.preventDefault();
        if(document.getElementById("autocomplete-input").value != "") {
            if(document.getElementById('fb').files.length != 0) {
                var workerName = document.getElementById("autocomplete-input").value;
                var file = document.getElementById("fb").files[0];
                var form_data = new FormData(); // Creating object of FormData class
                form_data.append("file", file); // Appending parameter named file with properties of file_field to form_data
                form_data.append("worker", workerName);
                $.ajax({
                    type: 'POST',
                    url: '/upload',
                    data: form_data,
                    contentType: false,
                    cache: false,
                    processData: false,
                    async: false,
                    success: function (data) {
                        console.log('Success!');
                    },
                });
            }
            else
            {
                 Materialize.toast('נא לטעון קובץ', 4000);
            }
        }
        else
        {
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


                                <div className="col s12" onClick={(e)=>this.handleSubmit(e)}  >
                                    <a className="waves-effect waves-light btn-large red right col s12"
                                    ><i
                                        className="material-icons right " onClick={(e)=>this.handleSubmit(e)} >cloud</i>העלה</a>
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