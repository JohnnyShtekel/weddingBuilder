import React from 'react';

class UploadButtons extends React.Component {
    constructor(props) {
        super(props);

    }


    upload() {
        console.log("bbb");
        var file = document.getElementById("fb").files[0];
        var form_data = new FormData(); // Creating object of FormData class
        form_data.append("file", file); // Appending parameter named file with properties of file_field to form_data

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

    render() {
        return (
                <div className="row">
                    <div className="animated fadeInDownBig">
                        <div className="file-field input-field" dir="rtl">
                            <div className="row">
                                <div className="col s12 ">
                                    <div className="btn col s12">
                                        <span>בחר קובץ</span>
                                        <input type="file" id="fb" multiple/>
                                    </div>
                                </div>


                                <div className="col s12">
                                    <a className="waves-effect waves-light btn-large red right col s12"
                                       ><i
                                        className="material-icons right " onclick={this.upload} >cloud</i>העלה</a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                );
                }
                }


export default UploadButtons