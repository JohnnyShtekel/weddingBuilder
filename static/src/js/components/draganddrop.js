import React from 'react';
import Dropzone from   'react-dropzone'


class DragAndDrop extends React.Component {
    constructor(props) {
        super(props);
         this.state = {
            files: []
        };

    }


   onDrop() {
        this.setState({
            files: files,

        });
    }

    onOpenClick() {
        this.refs.dropzone.open();
    }


    onDrop(files) {
        var file = files[0];
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
            <div className="row2">
                <div className="animated fadeInRight" id="drop">
                    <Dropzone ref="dropzone" onDrop={this.onDrop}>
                        <div>נא להשליך קובץ שברצונך לעלות או ללחוץ על הריבוע ולבחור קובץ.</div>
                        <i className="large material-icons">system_update_alt</i>
                    </Dropzone>
                </div>
            </div>

        );
    }
}


export default DragAndDrop