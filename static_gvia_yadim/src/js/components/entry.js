import React from 'react';
import AutoComplete   from '../components/autocomplete';
import Dropzone from   'react-dropzone';
import {Button, Card, Row, Col,Form} from 'react-materialize';
import { PropTypes } from 'react';
import { uploadFile } from '../utils/apis';


class App extends React.Component {
    constructor(props) {
        super(props);
        this.selector = {};

    }

    handleSubmit(e){
        e.preventDefault();
        uploadFile(document.getElementById("file_input").files)
    }

    render() {

        return (
            <Dropzone className="dropZone" activeClassName="activeDropZone" ref="activeClassName" onDrop={ uploadFile } disableClick={true}>
                <div className="container cardContainer animated">
                    <div className="card blue-grey darken-1">
                        <div className="card-content white-text" dir="rtl">
                            <span className="card-title">ברוכים הבאים</span>
                            <div className="orange-text text-lighten-4">
                                <p>ראשית, בחר עובד.</p>
                                <p>לאחר מכן, לחץ על כפתור בחירת הקובץ או גרור את הקובץ לחלון.</p>
                                <p>לבסוף, המתן עד לקבלת הודעה.</p>
                            </div>
                            <AutoComplete />
                            <div className="container fileInputContainer">
                                <div className="file-field input-field">
                                    <div className="btn white black-text container">
                                        <span>בחר קובץ</span>
                                        <input type="file" onChange={this.handleSubmit} id="file_input"/>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div className="card-action" style={{display: 'none'}}>
                            <div className="progress">
                                <div className="indeterminate"></div>
                            </div>
                        </div>
                    </div>
                </div>
            </Dropzone>
        );
    }
}


export default App