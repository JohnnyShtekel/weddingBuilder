import React from 'react';
import {Button, Card, Row, Col,Form} from 'react-materialize';
import UploadButtons  from '../js/uploadbutton'
import AutoComplete   from '../js/AutoSerach'
import DragAndDrop   from '../js/draganddrop'


class App extends React.Component {
    constructor(props) {
        super(props);
         this.selector = {};

    }


    render() {
        return (
            <div className="container">
                <div className="center">
                    <div className="animated fadeInDown">
                        <AutoComplete />
                        <UploadButtons/>
                        <div className="row1">
                            <div className="or">
                                <p>או </p>
                            </div>
                        </div>
                        <DragAndDrop/>
                    </div>
                </div>
            </div>
        );
    }
}


export default App