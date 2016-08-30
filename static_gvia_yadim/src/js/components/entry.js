import React from 'react';
import {Button, Card, Row, Col,Form} from 'react-materialize';
import UploadButtons  from '../components/uploadbutton'
import AutoComplete   from '../components/autocomplete'
import DragAndDrop   from '../components/draganddrop'
import { PropTypes } from 'react'


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
                        <UploadButtons onUpdate={this.onUpdate}/>
                        <div className="row1">
                            <div className="or" id="or">
                                <p>או </p>
                            </div>
                        </div>
                        <DragAndDrop onUpdate={this.onUpdate}/>
                    </div>
                </div>
            </div>
        );
    }
}


export default App