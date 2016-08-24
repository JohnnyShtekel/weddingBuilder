import React from 'react';
import {Button, Card, Row, Col,Form} from 'react-materialize';
import UploadButtons  from '../components/uploadbutton'
import AutoComplete   from '../components/autoserach'
import DragAndDrop   from '../components/draganddrop'


class Entry extends React.Component {
    constructor(props) {
        super(props);
        this.selector = {};

    }


    render() {
        return (
            <div className="container">

                <div className="row">
                    <AutoComplete/>
                </div>
                <div className="row">

                </div>
            </div>



        );
    }
}


export default Entry