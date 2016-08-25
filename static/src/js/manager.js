import React from 'react';
import ReactDOM from 'react-dom';
import App from '../js/entry';
import Nav from '../js/nav';

class Manager extends React.Component {
    constructor(props) {
        super(props);

    }


    render() {
        return (
            <div>
                <Nav/>
                <App/>
            </div>
        )

    }

}


ReactDOM
    .render(<Manager/>, document.getElementById(
        'app'
    ))
;
