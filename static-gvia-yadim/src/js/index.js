import React from 'react';
import ReactDOM from 'react-dom';
import App from '../js/components/entry';
import Nav from '../js/components/nav'

class Manager extends React.Component {
    constructor(props) {
        super(props);

    }


    render() {
        return (
            <div id="app-page">
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
