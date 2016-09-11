// import '../../vendor/materialize/dist/js/materialize'
import React from 'react';
import ReactDOM from 'react-dom';
import { render } from 'react-dom';
import createBrowserHistory from 'history/lib/createBrowserHistory';

import App from '../js/components/entry'
import Nav from '../js/components/nav'
import DailyDepartmentReport from '../js/components/dailydepartmentreport'


class Root extends React.Component {


    render() {
        return (
            <div id="app-id">
                <Nav/>
                <DailyDepartmentReport />
            </div>
        );
    }
}




ReactDOM.render(<Root/>, document.getElementById('app'));