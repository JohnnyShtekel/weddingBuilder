import React from 'react';
import App from '../js/app'
import UploadPage from '../js/components/entry'
import WelcomePage from './components/welcomepage'
import DailyDepartmentReport from './components/dailydepartmentreport'
import Route from 'react-router/lib/Route';
import IndexRoute from 'react-router/lib/IndexRoute';
import Router from 'react-router/lib/Router';


class Root extends React.Component {

    constructor(props) {
        super(props)


    }

    render() {
        const {history} = this.props;
        return (
            <Router history={history}>
                <Route path="/api/v1/gvia-yadim-report/" component={App}>
                    <IndexRoute component={WelcomePage}/>
                    <Route path="/api/v1/gvia-yadim-report/upload-page/" component={UploadPage}/>
                    <Route path="/api/v1/gvia-yadim-report/department-report/" component={DailyDepartmentReport}/>
                </Route>
            </Router>
        )
    }
}

export default Root;