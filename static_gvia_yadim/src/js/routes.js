import React from 'react';
import {Nav} from '../js/components/nav'
import {UploadPage} from '../js/components/entry'
import {WelcomePage} from '../js/components/welcomepage'
import Route from 'react-router/lib/Route';
import IndexRoute from 'react-router/lib/IndexRoute';
import Router from 'react-router/lib/Router';
import {browserHistory} from 'react-router/lib/browserHistory'


class Root extends React.Component {

    constructor(props) {
        super(props)


    }

    render() {
        return (
            <Router history={browserHistory}>
                <Route path="/api/v1/gvia-yadim-report/" component={Nav}>
                    <IndexRoute component={WelcomePage}/>
                    <Route path="/api/v1/gvia-yadim-report/upload-page/" component={UploadPage}/>
                </Route>
            </Router>
        )
    }
}

export default Root;