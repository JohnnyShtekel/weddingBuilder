import React from 'react';
import ReactDOM from 'react-dom'
import Root from '../js/routes';
import createBrowserHistory from 'history/lib/createBrowserHistory';


const history = createBrowserHistory();

const app = document.getElementById('app');

ReactDOM.render(<Root history={history}/>, app);