import React from 'react';
import * as quries from '../utils/queries'
import { autoComplete } from '../utils/auto_complete';

class AutoComplete extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            workers: {}
        };

    }

    componentDidMount() {
        autoComplete($('input.autocomplete'), this.state.workers);
    }


    componentDidUpdate() {
        autoComplete($('input.autocomplete'), this.state.workers);
    }


    componentWillMount() {


        $.ajax({
            url: 'http://52.41.181.217/esg-sql-service/search/',
            type: 'POST',
            data: {query: quries.GET_ALL_CONSULERS},
            dataType: 'json',
            success: function (data) {
                let tempData = {data:{}};
                for(let workerName of data){
                    tempData.data[workerName['FullName']] = null;
                }
                this.setState({
                    workers:tempData
                });
            }.bind(this)
        });

    }


    render() {
        return (
            <div className="container">
                <div className="input-field col s12">
                    <input type="text" id="autocomplete-input" className="autocomplete"/>
                    <label htmlFor="autocomplete-input">שם העובד</label>
                </div>
            </div>
        );
    }
}


export default AutoComplete