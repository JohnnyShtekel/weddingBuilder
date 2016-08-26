import React from 'react';
import * as quries from '../utils/queries'
import Select from 'react-select';

class AutoComplete extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            workers: []
        };
        this.workers = []

    }


    componentDidUpdate() {
          $('input.autocomplete').autocomplete(this.state.workers);
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
                this.workers = data;
            }.bind(this)
        });

    }


    render() {
        return (
            <div className="row">
                <div className="col s12">
                    <div className="row">
                        <div className="input-field col s12">
                            <i className="material-icons prefix">perm_identity</i>
                            <input type="text" id="autocomplete-input" onChange={this.test} className="autocomplete"/>
                                <label htmlFor="autocomplete-input">נא לבחור עובד</label>
                        </div>
                    </div>
                </div>
            </div>
    );
    }
    }


    export default AutoComplete